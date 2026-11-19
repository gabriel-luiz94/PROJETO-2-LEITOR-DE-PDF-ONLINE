"""
routers/assistente.py (TASK-191) — Assistente de IA com FERRAMENTAS e SKILLS.

POST /api/assistente/chat       conversa em NDJSON (uma linha JSON por evento: ferramenta, texto, plano, acao_ui, erro, fim). A tabela vem do navegador a cada pedido;
                                nada é alterado aqui: o PLANO volta para a tela, que mostra, pede confirmação e aplica (com desfazer).
GET  /api/assistente/skills     skills visíveis ao projeto;  GET/PUT/DELETE /api/assistente/skills/{escopo}/{nome}  ler/gravar/apagar (gravar e apagar: só admin)
POST /api/assistente/skills/gerar-catalogo   lê a Base de Orçamento e grava a skill `ativos` do projeto (só admin)
Sem chave de IA não há chat; as skills (arquivos) funcionam sempre.
"""
import json
import os
import re
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

import database
from config import PROMPT_ASSISTENTE_PATH, logger
from middleware.auth_middleware import require_role
from routers.ai_chat import _checar_rate_limit, resolver_provedor
from services.assistente import agente, catalogo, ferramentas, skills

router = APIRouter(prefix="/api/assistente", tags=["assistente"])

MODELOS_PADRAO = {
    "gemini": ["gemini-3.1-flash-lite", "gemini-3.6-flash", "gemini-3.5-flash"],
    "openai": ["gpt-4o-mini", "gpt-4o"],
    "claude": ["claude-haiku-5-5", "claude-sonnet-5-5"],
}
_FAST = re.compile(r"^(adicionar|add|adc|coloca|inserir|remover|rem|tira|excluir)\s+([\d.,]+)\s+(.+)$", re.IGNORECASE)


class AssistenteRequest(BaseModel):
    prompt: str
    history: List[Dict[str, Any]] = []
    provider: str = "gemini"
    openai_base_url: str = ""
    projeto_codigo: str = "229"
    cabos: List[Dict[str, Any]] = []          # as linhas das tabelas como estão na tela (cada uma com `id` = posição na tela)
    outros: List[Dict[str, Any]] = []
    obra: Optional[Dict[str, Any]] = None


class SkillRequest(BaseModel):
    texto: str


def _user(request: Request) -> dict:
    return getattr(request.state, "user", None) or {}


def _nome_do_projeto(codigo: str) -> str:
    conn = database.get_row_connection()
    try:
        r = conn.execute("SELECT nome FROM projetos WHERE codigo = ?", (codigo,)).fetchone()
    finally:
        conn.close()
    return r["nome"] if r else ""


def montar_system(ctx: ferramentas.Contexto, projeto_codigo: str) -> str:
    """Prompt de sistema: papel e regras (arquivo) + índice das skills + regras aprendidas do projeto + o estado resumido da tela."""
    base = "Você é o assistente do OrçaCore, especialista em redes elétricas. Use as ferramentas; não invente códigos."
    if os.path.exists(PROMPT_ASSISTENTE_PATH):
        with open(PROMPT_ASSISTENTE_PATH, "r", encoding="utf-8") as f:
            base = f.read().strip()
    partes = [base]
    idx = skills.indice_para_prompt(ctx.slug)
    if idx:
        partes.append(idx)
    try:
        from routers.regras import get_regras
        regras = get_regras(projeto_codigo=projeto_codigo)
        if regras:
            partes.append("REGRAS APRENDIDAS (ditas pelo usuário; valem sempre):\n" + "\n".join(f"- {r['conteudo']}" for r in regras))
    except Exception as e:
        logger.warning(f"Assistente: regras aprendidas indisponíveis: {e}")
    partes.append(f"ESTADO: projeto {ctx.projeto_nome or ctx.projeto_codigo} (código {ctx.projeto_codigo}); tela com {len(ctx.cabos)} linha(s) em Cabos e {len(ctx.outros)} em Outros.")
    return "\n\n".join(partes)


def _fast_path(prompt: str):
    """«adicionar 2 CON» / «remover 1 TR125»: sem IA. Devolve o evento de plano (auto-aplicável) ou None."""
    from services.obras_contexto import menciona_obra
    m = _FAST.match(prompt.strip())
    if not m or menciona_obra(prompt):
        return None
    op = "R" if m.group(1).lower().startswith("rem") or m.group(1).lower() in ("tira", "excluir") else "I"
    item = {"n": 1, "acao": "adicionar", "tabela": "outros", "id": None, "operacao": op, "ativo": f"{m.group(2).replace(',', '.')}-{m.group(3).upper().strip()}", "antes": None, "avisos": []}
    return {"tipo": "plano", "resumo": "Atalho rápido", "auto": True, "itens": [item]}


def _ndjson(eventos):
    async def gerar():
        async for ev in eventos:
            yield json.dumps(ev, ensure_ascii=False) + "\n"
    return StreamingResponse(gerar(), media_type="application/x-ndjson", headers={"Cache-Control": "no-store"})


@router.post("/chat")
async def chat(req: AssistenteRequest, request: Request):
    prov = resolver_provedor(request, req.provider, req.openai_base_url)          # 401 sem chave
    if prov["origem_chave"] == "padrao":
        _checar_rate_limit(request)
    rapido = _fast_path(req.prompt)
    if rapido:
        async def so_o_plano():
            yield rapido
            yield {"tipo": "fim"}
        return _ndjson(so_o_plano())

    user = _user(request)
    from routers.obras import buscar_obra, listar_visiveis
    ctx = ferramentas.Contexto(
        projeto_codigo=req.projeto_codigo, projeto_nome=_nome_do_projeto(req.projeto_codigo), user=user, cabos=req.cabos[:2000], outros=req.outros[:2000], obra=req.obra,
        carregar_obras=lambda: listar_visiveis(user, req.projeto_codigo) if user else [],
        buscar_obra=lambda oid: buscar_obra(user, oid, req.projeto_codigo) if user else None)
    modelos = [prov["custom_model"]] if prov["custom_model"] else MODELOS_PADRAO[prov["provider"]]
    adaptador = agente.criar_adaptador(prov["provider"], prov["api_key"], modelos, prov["openai_base_url"])
    return _ndjson(agente.rodar(adaptador, montar_system(ctx, req.projeto_codigo), agente.historico_neutro(req.history), req.prompt.strip(), ctx))


# ── skills ────────────────────────────────────────────────────────────────────────────────────────────
def _escopo(escopo: str, codigo: str = "") -> str:
    """"_geral" ou o slug do projeto (um slug só vale se for o do projeto pedido ou um nome bem formado)."""
    return skills.GERAL if escopo in ("_geral", "geral") else escopo


@router.get("/skills")
def listar_skills(request: Request, projeto: str = ""):
    nome = _nome_do_projeto(projeto) if projeto else ""
    slug = skills.slug(nome)
    return {"projeto": nome, "escopo_projeto": slug, "skills": skills.listar(slug), "pode_editar": _user(request).get("role") == "admin"}


@router.get("/skills/{escopo}/{nome}")
def ler_skill(escopo: str, nome: str):
    try:
        return {"escopo": _escopo(escopo), "nome": nome, "texto": skills.texto_bruto(_escopo(escopo), nome)}
    except skills.ErroSkill as e:
        raise HTTPException(status_code=404, detail=str(e))


@router.put("/skills/{escopo}/{nome}", dependencies=[Depends(require_role("admin"))])
def gravar_skill(escopo: str, nome: str, req: SkillRequest):
    try:
        return skills.salvar(_escopo(escopo), nome, req.texto)
    except skills.ErroSkill as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.delete("/skills/{escopo}/{nome}", dependencies=[Depends(require_role("admin"))])
def apagar_skill(escopo: str, nome: str):
    try:
        if not skills.apagar(_escopo(escopo), nome):
            raise HTTPException(status_code=404, detail="Não há skill sua com esse nome (as que vêm com o programa não se apagam; salve uma própria para sobrepor).")
    except skills.ErroSkill as e:
        raise HTTPException(status_code=400, detail=str(e))
    return {"ok": True}


@router.post("/skills/gerar-catalogo", dependencies=[Depends(require_role("admin"))])
def gerar_catalogo(request: Request, projeto: str):
    """Lê a Base de Orçamento do usuário/mestre e grava a skill `ativos` do projeto."""
    nome = _nome_do_projeto(projeto)
    if not nome:
        raise HTTPException(status_code=404, detail="Projeto não cadastrado.")
    from services.sync_service import get_merged_orcamento
    indice = catalogo.indexar(get_merged_orcamento(_user(request).get("user_id")), nome, projeto)
    if not indice:
        raise HTTPException(status_code=400, detail="A Base de Orçamento não tem ativos para este projeto.")
    texto = catalogo.gerar_markdown(indice, nome)
    try:
        r = skills.salvar(skills.slug(nome), "ativos", texto)
    except skills.ErroSkill as e:
        raise HTTPException(status_code=400, detail=str(e))
    return {"ok": True, "escopo": r["escopo"], "nome": "ativos", "ativos": len(indice), "chars": r["chars"]}
