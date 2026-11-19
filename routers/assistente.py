"""
routers/assistente.py (TASK-191) — Assistente de IA com FERRAMENTAS e SKILLS.

POST /api/assistente/chat       conversa em NDJSON (uma linha JSON por evento: ferramenta, texto, plano, acao_ui, erro, fim). A tabela vem do navegador a cada pedido;
                                nada é alterado aqui: o PLANO volta para a tela, que mostra, pede confirmação e aplica (com desfazer).
GET/PUT /api/assistente/reserva  provedor reserva automático (PUT/testar: só admin; a chave nunca volta para a tela);  POST /api/assistente/reserva/testar
GET  /api/assistente/skills     skills visíveis ao projeto;  GET/PUT/DELETE /api/assistente/skills/{escopo}/{nome}  ler/gravar/apagar (gravar e apagar: só admin)
POST /api/assistente/skills/gerar-catalogo   lê a Base de Orçamento e grava a skill `ativos` do projeto (só admin)
Sem chave de IA não há chat; as skills (arquivos) funcionam sempre.
"""
import asyncio
import json
import os
import re
import time
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

import database
from config import PROMPT_ASSISTENTE_PATH, logger
from middleware.auth_middleware import require_role
from routers.ai_chat import _checar_rate_limit, _ler_configuracao, resolver_credencial, resolver_provedor
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


class ReservaRequest(BaseModel):
    ativa: bool = False
    provider: str = "openai"
    base_url: str = ""
    model: str = ""
    api_key: Optional[str] = None            # None/vazio = mantém a chave já salva
    limpar_chave: bool = False


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


def _rotulo(provider: str, modelos: list) -> str:
    return f"{ {'gemini': 'Gemini', 'openai': 'OpenAI-compatível', 'claude': 'Claude'}.get(provider, provider) } · {modelos[0] if modelos else ''}".strip(" ·")


def _reserva_ativa() -> Optional[dict]:
    """A configuração do provedor reserva pronta para uso ({provider, api_key, base_url, modelos}) ou None (desligado, sem chave ou inválido)."""
    if _ler_configuracao("ai_reserva_ativa") != "1":
        return None
    provider = _ler_configuracao("ai_reserva_provider", "openai")
    if provider not in MODELOS_PADRAO:
        return None
    chave, _origem = resolver_credencial(None, _ler_configuracao("ai_reserva_api_key"), provider)
    if not chave:
        return None
    modelo = (_ler_configuracao("ai_reserva_model") or "").strip()
    return {"provider": provider, "api_key": chave, "base_url": _ler_configuracao("ai_reserva_base_url", "") or "", "modelos": [modelo] if modelo else list(MODELOS_PADRAO[provider])}


@router.post("/chat")
async def chat(req: AssistenteRequest, request: Request):
    rapido = _fast_path(req.prompt)                       # sem IA: vale mesmo sem chave nenhuma
    if rapido:
        async def so_o_plano():
            yield rapido
            yield {"tipo": "fim"}
        return _ndjson(so_o_plano())

    reserva = _reserva_ativa()
    prov = None
    try:
        prov = resolver_provedor(request, req.provider, req.openai_base_url)          # 401 sem chave
    except HTTPException as e:
        if e.status_code != 401 or not reserva:
            raise                                          # sem chave no principal mas com reserva: o reserva atende
    if prov and prov["origem_chave"] == "padrao":
        _checar_rate_limit(request)

    user = _user(request)
    from routers.obras import buscar_obra, listar_visiveis
    ctx = ferramentas.Contexto(
        projeto_codigo=req.projeto_codigo, projeto_nome=_nome_do_projeto(req.projeto_codigo), user=user, cabos=req.cabos[:2000], outros=req.outros[:2000], obra=req.obra,
        carregar_obras=lambda: listar_visiveis(user, req.projeto_codigo) if user else [],
        buscar_obra=lambda oid: buscar_obra(user, oid, req.projeto_codigo) if user else None)
    adaptador_reserva = agente.criar_adaptador(reserva["provider"], reserva["api_key"], reserva["modelos"], reserva["base_url"]) if reserva else None
    if prov:
        modelos = [prov["custom_model"]] if prov["custom_model"] else MODELOS_PADRAO[prov["provider"]]
        adaptador = agente.criar_adaptador(prov["provider"], prov["api_key"], modelos, prov["openai_base_url"])
        if adaptador_reserva:
            chave = f"{prov['provider']}|{prov['openai_base_url']}|{modelos[0]}"
            adaptador = agente.AdaptadorComReserva(adaptador, adaptador_reserva, chave, _rotulo(reserva["provider"], reserva["modelos"]))
    else:
        adaptador = adaptador_reserva
    return _ndjson(agente.rodar(adaptador, montar_system(ctx, req.projeto_codigo), agente.historico_neutro(req.history), req.prompt.strip(), ctx))


# ── provedor reserva ──────────────────────────────────────────────────────────────────────────────────
def _gravar_config(pares: dict):
    conn = database.get_connection()
    try:
        for k, v in pares.items():
            conn.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES (?, ?)", (k, v))
        conn.commit()
    finally:
        conn.close()


@router.get("/reserva")
def ler_reserva(request: Request):
    provider = _ler_configuracao("ai_reserva_provider", "openai")
    return {"ativa": _ler_configuracao("ai_reserva_ativa") == "1", "provider": provider if provider in MODELOS_PADRAO else "openai",
            "base_url": _ler_configuracao("ai_reserva_base_url", "") or "", "model": _ler_configuracao("ai_reserva_model", "") or "",
            "tem_chave": bool(_ler_configuracao("ai_reserva_api_key")), "pode_editar": _user(request).get("role") == "admin"}


@router.put("/reserva", dependencies=[Depends(require_role("admin"))])
def gravar_reserva(req: ReservaRequest):
    provider, base, modelo = req.provider.strip().lower(), req.base_url.strip(), req.model.strip()
    if provider == "anthropic":
        provider = "claude"
    if provider not in MODELOS_PADRAO:
        raise HTTPException(status_code=400, detail="Provedor deve ser gemini, openai (compatível) ou claude.")
    if base and (not re.match(r"^https?://\S+$", base) or len(base) > 300):
        raise HTTPException(status_code=400, detail="A URL base deve começar com http:// ou https://.")
    if provider != "openai" and base:
        raise HTTPException(status_code=400, detail="URL base só vale para o provedor OpenAI-compatível.")
    if provider == "openai" and base and not modelo:
        raise HTTPException(status_code=400, detail="Com URL base própria, informe o nome do modelo (ex.: deepseek-chat).")
    if len(modelo) > 120:
        raise HTTPException(status_code=400, detail="Nome de modelo grande demais.")
    chave_nova = (req.api_key or "").strip()
    if req.limpar_chave:
        chave_nova = ""
    chave_final = chave_nova or ("" if req.limpar_chave else (_ler_configuracao("ai_reserva_api_key") or ""))
    if req.ativa and not resolver_credencial(None, chave_final, provider)[0]:
        raise HTTPException(status_code=400, detail="Informe a chave do provedor reserva para ativá-lo.")
    pares = {"ai_reserva_ativa": "1" if req.ativa else "0", "ai_reserva_provider": provider, "ai_reserva_base_url": base, "ai_reserva_model": modelo}
    if chave_nova or req.limpar_chave:
        pares["ai_reserva_api_key"] = chave_nova
    _gravar_config(pares)
    agente.limpar_pausas()
    return {"ok": True, "ativa": req.ativa, "tem_chave": bool(chave_final)}


@router.post("/reserva/testar", dependencies=[Depends(require_role("admin"))])
async def testar_reserva():
    """Uma chamada mínima (sem ferramentas) ao reserva SALVO: diz se a chave, a URL e o modelo funcionam."""
    provider = _ler_configuracao("ai_reserva_provider", "openai")
    if provider not in MODELOS_PADRAO:
        return {"ok": False, "erro": "Provedor reserva não configurado."}
    chave, _o = resolver_credencial(None, _ler_configuracao("ai_reserva_api_key"), provider)
    if not chave:
        return {"ok": False, "erro": "Sem chave salva para o provedor reserva."}
    modelo = (_ler_configuracao("ai_reserva_model") or "").strip()
    modelos = [modelo] if modelo else list(MODELOS_PADRAO[provider])
    ad = agente.criar_adaptador(provider, chave, modelos, _ler_configuracao("ai_reserva_base_url", "") or "")
    ad.esperas = (0,)
    t0 = time.monotonic()
    try:
        t = await asyncio.wait_for(ad.turno("Responda apenas: ok", [{"role": "user", "texto": "Teste de conexão."}], []), timeout=40)
    except Exception as e:
        return {"ok": False, "erro": agente.mensagem_amigavel(e)}
    return {"ok": True, "modelo": modelos[0], "ms": int((time.monotonic() - t0) * 1000), "resposta": (t.texto or "")[:80]}


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
