"""
routers/assistente.py (TASK-191) — Assistente de IA com FERRAMENTAS e SKILLS.

POST /api/assistente/chat       conversa em NDJSON (uma linha JSON por evento: ferramenta, texto, plano, acao_ui, erro, fim). A tabela vem do navegador a cada pedido;
                                nada é alterado aqui: o PLANO volta para a tela, que mostra, pede confirmação e aplica (com desfazer).
POST /api/assistente/previa     compara as quantidades dos dois cenários montados pela Totalizadora do Resumo.
GET/PUT /api/assistente/reserva  provedor reserva automático (PUT/testar: só admin; a chave nunca volta para a tela);  POST /api/assistente/reserva/testar
GET  /api/assistente/skills     skills visíveis ao projeto;  GET/PUT/DELETE /api/assistente/skills/{escopo}/{nome}  ler/gravar/apagar (gravar e apagar: só admin)
POST /api/assistente/skills/gerar-catalogo   lê a Base de Orçamento e grava a skill `ativos` do projeto (só admin)
O atalho de ativos e a prévia funcionam sem chave de IA. As skills são arquivos locais.
"""
import asyncio
import hashlib
import json
import os
import re
import time
from contextlib import aclosing, closing
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

import database
from config import PROMPT_ASSISTENTE_PATH, logger
from middleware.auth_middleware import require_role
from routers.ai_chat import SENTINELA_CHAVE, _checar_rate_limit, resolver_provedor
from services.assistente import agente, catalogo, ferramentas, skills
from services.provedores_ia import OPENAI_BASE, modelo_openai_compativel, validar_base_openai

router = APIRouter(prefix="/api/assistente", tags=["assistente"])

MODELOS_PADRAO = {
    "gemini": ["gemini-3.1-flash-lite", "gemini-3.6-flash", "gemini-3.5-flash"],
    "openai": ["gpt-4o-mini", "gpt-4o"],
    "claude": ["claude-haiku-5-5", "claude-sonnet-5-5"],
}
_VERBOS_RETIRAR = ("remover", "remova", "rem", "tira", "tire", "retirar", "retire", "excluir", "exclua")
_VERBOS_ADICIONAR = ("adicionar", "adicione", "add", "adc", "coloca", "coloque", "inserir", "insira")
_FAST = re.compile(r"^(" + "|".join(_VERBOS_ADICIONAR + _VERBOS_RETIRAR) + r")\s+(.+)$", re.IGNORECASE)
_ITEM_FAST = re.compile(r"^(\d+(?:[.,]\d+)?)(?:\s+|-|\s*[xX]\s+)([A-Za-z][A-Za-z0-9_/().+*-]*)$")      # «2 CON», «2-CON», «2x CON»: um ativo só, sem espaços ("11300D" não é «11300 D»)


class AssistenteRequest(BaseModel):
    prompt: str
    history: List[Dict[str, Any]] = []
    provider: str = "gemini"
    openai_base_url: str = ""
    projeto_codigo: str = "229"
    cabos: List[Dict[str, Any]] = []          # as linhas das tabelas como estão na tela (cada uma com `id` = posição na tela)
    outros: List[Dict[str, Any]] = []
    obra: Optional[Dict[str, Any]] = None
    selecao_mapa: Optional[Dict[str, Any]] = None
    tabela_incompleta: bool = False


class TabelasOrcamento(BaseModel):
    cabos: List[Dict[str, Any]]
    outros: List[Dict[str, Any]]


class PreviaRequest(BaseModel):
    projeto_codigo: str
    obra: Optional[Dict[str, Any]] = None
    antes: TabelasOrcamento
    depois: TabelasOrcamento


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
    idx = skills.indice_para_prompt(ctx.slug, ctx.projeto_codigo)
    if idx:
        partes.append(idx)
    try:
        from routers.regras import get_regras
        regras = get_regras(projeto_codigo=projeto_codigo)
        if regras:
            partes.append("REGRAS APRENDIDAS (ditas pelo usuário; valem sempre):\n" + "\n".join(f"- {r['conteudo']}" for r in regras))
    except Exception as e:
        logger.warning(f"Assistente: regras aprendidas indisponíveis: {e}")
    try:                                              # o formato das linhas vem já no prompt: evita uma volta ao modelo só para «carregar a skill»
        partes.append("SKILL formato_ativos (JÁ CARREGADA; não chame carregar_skill para ela):\n" + skills.carregar(ctx.slug, "formato_ativos", ctx.projeto_codigo))
    except skills.ErroSkill:
        pass
    partes.append(f"ESTADO: projeto {ctx.projeto_nome or ctx.projeto_codigo} (código {ctx.projeto_codigo}); tela com {len(ctx.cabos)} linha(s) em Cabos e {len(ctx.outros)} em Outros.")
    if ctx.selecao_mapa:
        partes.append("Há um poste ou vão selecionado. Use consultar_vinculos para ler os IDs e diagnosticar a seleção.")
    if ctx.tabela_incompleta:
        partes.append("Cópia parcial da tabela. Indique essa limitação ao responder.")
    partes.append(_tela_resumida(ctx))
    return "\n\n".join(partes)


LIMITE_TELA_NO_PROMPT = 60


def _tela_resumida(ctx: ferramentas.Contexto) -> str:
    """As linhas da tela com os ids (`[outros-3]`) quando são poucas: o modelo não precisa chamar `ler_tabela` (uma volta a menos)."""
    linhas = []
    for t in ("cabos", "outros"):
        for r in ctx.cabos if t == "cabos" else ctx.outros:
            if isinstance(r, dict) and not r.get("congelada"):
                linhas.append(f"[{t}-{r.get('id')}] {r.get('operacao') or '-'} | {str(r.get('ativo') or '')[:160]}")
    if not linhas:
        return "TELA ATUAL: vazia."
    if len(linhas) > LIMITE_TELA_NO_PROMPT:
        return f"TELA ATUAL: {len(linhas)} linhas (muitas para listar aqui); use ler_tabela."
    return "TELA ATUAL (id entre colchetes; use o número como `id`):\n" + "\n".join(linhas)


def _fast_path(prompt: str):
    """«adicionar 2 CON» / «remover 1 TR125» / «adicionar 2 CON, 3 PR220 e 1 CFU»: sem IA (instantâneo). Devolve os itens candidatos ou None
    (qualquer item que não seja «quantidade ativo» manda o pedido inteiro para a IA)."""
    from services.obras_contexto import menciona_obra
    m = _FAST.match(prompt.strip())
    if not m or menciona_obra(prompt):
        return None
    op = "R" if m.group(1).lower() in _VERBOS_RETIRAR else "I"
    partes = [x for x in re.split(r",\s+|;\s*|\s+e\s+", m.group(2).strip()) if x.strip()]
    itens = []
    for i, parte in enumerate(partes, 1):
        mi = _ITEM_FAST.match(parte.strip())
        if not mi:
            return None
        itens.append({"n": i, "acao": "adicionar", "tabela": "outros", "id": None, "operacao": op, "ativo": f"{mi.group(1).replace(',', '.')}-{mi.group(2).upper()}", "antes": None, "avisos": []})
    return {"tipo": "plano", "resumo": "Atalho rápido", "itens": itens} if itens and len(itens) <= 20 else None


def _ndjson(eventos):
    async def gerar():
        async with aclosing(eventos):
            async for ev in eventos:
                yield json.dumps(ev, ensure_ascii=False) + "\n"
    return StreamingResponse(gerar(), media_type="application/x-ndjson", headers={"Cache-Control": "no-store"})


def _rotulo(provider: str, modelos: list) -> str:
    return f"{ {'gemini': 'Gemini', 'openai': 'OpenAI-compatível', 'claude': 'Claude'}.get(provider, provider) } · {modelos[0] if modelos else ''}".strip(" ·")


def _configuracao_reserva(conn) -> dict:
    return dict(conn.execute("SELECT chave, valor FROM configuracoes WHERE chave GLOB 'ai_reserva_*'"))


def _ler_configuracao_reserva() -> dict:
    with closing(database.get_connection()) as conn:
        return _configuracao_reserva(conn)


def _configurar_adaptador_reserva(cfg: dict) -> dict:
    provider = cfg.get("ai_reserva_provider", "openai")
    if provider not in MODELOS_PADRAO:
        raise ValueError("Provedor reserva não configurado.")
    chave = cfg.get("ai_reserva_api_key")
    if not chave or chave == SENTINELA_CHAVE:
        raise ValueError("Sem chave salva para o provedor reserva.")
    base = validar_base_openai(cfg.get("ai_reserva_base_url") or "") if provider == "openai" else ""
    modelo = modelo_openai_compativel(base, (cfg.get("ai_reserva_model") or "").strip())
    return {"provider": provider, "api_key": chave, "base_url": base,
            "modelos": [modelo] if modelo else list(MODELOS_PADRAO[provider])}


def _reserva_ativa() -> Optional[dict]:
    cfg = _ler_configuracao_reserva()
    if cfg.get("ai_reserva_ativa") == "1":
        try:
            return _configurar_adaptador_reserva(cfg)
        except ValueError:
            logger.warning("Configuração inválida do provedor reserva. Revise a URL e a chave de IA.")
    return None


def _contexto(req: AssistenteRequest, request: Request) -> ferramentas.Contexto:
    user = _user(request)
    from routers.obras import buscar_obra, listar_visiveis
    return ferramentas.Contexto(
        projeto_codigo=req.projeto_codigo, projeto_nome=_nome_do_projeto(req.projeto_codigo), user=user,
        cabos=req.cabos[:2000], outros=req.outros[:2000], obra=req.obra, selecao_mapa=req.selecao_mapa,
        tabela_incompleta=req.tabela_incompleta or len(req.cabos) > 2000 or len(req.outros) > 2000,
        carregar_obras=lambda: listar_visiveis(user, req.projeto_codigo) if user else [],
        buscar_obra=lambda oid: buscar_obra(user, oid, req.projeto_codigo) if user else None,
    )


@router.post("/previa")
def previa(req: PreviaRequest, request: Request):
    if any(len(t.cabos) + len(t.outros) > 20000 for t in (req.antes, req.depois)):
        raise HTTPException(status_code=400, detail="A prévia excede o limite. Calcule o orçamento pelo Resumo.")
    ctx = _contexto(AssistenteRequest(prompt="", projeto_codigo=req.projeto_codigo, obra=req.obra), request)
    try:
        return ferramentas.comparar_orcamento(ctx, req.antes.model_dump(), req.depois.model_dump())
    except Exception:
        logger.exception("Falha na prévia do orçamento do assistente")
        raise HTTPException(status_code=422, detail="Não foi possível calcular a prévia com a base e as regras do projeto.")


@router.post("/chat")
async def chat(req: AssistenteRequest, request: Request):
    prazo = agente.novo_prazo()
    rapido = _fast_path(req.prompt)
    if rapido:
        async def so_o_plano():
            try:
                ctx = await agente.executar_em_thread(prazo, _contexto, req, request)
                resultado = await agente.executar_em_thread(
                    prazo, ferramentas.executar, ctx, "propor_alteracoes",
                    {"alteracoes": rapido["itens"], "resumo": rapido["resumo"]})
                if ctx.plano:
                    yield {"tipo": "plano", "resumo": ctx.resumo_plano, "projeto": ctx.projeto_codigo, "itens": ctx.plano}
                else:
                    erros = "; ".join(r["motivo"] for r in resultado.get("rejeitadas", []))
                    yield {"tipo": "erro", "mensagem": resultado.get("erro") or erros}
            except Exception as e:
                yield {"tipo": "erro", "mensagem": agente.mensagem_amigavel(e)}
            yield {"tipo": "fim"}
        return _ndjson(so_o_plano())

    reserva = _reserva_ativa()
    prov = None
    try:
        prov = resolver_provedor(request, req.provider, req.openai_base_url)
    except HTTPException as e:
        if e.status_code != 401 or not reserva:
            raise
    if prov and prov["origem_chave"] == "padrao":
        _checar_rate_limit(request)

    async def conversar():
        try:
            ctx = await agente.executar_em_thread(prazo, _contexto, req, request)
            system = await agente.executar_em_thread(prazo, montar_system, ctx, req.projeto_codigo)
        except Exception as e:
            yield {"tipo": "erro", "mensagem": agente.mensagem_amigavel(e)}
            yield {"tipo": "fim"}
            return

        adaptador_reserva = agente.criar_adaptador(reserva["provider"], reserva["api_key"], reserva["modelos"], reserva["base_url"]) if reserva else None
        if prov:
            modelos = [prov["custom_model"]] if prov["custom_model"] else MODELOS_PADRAO[prov["provider"]]
            adaptador = agente.criar_adaptador(prov["provider"], prov["api_key"], modelos, prov["openai_base_url"])
            if adaptador_reserva:
                credencial = hashlib.sha256(prov["api_key"].encode()).hexdigest()
                chave = f"{prov['provider']}|{prov['openai_base_url']}|{modelos[0]}|{credencial}"
                adaptador = agente.AdaptadorComReserva(adaptador, adaptador_reserva, chave, _rotulo(reserva["provider"], reserva["modelos"]))
        else:
            adaptador = adaptador_reserva
        try:
            async with aclosing(agente.rodar(adaptador, system, agente.historico_neutro(req.history), req.prompt.strip(), ctx, prazo=prazo)) as eventos:
                async for evento in eventos:
                    yield evento
        finally:
            await adaptador.fechar()

    return _ndjson(conversar())


# ── provedor reserva ──────────────────────────────────────────────────────────────────────────────────
@router.get("/reserva")
def ler_reserva(request: Request):
    cfg = _ler_configuracao_reserva()
    provider = cfg.get("ai_reserva_provider", "openai")
    base = cfg.get("ai_reserva_base_url") or ""
    modelo = modelo_openai_compativel(base.rstrip("/"), cfg.get("ai_reserva_model") or "")
    return {"ativa": cfg.get("ai_reserva_ativa") == "1", "provider": provider if provider in MODELOS_PADRAO else "openai",
            "base_url": base, "model": modelo,
            "tem_chave": bool(cfg.get("ai_reserva_api_key")), "pode_editar": _user(request).get("role") == "admin"}


@router.put("/reserva", dependencies=[Depends(require_role("admin"))])
def gravar_reserva(req: ReservaRequest):
    provider, base, modelo = req.provider.strip().lower(), req.base_url.strip(), req.model.strip()
    if provider == "anthropic":
        provider = "claude"
    if provider not in MODELOS_PADRAO:
        raise HTTPException(status_code=400, detail="Provedor deve ser gemini, openai (compatível) ou claude.")
    if provider != "openai" and base:
        raise HTTPException(status_code=400, detail="URL base só vale para o provedor OpenAI-compatível.")
    if base:
        try:
            base = validar_base_openai(base)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e)) from e
        if base != OPENAI_BASE and not modelo:
            raise HTTPException(status_code=400, detail="Com URL base própria, informe o nome do modelo (ex.: deepseek-flash).")
    if len(modelo) > 120:
        raise HTTPException(status_code=400, detail="Nome de modelo grande demais.")

    chave_nova = (req.api_key or "").strip() if not req.limpar_chave else ""
    if chave_nova == SENTINELA_CHAVE:
        raise HTTPException(status_code=400, detail="Informe uma chave própria do provedor reserva.")
    with closing(database.get_connection()) as conn, conn:
        conn.execute("BEGIN IMMEDIATE")
        cfg = _configuracao_reserva(conn)
        mesmo_provedor = provider == cfg.get("ai_reserva_provider", "openai")
        base_anterior = (cfg.get("ai_reserva_base_url") or "").rstrip("/")
        mesma_base = provider != "openai" or (base or OPENAI_BASE) == (base_anterior or OPENAI_BASE)
        chave_final = chave_nova
        if not chave_final and not req.limpar_chave and mesmo_provedor and mesma_base:
            chave_final = cfg.get("ai_reserva_api_key") or ""
        if req.ativa and not chave_final:
            raise HTTPException(status_code=400, detail="Informe a chave do provedor reserva para ativá-lo. Ao trocar de provedor ou URL, informe a chave correspondente.")
        pares = {"ai_reserva_ativa": "1" if req.ativa else "0", "ai_reserva_provider": provider,
                 "ai_reserva_base_url": base, "ai_reserva_model": modelo_openai_compativel(base, modelo),
                 "ai_reserva_api_key": chave_final}
        conn.executemany("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES (?, ?)", pares.items())
    agente.limpar_pausas()
    return {"ok": True, "ativa": req.ativa, "tem_chave": bool(chave_final)}


@router.post("/reserva/testar", dependencies=[Depends(require_role("admin"))])
async def testar_reserva():
    """Valida a chamada, a execução da ferramenta e a resposta final do modelo."""
    try:
        cfg = _configurar_adaptador_reserva(_ler_configuracao_reserva())
    except ValueError as e:
        return {"ok": False, "erro": str(e)}
    ad = agente.criar_adaptador(cfg["provider"], cfg["api_key"], cfg["modelos"], cfg["base_url"])
    ad.esperas = (0,)
    t0 = time.monotonic()
    spec = next(s for s in ferramentas.especificacoes() if s["name"] == "ler_tabela")
    system = "Use a ferramenta ler_tabela com tabela=cabos, e responda após receber o resultado."
    mensagens = [{"role": "user", "texto": "Consulte a tabela de Cabos pela ferramenta. Depois informe se há linhas."}]

    async def verificar():
        t = await ad.turno(system, mensagens, [spec])
        if len(t.chamadas) != 1 or t.chamadas[0].nome != "ler_tabela":
            raise ValueError("O modelo não chamou apenas a ferramenta de leitura esperada.")
        chamada = t.chamadas[0]
        resultado = ferramentas.executar(ferramentas.Contexto(), chamada.nome, chamada.args)
        if "cabos" not in resultado:
            raise ValueError("O modelo enviou argumentos inválidos para ler_tabela.")
        mensagens.append(agente.mensagem_do_turno(t))
        mensagens.append({"role": "tool", "resultados": [{"id": chamada.id, "nome": chamada.nome, "resultado": resultado}]})
        final = await ad.turno(system, mensagens, [spec])
        if final.chamadas or not final.texto.strip():
            raise ValueError("O modelo não retornou uma resposta final após receber o resultado da ferramenta.")

    try:
        await asyncio.wait_for(verificar(), timeout=40)
    except asyncio.TimeoutError:
        return {"ok": False, "erro": "O teste completo do provedor reserva excedeu 40 segundos."}
    except Exception as e:
        return {"ok": False, "erro": agente.mensagem_amigavel(e)}
    finally:
        await ad.fechar()
    return {"ok": True, "modelo": cfg["modelos"][0], "ms": int((time.monotonic() - t0) * 1000),
            "resposta": "Consulta de teste e resposta final concluídas."}


# ── skills ────────────────────────────────────────────────────────────────────────────────────────────
def _escopo(escopo: str) -> str:
    """Normaliza o alias do escopo geral."""
    return skills.GERAL if escopo in ("_geral", "geral") else escopo


@router.get("/skills")
def listar_skills(request: Request, projeto: str = ""):
    nome = _nome_do_projeto(projeto) if projeto else ""
    slug = skills.slug(nome)
    return {"projeto": nome, "escopo_projeto": skills.escopo_do_projeto(projeto) if nome else "", "skills": skills.listar(slug, projeto if nome else ""), "pode_editar": _user(request).get("role") == "admin"}


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
        r = skills.salvar(skills.escopo_do_projeto(projeto), "ativos", texto)
    except skills.ErroSkill as e:
        raise HTTPException(status_code=400, detail=str(e))
    return {"ok": True, "escopo": r["escopo"], "nome": "ativos", "ativos": len(indice), "chars": r["chars"]}
