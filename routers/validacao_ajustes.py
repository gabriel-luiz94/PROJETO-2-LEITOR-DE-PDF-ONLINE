"""
routers/validacao_ajustes.py — Ajustes das planilhas: pré-visualização (TASK-022) e cadastro de receitas (TASK-023, ADR-006).

• `/preview` e `/descrever` só simulam: devolvem o diff; quem aplica é o frontend, depois do aceite do usuário.
• O cadastro de receitas (ações recorrentes) é em camadas, como as regras de domínio (services/ajustes_camadas.py): o
  DEFAULT guarda o container completo; cada projeto guarda só o overlay. Nuvem primeiro com fallback local, histórico com
  reversão; escrita só admin, leitura para qualquer autenticado. A gravação só acontece quando o admin clica em Salvar.
• TASK-044: terceira camada opcional, o "contexto" — um nome cadastrado dentro de um projeto (services/contextos.py),
  com seu próprio overlay de ajustes (services/repo_json.py::RepoJsonContexto). Resolução em cadeia: efetivo(padrão,
  overlay_projeto) → efetivo(<isso>, overlay_contexto), reaproveitando a MESMA função `efetivo()` sem alterá-la. Toda
  rota aceita `contexto` opcional; sem ele, comportamento idêntico ao de antes desta tarefa.
Os grupos de ativos são os efetivos do projeto (os mesmos das regras de domínio). Não registrar o conteúdo das planilhas em log.
"""
import json
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, Depends, HTTPException, Request
from pydantic import BaseModel

from config import AJUSTES_SEED_PATH
from middleware.auth_middleware import require_role
from routers.validacao_regras import DEFAULT, regras_efetivas
from services import contextos as contextos_service
from services.ajustes_camadas import (CAMPOS_META, efetivo, lista_de_bruto, normalizar_overlay, overlay_de_bruto,
                                      overlay_de_efetivo, overlay_vazio, overlay_vazio_de)
from services.ajustes_planilhas import ajustar, descrever_acao, validar_acoes, validar_receitas
from services.repo_json import ErroSincronizacao, RepoJson, RepoJsonContexto

router = APIRouter(prefix="/api/validacao/ajustes", tags=["validacao-ajustes"])
repo = RepoJson("ajustes_planilhas", "ajustes_planilhas_historico", "ajustes_json")
repo_contexto = RepoJsonContexto("ajustes_contextos", "ajustes_contextos_historico", "ajustes_json")
TAMANHO_MAX_CONTEXTO = 60


class AjustesPreviewRequest(BaseModel):
    acoes: List[Dict[str, Any]] = []
    receitas: Optional[List[str]] = None      # ids de receitas cadastradas, rodadas em ordem (antes de `acoes`)
    cabos: List[Dict[str, Any]] = []
    outros: List[Dict[str, Any]] = []
    projeto_codigo: Optional[str] = None
    contexto: Optional[str] = None


class AjustesLoteRequest(BaseModel):
    cabos: List[Dict[str, Any]] = []
    outros: List[Dict[str, Any]] = []
    projeto_codigo: Optional[str] = None
    contexto: Optional[str] = None


class DescreverAcoesRequest(BaseModel):
    acoes: List[Dict[str, Any]]
    projeto_codigo: Optional[str] = None


class SalvarAjustesPayload(BaseModel):
    projeto_codigo: str = DEFAULT
    ajustes: List[Dict[str, Any]]
    contexto: Optional[str] = None


class ProjetoPayload(BaseModel):
    projeto_codigo: str = DEFAULT


class ReverterPayload(BaseModel):
    projeto_codigo: str = DEFAULT
    historico_id: int


class ContextoPayload(BaseModel):
    projeto_codigo: str
    contexto: str


def _email(request: Request):
    user = getattr(request.state, "user", None)
    return user.get("email") if user else None


def _grupos(projeto_codigo):
    return regras_efetivas(projeto_codigo or DEFAULT)[1]


def _limpa(r: dict) -> dict:
    return {k: v for k, v in r.items() if k not in CAMPOS_META}


def _nome_contexto_valido(contexto: str) -> str:
    nome = (contexto or "").strip()
    if not nome:
        raise HTTPException(status_code=400, detail="Nome do contexto não pode ser vazio.")
    if len(nome) > TAMANHO_MAX_CONTEXTO:
        raise HTTPException(status_code=400, detail=f"Nome do contexto muito longo (máx. {TAMANHO_MAX_CONTEXTO} caracteres).")
    return nome


# ═══════════════════════════════════════════════════════════════════════════
# CADASTRO EM CAMADAS
# ═══════════════════════════════════════════════════════════════════════════
def _padrao() -> list:
    return lista_de_bruto(repo.buscar(DEFAULT))


def resolver(projeto_codigo: str = DEFAULT, contexto: Optional[str] = None):
    """(ajustes efetivos, avisos, overlay_projeto|None, overlay_contexto|None).
    `overlay_projeto` None = projeto sem versão própria (ou o próprio DEFAULT).
    `overlay_contexto` None = nenhum contexto pedido, ou contexto sem overlay próprio ainda.
    TASK-044: com `contexto`, encadeia uma 2ª chamada a `efetivo()` sobre o resultado da 1ª —
    mesma função do overlay de projeto, sem alteração nenhuma nela."""
    padrao = _padrao()
    overlay_projeto = None
    if (projeto_codigo or DEFAULT) != DEFAULT:
        bruto = repo.buscar(projeto_codigo)
        overlay_projeto = None if bruto is None else overlay_de_bruto(padrao, bruto)
    ajustes_projeto, avisos = efetivo(padrao, overlay_projeto or overlay_vazio())
    if not contexto:
        return ajustes_projeto, avisos, overlay_projeto, None
    bruto_contexto = repo_contexto.buscar(projeto_codigo or DEFAULT, contexto)
    overlay_contexto = None if bruto_contexto is None else normalizar_overlay(bruto_contexto)
    ajustes_final, avisos2 = efetivo(ajustes_projeto, overlay_contexto or overlay_vazio())
    return ajustes_final, avisos + avisos2, overlay_projeto, overlay_contexto


def ajustes_efetivos(projeto_codigo: str = DEFAULT, contexto: Optional[str] = None) -> list:
    return resolver(projeto_codigo, contexto)[0]


def _frases(ajustes: list) -> list:
    saida = []
    for r in ajustes:
        try:
            frases = [descrever_acao(a) for a in r.get("acoes") or []]
        except Exception:  # ajuste malformado gravado por outra versão: não derruba a listagem
            frases = []
        saida.append({**r, "frases": frases})
    return saida


def _por_regra(ajustes: list) -> dict:
    mapa = {}
    for r in ajustes:
        if r.get("ativa") and not r.get("oculta"):
            for rid in r.get("regras") or []:
                mapa.setdefault(rid, []).append(r["id"])
    return mapa


def _gravar(projeto: str, objeto, email):
    try:
        repo.gravar(projeto, objeto, email)
    except ErroSincronizacao as e:
        raise HTTPException(status_code=500, detail=f"Ajustes salvos localmente, mas falharam ao sincronizar com a nuvem (Supabase): {e}")
    return {"status": "success"}


def _erros(ajustes: list, grupos: dict):
    erros = validar_receitas(ajustes, grupos)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})


def _salvar(projeto: str, ajustes: list, email, contexto: Optional[str] = None):
    """DEFAULT: container completo. Projeto: `ajustes` é o efetivo editado; grava só a diferença
    para o padrão. TASK-044: com `contexto`, `ajustes` é o efetivo editado NO CONTEXTO (padrão +
    overlay do projeto + o que o admin mudou) — grava só a diferença para o efetivo do projeto."""
    if contexto:
        ajustes_projeto = resolver(projeto)[0]
        return _salvar_overlay_contexto(projeto, contexto, overlay_de_efetivo(ajustes_projeto, ajustes), email)
    grupos = _grupos(projeto)
    _erros(ajustes, grupos)
    if projeto == DEFAULT:
        return _gravar(DEFAULT, {"versao": 1, "ajustes": [_limpa(r) for r in ajustes]}, email)
    return _salvar_overlay(projeto, overlay_de_efetivo(_padrao(), ajustes), email)


def _salvar_overlay(projeto: str, overlay: dict, email):
    efetivos, _ = efetivo(_padrao(), overlay)
    _erros(efetivos, _grupos(projeto))
    return _gravar(projeto, overlay, email)


def _gravar_contexto(projeto: str, contexto: str, objeto, email):
    try:
        repo_contexto.gravar(projeto, contexto, objeto, email)
    except ErroSincronizacao as e:
        raise HTTPException(
            status_code=500,
            detail=f"Ajustes do contexto salvos localmente, mas falharam ao sincronizar com a nuvem (Supabase): {e}"
        )
    return {"status": "success"}


def _salvar_overlay_contexto(projeto: str, contexto: str, overlay: dict, email):
    ajustes_projeto = resolver(projeto)[0]
    efetivos, _ = efetivo(ajustes_projeto, overlay)
    _erros(efetivos, _grupos(projeto))
    return _gravar_contexto(projeto, contexto, overlay, email)


@router.get("")
def obter(projeto_codigo: str = DEFAULT, contexto: Optional[str] = None):
    ajustes, avisos, overlay_projeto, overlay_contexto = resolver(projeto_codigo, contexto)
    if contexto:
        personalizado = overlay_contexto is not None and not overlay_vazio_de(overlay_contexto)
    else:
        personalizado = overlay_projeto is not None and not overlay_vazio_de(overlay_projeto)
    return {"ajustes": _frases(ajustes), "avisos": avisos, "por_regra": _por_regra(ajustes),
            "versao_de": projeto_codigo if overlay_projeto is not None else DEFAULT,
            "personalizado": personalizado, "contexto": contexto}


@router.post("", dependencies=[Depends(require_role("admin"))])
def salvar(payload: SalvarAjustesPayload, request: Request):
    return _salvar(payload.projeto_codigo, payload.ajustes, _email(request), payload.contexto)


# ═══════════════════════════════════════════════════════════════════════════
# CONTEXTOS (TASK-044)
# ═══════════════════════════════════════════════════════════════════════════
@router.get("/contextos")
def listar_contextos(projeto_codigo: str = DEFAULT):
    return {"contextos": contextos_service.listar(projeto_codigo)}


@router.post("/contextos", dependencies=[Depends(require_role("admin"))])
def criar_contexto(payload: ContextoPayload, request: Request):
    nome = _nome_contexto_valido(payload.contexto)
    contextos_service.criar(payload.projeto_codigo, nome, _email(request))
    return {"status": "success"}


@router.post("/contextos/excluir", dependencies=[Depends(require_role("admin"))])
def excluir_contexto(payload: ContextoPayload):
    contextos_service.excluir(payload.projeto_codigo, payload.contexto)
    repo_contexto.excluir(payload.projeto_codigo, payload.contexto)
    return {"status": "success"}


@router.get("/historico", dependencies=[Depends(require_role("admin"))])
def historico(projeto_codigo: str = DEFAULT):
    padrao = _padrao()
    itens = []
    for h in repo.historico(projeto_codigo):
        if projeto_codigo == DEFAULT:
            ajustes = lista_de_bruto(h["bruto"])
        else:
            ajustes, _ = efetivo(padrao, overlay_de_bruto(padrao, h["bruto"]))
        itens.append({"id": h["id"], "criado_em": h["criado_em"], "criado_por": h["criado_por"], "ajustes": ajustes})
    return {"historico": itens}


@router.post("/reverter", dependencies=[Depends(require_role("admin"))])
def reverter(payload: ReverterPayload, request: Request):
    versao = next((h for h in repo.historico(payload.projeto_codigo) if h["id"] == payload.historico_id), None)
    if versao is None:
        raise HTTPException(status_code=404, detail="Versão de histórico não encontrada para este projeto.")
    if payload.projeto_codigo == DEFAULT:
        return _salvar(DEFAULT, lista_de_bruto(versao["bruto"]), _email(request))
    return _salvar_overlay(payload.projeto_codigo, overlay_de_bruto(_padrao(), versao["bruto"]), _email(request))


@router.post("/restaurar-semente", dependencies=[Depends(require_role("admin"))])
def restaurar_semente(payload: ProjetoPayload, request: Request):
    """DEFAULT: volta ao conteúdo da semente. Projeto: descarta os ajustes do projeto (volta a ser o padrão puro)."""
    if payload.projeto_codigo != DEFAULT:
        return _gravar(payload.projeto_codigo, overlay_vazio(), _email(request))
    try:
        with open(AJUSTES_SEED_PATH, "r", encoding="utf-8") as f:
            semente = lista_de_bruto(json.load(f))
    except (FileNotFoundError, json.JSONDecodeError):
        raise HTTPException(status_code=404, detail="Semente dos ajustes não encontrada.")
    return _salvar(DEFAULT, semente, _email(request))


# ═══════════════════════════════════════════════════════════════════════════
# PRÉ-VISUALIZAÇÃO (só simula)
# ═══════════════════════════════════════════════════════════════════════════
@router.post("/preview")
def preview(req: AjustesPreviewRequest):
    grupos = _grupos(req.projeto_codigo)
    acoes = []
    if req.receitas:
        por_id = {r["id"]: r for r in ajustes_efetivos(req.projeto_codigo or DEFAULT, req.contexto)}
        for rid in req.receitas:
            r = por_id.get(rid)
            if r is None:
                raise HTTPException(status_code=404, detail=f"Ajuste '{rid}' não encontrado.")
            if r.get("oculta"):
                raise HTTPException(status_code=400, detail=f"O ajuste '{rid}' está oculto neste projeto.")
            acoes += r.get("acoes") or []
    acoes += req.acoes
    erros = validar_acoes(acoes, grupos)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})
    try:
        return ajustar(acoes, req.cabos, req.outros, grupos)
    except ValueError as e:
        raise HTTPException(status_code=400, detail={"erros": e.args[0]})


LIMITE_LOTE = 30                       # ajustes habilitados calculados por pedido
ACOES_DESTRUTIVAS = {"excluir_linhas", "remover_ativo"}


def _destrutivo(acoes: list, diff: dict) -> bool:
    return any(a.get("acao") in ACOES_DESTRUTIVAS for a in acoes) or diff["resumo"]["excluir"] > 0


@router.post("/preview-lote")
def preview_lote(req: AjustesLoteRequest):
    """TASK-029: roda TODOS os ajustes habilitados do projeto (ligados e não ocultos) nas tabelas dadas e devolve, numa só
    chamada, (a) o diff de cada ajuste calculado sozinho sobre o estado atual, (b) os que não mudariam nada e (c) a CADEIA —
    todos em sequência, na ordem do cadastro, cada um enxergando o resultado do anterior — usada por "Executar todas".
    Só simula (nada é aplicado aqui); a camada 1 continua sendo a guarda de cada linha."""
    grupos = _grupos(req.projeto_codigo)
    habilitados = [a for a in ajustes_efetivos(req.projeto_codigo or DEFAULT, req.contexto)
                   if a.get("ativa") and not a.get("oculta")]
    avisos = []
    if len(habilitados) > LIMITE_LOTE:
        habilitados = habilitados[:LIMITE_LOTE]
        avisos.append(f"Só os {LIMITE_LOTE} primeiros ajustes habilitados foram calculados.")
    itens, sem_mudanca, problemas, validos = [], [], [], []
    for r in habilitados:
        acoes = r.get("acoes") or []
        try:
            diff = ajustar(acoes, req.cabos, req.outros, grupos)
        except ValueError as e:
            problemas.append({"id": r["id"], "nome": r.get("nome", r["id"]), "erros": e.args[0]})
            continue
        validos.append(r)
        base = {"id": r["id"], "nome": r.get("nome", r["id"]), "descricao": r.get("descricao", ""),
                "frases": [descrever_acao(a) for a in acoes]}
        if sum(diff["resumo"].values()):
            itens.append({**base, "destrutivo": _destrutivo(acoes, diff), "diff": diff})
        else:
            sem_mudanca.append({**base, "descartadas": diff["descartadas"]})
    cadeia, cadeia_erro = None, None
    todas = [a for r in validos for a in (r.get("acoes") or [])]
    if len(todas) > 50:
        cadeia_erro = "Muitas ações habilitadas para executar em sequência de uma vez; execute as correções uma a uma."
    elif itens:
        cadeia = ajustar(todas, req.cabos, req.outros, grupos)
        cadeia["destrutivo"] = _destrutivo(todas, cadeia)
    return {"total_habilitados": len(habilitados), "itens": itens, "sem_mudanca": sem_mudanca, "problemas": problemas,
            "cadeia": cadeia, "cadeia_erro": cadeia_erro, "avisos": avisos}


@router.post("/descrever")
def descrever(req: DescreverAcoesRequest):
    """Frase em português de cada ação; erros de schema voltam em `erros` (sem frases)."""
    erros = validar_acoes(req.acoes, _grupos(req.projeto_codigo))
    if erros:
        return {"frases": [], "erros": erros}
    return {"frases": [descrever_acao(a) for a in req.acoes], "erros": []}
