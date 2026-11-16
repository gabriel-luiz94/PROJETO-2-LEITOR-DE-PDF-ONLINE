"""
routers/casos.py — Casos de referência por projeto (TASK-178): guardar pedidos com o resultado esperado e conferir a base (atual ou uma candidata) antes de gravar.

  GET    /api/casos?projeto=                         lista os casos do projeto
  POST   /api/casos                                  {projeto, caso}: guarda um caso (pedido + resultado esperado)
  DELETE /api/casos/{id}?projeto=
  POST   /api/casos/conferir                         {projeto, ids?, base_csv?}: roda os casos (base atual, ou a candidata de `base_csv`, sem gravar) → igual / mudou / erro
  POST   /api/casos/aceitar                          {projeto, ids, nota, base_csv?}: marca "mudança esperada" (o resultado de hoje vira o esperado, com a nota)
"""
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel

from middleware.auth_middleware import get_current_user_from_state
from services import casos_referencia as cr
from services import pacote_projeto as pp
from services.sync_service import get_merged_orcamento, normalizar_projeto

router = APIRouter(prefix="/api/casos", tags=["casos"])


class CasoPayload(BaseModel):
    projeto: str
    caso: Dict[str, Any]


class ConferirPayload(BaseModel):
    projeto: str
    ids: Optional[List[int]] = None
    base_csv: Optional[str] = None


class AceitarPayload(ConferirPayload):
    nota: str = ""


def _projeto(codigo) -> dict:
    p = pp.projeto_do_cadastro(codigo)
    if not p:
        raise HTTPException(status_code=404, detail=f"Projeto '{codigo}' não está cadastrado.")
    return p


def linhas_da_base(proj, user_id=None, base_csv=None) -> list:
    """Linhas da base para o cálculo: a do usuário/master; com `base_csv`, a candidata no lugar das linhas deste projeto (sem gravar)."""
    rows = get_merged_orcamento(user_id)
    if not base_csv:
        return rows
    try:
        novas = pp._linhas_do_csv(base_csv)
    except pp.ErroPacote as e:
        raise HTTPException(status_code=400, detail=str(e))
    alvo = normalizar_projeto(proj["nome"])
    novas = [dict(l, projeto=alvo) for l in novas]                                        # a base candidata vale para este projeto
    return [r for r in rows if normalizar_projeto(r.get("projeto")) != alvo] + novas


def _contexto_calculo(proj):
    from routers.vinculacao_config import perfil_do_projeto, postes_do_projeto
    return postes_do_projeto(proj["codigo"]), perfil_do_projeto(proj["codigo"])


def conferir_projeto(proj, user_id=None, ids=None, base_csv=None, casos=None) -> dict:
    todos = casos if casos is not None else cr.listar(proj["codigo"])
    if ids is not None:
        todos = [c for c in todos if c.get("id") in set(ids)]
    postes, perfil = _contexto_calculo(proj)
    res = cr.conferir(todos, linhas_da_base(proj, user_id, base_csv), proj["nome"], postes, perfil)
    res["relatorio"] = cr.relatorio_markdown(res)
    return res


@router.get("")
def listar(request: Request, projeto: str):
    get_current_user_from_state(request)
    p = _projeto(projeto)
    return {"projeto": p, "casos": [{"id": c["id"], "nome": c["nome"], "origem": c["origem"], "nota": c.get("nota", ""), "itens": len(c["esperado"]), "historico": len(c.get("historico") or [])} for c in cr.listar(p["codigo"])]}


@router.post("")
def guardar(request: Request, payload: CasoPayload):
    user = get_current_user_from_state(request)
    p = _projeto(payload.projeto)
    try:
        return cr.gravar(p["codigo"], payload.caso, user.get("email"))
    except cr.ErroCaso as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.delete("/{caso_id}")
def excluir(request: Request, caso_id: int, projeto: str):
    get_current_user_from_state(request)
    if not cr.excluir(_projeto(projeto)["codigo"], caso_id):
        raise HTTPException(status_code=404, detail="Caso não encontrado.")
    return {"status": "ok"}


@router.post("/conferir")
def conferir(request: Request, payload: ConferirPayload):
    user = get_current_user_from_state(request)
    return conferir_projeto(_projeto(payload.projeto), user.get("user_id"), payload.ids, payload.base_csv)


@router.post("/aceitar")
def aceitar(request: Request, payload: AceitarPayload):
    user = get_current_user_from_state(request)
    p = _projeto(payload.projeto)
    casos = [c for c in cr.listar(p["codigo"]) if not payload.ids or c["id"] in set(payload.ids)]
    postes, perfil = _contexto_calculo(p)
    rows = linhas_da_base(p, user.get("user_id"), payload.base_csv)
    feitos = []
    try:
        for c in casos:
            obtido = cr.calcular(c, rows, p["nome"], postes, perfil)
            if cr.comparar(c.get("esperado") or [], obtido, c.get("tolerancia", cr.TOLERANCIA_PADRAO)):
                cr.aceitar_mudanca(p["codigo"], c, obtido, user.get("email"), payload.nota)
                feitos.append(c["id"])
    except cr.ErroCaso as e:
        raise HTTPException(status_code=400, detail=str(e))
    return {"aceitos": feitos}
