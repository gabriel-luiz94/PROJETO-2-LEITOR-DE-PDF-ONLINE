"""
routers/validacao_ajustes.py — Pré-visualização dos ajustes das planilhas (TASK-022, ADR-006).

Só simula: devolve o diff e as tabelas resultantes; quem aplica é o frontend, depois do aceite do usuário. Leitura aberta a
qualquer usuário autenticado (como /api/validacao/planilhas). Os grupos de ativos são os efetivos do projeto (os mesmos
das regras de domínio). Não registrar o conteúdo das planilhas em log.
"""
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from routers.validacao_regras import DEFAULT, regras_efetivas
from services.ajustes_planilhas import ajustar, descrever_acao, validar_acoes

router = APIRouter(prefix="/api/validacao/ajustes", tags=["validacao-ajustes"])


class AjustesPreviewRequest(BaseModel):
    acoes: List[Dict[str, Any]]
    cabos: List[Dict[str, Any]] = []
    outros: List[Dict[str, Any]] = []
    projeto_codigo: Optional[str] = None


class DescreverAcoesRequest(BaseModel):
    acoes: List[Dict[str, Any]]
    projeto_codigo: Optional[str] = None


def _grupos(projeto_codigo):
    return regras_efetivas(projeto_codigo or DEFAULT)[1]


@router.post("/preview")
def preview(req: AjustesPreviewRequest):
    grupos = _grupos(req.projeto_codigo)
    erros = validar_acoes(req.acoes, grupos)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})
    try:
        return ajustar(req.acoes, req.cabos, req.outros, grupos)
    except ValueError as e:
        raise HTTPException(status_code=400, detail={"erros": e.args[0]})


@router.post("/descrever")
def descrever(req: DescreverAcoesRequest):
    """Frase em português de cada ação; erros de schema voltam em `erros` (sem frases)."""
    erros = validar_acoes(req.acoes, _grupos(req.projeto_codigo))
    if erros:
        return {"frases": [], "erros": erros}
    return {"frases": [descrever_acao(a) for a in req.acoes], "erros": []}
