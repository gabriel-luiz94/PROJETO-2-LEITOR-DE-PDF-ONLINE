"""
routers/validacao.py — Validação das planilhas Cabos e Outros (camada 1, contrato).
"""
from fastapi import APIRouter, Request
from models import ValidacaoPlanilhasRequest
from services.sync_service import get_merged_orcamento
from routers.validacao_regras import buscar_regras
from services.regras_dominio import avaliar
from services.validacao_planilhas import validar_planilhas, validar_base, resumir

router = APIRouter(prefix="/api/validacao", tags=["validacao"])


@router.post("/planilhas")
def validar(req: ValidacaoPlanilhasRequest, request: Request):
    # Não registrar o conteúdo das planilhas em log (RULES Regra 11).
    achados = validar_planilhas(req.cabos, req.outros)
    if req.incluir_dominio:
        regras, _ = buscar_regras(req.projeto_codigo or "DEFAULT")
        achados += avaliar(regras, req.cabos, req.outros)
    if req.payload_calculo is not None:
        user = getattr(request.state, "user", None)
        base_rows = get_merged_orcamento(user["user_id"] if user else None)
        achados += validar_base(
            req.payload_calculo.cabos, req.payload_calculo.outros, req.projeto, base_rows
        )
    return {"achados": achados, "resumo": resumir(achados)}
