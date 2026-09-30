"""
routers/validacao.py — Validação das planilhas Cabos e Outros (camada 1, contrato).
"""
from fastapi import APIRouter, Request
from models import ValidacaoPlanilhasRequest, ValidacaoIARequest
from services.sync_service import get_merged_orcamento
from routers.ai_chat import _checar_rate_limit, _ler_configuracao, resolver_credencial
from routers.validacao_prompts import buscar_prompt
from routers.validacao_regras import buscar_regras
from services.regras_dominio import avaliar
from services.prompts_validacao import parse_prompt, resumo_meta
from services.validacao_ia import MODELOS_RESERVA, chamar_gemini, validar_com_ia
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


@router.post("/ia")
async def validar_ia(req: ValidacaoIARequest, request: Request):
    """Camada 3: revisão por IA com o prompt salvo. Nunca devolve erro HTTP por falha da IA — devolve `status`
    para a tela mostrar os achados das camadas 1 e 2 (que vêm de /planilhas) mesmo sem IA.

    Não registrar o conteúdo das planilhas em log.
    """
    def sem_ia(status, mensagem):
        return {"status": status, "achados": [], "mensagem": mensagem, "lotes": 0, "truncado": False,
                "descartados": 0, "resumo": resumir([])}

    achado = buscar_prompt(req.prompt_id, req.projeto_codigo or "DEFAULT")
    if achado is None:
        return sem_ia("indisponivel", f"Prompt '{req.prompt_id}' não encontrado.")
    texto_prompt, _ = achado
    meta = resumo_meta(texto_prompt)
    if not meta.get("ativo"):
        return sem_ia("desativado", "A revisão por IA está desativada para este prompt.")

    api_key, origem = resolver_credencial(request.headers.get("X-Gemini-Key"), _ler_configuracao("gemini_api_key"))
    if not api_key:
        return sem_ia("sem_chave", "Sem chave de IA: informe a sua na configuração de IA ou peça ao administrador a chave padrão.")
    if origem == "padrao":
        _checar_rate_limit(request)  # 429 com mensagem clara (cota compartilhada)

    try:
        temperatura = float(parse_prompt(texto_prompt)[0].get("temperatura") or 0)
    except ValueError:
        temperatura = 0.0
    modelos = list(dict.fromkeys([m for m in [meta.get("modelo")] + MODELOS_RESERVA if m]))

    async def chamar(texto):
        return await chamar_gemini(api_key, modelos, texto, temperatura)

    resultado = await validar_com_ia(chamar, texto_prompt, req.cabos, req.outros, req.achados_previos)
    resultado["resumo"] = resumir(resultado["achados"])
    return resultado
