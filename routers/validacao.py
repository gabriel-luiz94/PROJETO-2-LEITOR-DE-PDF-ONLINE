"""
routers/validacao.py — Validação das planilhas Cabos e Outros (camada 1, contrato).
"""
from fastapi import APIRouter, Request
from models import AjustesIARequest, CorrecaoIARequest, ValidacaoPlanilhasRequest, ValidacaoIARequest
from services.sync_service import get_merged_orcamento
from routers.ai_chat import _checar_rate_limit, _ler_configuracao, resolver_credencial
from routers.validacao_prompts import buscar_prompt
from routers.validacao_regras import regras_efetivas
from services.regras_dominio import avaliar
from services.correcao_ia import ajustar_com_ia, corrigir_com_ia
from services.prompts_validacao import parse_prompt, resumo_meta
from services.validacao_ia import MODELOS_RESERVA, chamar_gemini, validar_com_ia
from services.validacao_planilhas import validar_planilhas, validar_base, resumir

router = APIRouter(prefix="/api/validacao", tags=["validacao"])


@router.post("/planilhas")
def validar(req: ValidacaoPlanilhasRequest, request: Request):
    # Não registrar o conteúdo das planilhas em log (RULES Regra 11).
    achados = validar_planilhas(req.cabos, req.outros)
    if req.incluir_dominio:
        regras, grupos, _ = regras_efetivas(req.projeto_codigo or "DEFAULT")
        achados += avaliar(regras, req.cabos, req.outros, grupos)
    if req.payload_calculo is not None:
        user = getattr(request.state, "user", None)
        base_rows = get_merged_orcamento(user["user_id"] if user else None)
        achados += validar_base(
            req.payload_calculo.cabos, req.payload_calculo.outros, req.projeto, base_rows
        )
    return {"achados": achados, "resumo": resumir(achados)}


def _preparar_ia(request: Request, prompt_id: str, projeto_codigo):
    """Prompt salvo + chave + limite, iguais para revisão e correção.

    Devolve (contexto, None) ou (None, (status, mensagem)) quando a IA não pode ser usada. `contexto` =
    (chamar, texto_prompt). Só levanta o 429 do limite por minuto (chave padrão compartilhada).
    """
    achado = buscar_prompt(prompt_id, projeto_codigo or "DEFAULT")
    if achado is None:
        return None, ("indisponivel", f"Prompt '{prompt_id}' não encontrado.")
    texto_prompt, _ = achado
    meta = resumo_meta(texto_prompt)
    if not meta.get("ativo"):
        return None, ("desativado", "Este prompt está desativado.")

    api_key, origem = resolver_credencial(request.headers.get("X-Gemini-Key"), _ler_configuracao("gemini_api_key"))
    if not api_key:
        return None, ("sem_chave", "Sem chave de IA: informe a sua na configuração de IA ou peça ao administrador a chave padrão.")
    if origem == "padrao":
        _checar_rate_limit(request)  # 429 com mensagem clara (cota compartilhada)

    try:
        temperatura = float(parse_prompt(texto_prompt)[0].get("temperatura") or 0)
    except ValueError:
        temperatura = 0.0
    modelos = list(dict.fromkeys([m for m in [meta.get("modelo")] + MODELOS_RESERVA if m]))

    async def chamar(texto):
        return await chamar_gemini(api_key, modelos, texto, temperatura)

    return (chamar, texto_prompt), None


@router.post("/ia")
async def validar_ia(req: ValidacaoIARequest, request: Request):
    """Camada 3: revisão por IA com o prompt salvo. Nunca devolve erro HTTP por falha da IA — devolve `status`
    para a tela mostrar os achados das camadas 1 e 2 (que vêm de /planilhas) mesmo sem IA.

    Não registrar o conteúdo das planilhas em log.
    """
    contexto, indisponivel = _preparar_ia(request, req.prompt_id, req.projeto_codigo)
    if indisponivel:
        return {"status": indisponivel[0], "achados": [], "mensagem": indisponivel[1], "lotes": 0,
                "truncado": False, "descartados": 0, "resumo": resumir([])}
    chamar, texto_prompt = contexto
    resultado = await validar_com_ia(chamar, texto_prompt, req.cabos, req.outros, req.achados_previos)
    resultado["resumo"] = resumir(resultado["achados"])
    return resultado


@router.post("/corrigir")
async def corrigir_ia(req: CorrecaoIARequest, request: Request):
    """Correção assistida: a IA PROPÕE correções só para as linhas citadas em `achados`; nada é aplicado aqui.
    Cada proposta é reconferida (linha existente, muda algo, passa de novo na camada 1). Falha da IA volta como
    `status` com HTTP 200.

    Não registrar o conteúdo das planilhas em log.
    """
    contexto, indisponivel = _preparar_ia(request, req.prompt_id, req.projeto_codigo)
    if indisponivel:
        return {"status": indisponivel[0], "correcoes": [], "descartadas": [], "mensagem": indisponivel[1],
                "lotes": 0, "truncado": False}
    chamar, texto_prompt = contexto
    return await corrigir_com_ia(chamar, texto_prompt, req.cabos, req.outros, req.achados)


@router.post("/ajustes-ia")
async def ajustes_ia(req: AjustesIARequest, request: Request):
    """Ajustes por IA (TASK-025): a IA PROPÕE ações estruturadas só para os achados enviados; o motor determinístico
    calcula o diff (nada é aplicado aqui). Falha da IA volta como `status` com HTTP 200.

    Não registrar o conteúdo das planilhas em log.
    """
    contexto, indisponivel = _preparar_ia(request, req.prompt_id, req.projeto_codigo)
    if indisponivel:
        return {"status": indisponivel[0], "acoes": [], "diff": None, "destrutivo": False, "descartadas": [],
                "mensagem": indisponivel[1], "lotes": 0, "truncado": False}
    chamar, texto_prompt = contexto
    grupos = regras_efetivas(req.projeto_codigo or "DEFAULT")[1]
    return await ajustar_com_ia(chamar, texto_prompt, req.cabos, req.outros, req.achados, grupos)
