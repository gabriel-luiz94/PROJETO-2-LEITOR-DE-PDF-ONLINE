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
from services.validacao_ia import (
    MODELOS_RESERVA, MODELOS_RESERVA_CLAUDE, MODELOS_RESERVA_OPENAI,
    chamar_claude, chamar_gemini, chamar_openai, validar_com_ia,
)
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


# Por provedor (TASK-055): header da chave do usuário, nome da chave salva em `configuracoes`, o
# NOME da função de chamada (resolvido por `globals()` em `_preparar_ia`, nunca a função em si —
# os testes substituem `chamar_gemini` por um fake via `monkeypatch.setattr(rota, "chamar_gemini",
# falso)`, que só tem efeito se a busca for pelo nome no momento da chamada) e a lista de modelos
# de reserva. O campo "modelo" do cabeçalho do prompt salvo (ex. `modelo: gemini-3.1-flash-lite`)
# só é usado para Gemini — é sempre um ID Gemini, não existe equivalente para OpenAI/Claude.
_PROVEDORES_IA = {
    "gemini": ("X-Gemini-Key", "gemini_api_key", "chamar_gemini", MODELOS_RESERVA, True),
    "openai": ("X-OpenAI-Key", "openai_api_key", "chamar_openai", MODELOS_RESERVA_OPENAI, False),
    "claude": ("X-Anthropic-Key", "anthropic_api_key", "chamar_claude", MODELOS_RESERVA_CLAUDE, False),
}


def _preparar_ia(request: Request, prompt_id: str, projeto_codigo):
    """Prompt salvo + chave + limite, iguais para revisão e correção.

    Devolve (contexto, None) ou (None, (status, mensagem)) quando a IA não pode ser usada. `contexto` =
    (chamar, texto_prompt). Só levanta o 429 do limite por minuto (chave padrão compartilhada).
    Provider-aware (TASK-055): lê `ai_provider` salvo e monta a chamada certa (Gemini/OpenAI/Claude).
    """
    achado = buscar_prompt(prompt_id, projeto_codigo or "DEFAULT")
    if achado is None:
        return None, ("indisponivel", f"Prompt '{prompt_id}' não encontrado.")
    texto_prompt, _ = achado
    meta = resumo_meta(texto_prompt)
    if not meta.get("ativo"):
        return None, ("desativado", "Este prompt está desativado.")

    provider = _ler_configuracao("ai_provider", "gemini") or "gemini"
    if provider == "anthropic":
        provider = "claude"
    if provider not in _PROVEDORES_IA:
        provider = "gemini"
    header_nome, chave_config, nome_chamar, modelos_reserva, usa_modelo_do_prompt = _PROVEDORES_IA[provider]

    api_key, origem = resolver_credencial(request.headers.get(header_nome), _ler_configuracao(chave_config), provider)
    if not api_key:
        return None, ("sem_chave", "Sem chave de IA: informe a sua na configuração de IA ou peça ao administrador a chave padrão.")
    if origem == "padrao":
        _checar_rate_limit(request)  # 429 com mensagem clara (cota compartilhada)

    try:
        temperatura = float(parse_prompt(texto_prompt)[0].get("temperatura") or 0)
    except ValueError:
        temperatura = 0.0
    modelo_prompt = meta.get("modelo") if usa_modelo_do_prompt else None
    modelos = list(dict.fromkeys([m for m in [modelo_prompt] + modelos_reserva if m]))

    async def chamar(texto):
        return await globals()[nome_chamar](api_key, modelos, texto, temperatura)

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
