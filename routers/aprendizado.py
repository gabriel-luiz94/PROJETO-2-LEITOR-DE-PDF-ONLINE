"""
routers/aprendizado.py — Aprendizado supervisionado do modo autônomo (TASK-059). Só administrador (as propostas viram AJUSTES do projeto).

O registro das correções acontece em `POST /api/obras` (quando a obra veio do autônomo). Aqui ficam: resumo, propostas (estatísticas e da IA),
aprovar/recusar, apagar o histórico e a análise opcional por IA (liberada depois de `IA_MINIMO_OBRAS` obras corrigidas).
Nada é aplicado sem o clique em "Aprovar". Não registrar conteúdo de planilhas em log (RULES Regra 11).
"""
import hashlib
import json
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Request

from middleware.auth_middleware import require_role
from routers.ai_chat import _checar_rate_limit, _ler_configuracao, resolver_credencial
from routers.validacao_regras import DEFAULT, regras_efetivas
from services.ajustes_planilhas import descrever_acao, validar_acoes
from services.autonomo import aprendizado, aprendizado_repo
from services.correcao_ia import interpretar_acoes
from services.validacao_ia import MODELOS_RESERVA, MODELOS_RESERVA_CLAUDE, MODELOS_RESERVA_OPENAI

router = APIRouter(prefix="/api/aprendizado", tags=["aprendizado"], dependencies=[Depends(require_role("admin"))])

_PROVEDORES = {
    "gemini": ("X-Gemini-Key", "gemini_api_key", "chamar_gemini", MODELOS_RESERVA),
    "openai": ("X-OpenAI-Key", "openai_api_key", "chamar_openai", MODELOS_RESERVA_OPENAI),
    "claude": ("X-Anthropic-Key", "anthropic_api_key", "chamar_claude", MODELOS_RESERVA_CLAUDE),
}


def _user(request: Request) -> dict:
    return getattr(request.state, "user", None) or {}


def _publica(p: dict) -> dict:
    d = {k: v for k, v in p.items() if k != "user_id"}
    d["pode_aprovar"] = bool(p.get("receita"))
    return d


@router.get("/resumo")
def resumo(request: Request, projeto: str):
    return aprendizado_repo.resumo(_user(request).get("user_id"), projeto)


@router.get("/propostas")
def propostas(request: Request, projeto: str, status: Optional[str] = None):
    if status and status not in aprendizado_repo.STATUS_PROPOSTA:
        raise HTTPException(status_code=400, detail="Status inválido.")
    return [_publica(p) for p in aprendizado_repo.listar_propostas(_user(request).get("user_id"), projeto, status)]


@router.get("/correcoes")
def correcoes(request: Request, projeto: str):
    """Obras corrigidas (sem os eventos brutos): arquivo, data e quantas correções."""
    return [{"execucao_id": r["execucao_id"], "arquivo": r["arquivo"], "atualizado_em": r["atualizado_em"], "correcoes": len(r["eventos"]),
             "descricoes": [aprendizado.descrever(e) for e in r["eventos"][:20]]}
            for r in aprendizado_repo.listar_registros(_user(request).get("user_id"), projeto)]


@router.post("/analisar")
def analisar(request: Request, projeto: str):
    """Recalcula os padrões (estatística) e guarda as propostas novas como pendentes."""
    return aprendizado_repo.reanalisar(_user(request).get("user_id"), projeto, regras_efetivas(projeto)[1])


@router.post("/propostas/{proposta_id}/recusar")
def recusar(proposta_id: int, request: Request):
    uid = _user(request).get("user_id")
    p = aprendizado_repo.buscar_proposta(uid, proposta_id)
    if not p:
        raise HTTPException(status_code=404, detail="Proposta não encontrada.")
    if p["status"] != "pendente":
        raise HTTPException(status_code=400, detail=f"A proposta já está {p['status']}.")
    aprendizado_repo.decidir_proposta(uid, proposta_id, "recusada")
    return {"status": "success"}


@router.post("/propostas/{proposta_id}/aprovar")
def aprovar(proposta_id: int, request: Request):
    """Cria o ajuste (receita) no projeto — o mesmo cadastro da tela de Ajustes, com validação e histórico — e marca a proposta aprovada."""
    from routers import validacao_ajustes as va
    uid = _user(request).get("user_id")
    p = aprendizado_repo.buscar_proposta(uid, proposta_id)
    if not p:
        raise HTTPException(status_code=404, detail="Proposta não encontrada.")
    if p["status"] != "pendente":
        raise HTTPException(status_code=400, detail=f"A proposta já está {p['status']}.")
    if not p.get("receita"):
        raise HTTPException(status_code=400, detail="Esta sugestão é só informativa (não vira ajuste automático). Recuse-a ou crie a regra manualmente.")
    receita = dict(p["receita"])
    existentes = va.ajustes_efetivos(p["projeto_codigo"])
    if any(r.get("id") == receita["id"] for r in existentes):
        raise HTTPException(status_code=400, detail="Este ajuste já existe no projeto.")
    va._salvar(p["projeto_codigo"], [va._limpa(r) for r in existentes] + [receita], _user(request).get("email"))
    aprendizado_repo.decidir_proposta(uid, proposta_id, "aprovada")
    return {"status": "success", "ajuste_id": receita["id"]}


@router.delete("/historico")
def apagar(request: Request, projeto: str):
    return aprendizado_repo.apagar_historico(_user(request).get("user_id"), projeto)


# ── análise opcional por IA ─────────────────────────────────────────────────
PROMPT_IA = """Você ajuda a configurar o AJUSTE automático das planilhas Cabos e Outros de orçamento de redes elétricas.
Abaixo estão correções que o usuário repetiu depois do processamento automático. Proponha AJUSTES que as reproduzam.
Responda SOMENTE com JSON: {"acoes": [ {"acao": ..., "tabela": "cabos"|"outros", ..., "motivo": "por que"} ]}.
Ações permitidas: substituir {"de","para","modo":"item"|"texto","palavra_inteira":true}; remover_ativo {"ativo"};
adicionar_ativo {"ativo","qtd","se_ja_existe":"ignorar","quando":{"tem":"CÓDIGO"}}; excluir_linhas {"onde":{"texto":"regex"}}.
Nunca mude a operação das linhas. Só proponha o que as correções sustentam; na dúvida, não proponha.

%s
"""


def _chamador(request: Request):
    """(chamar, None) ou (None, mensagem). Mesmo provedor/chave da configuração de IA do programa."""
    provider = _ler_configuracao("ai_provider", "gemini") or "gemini"
    provider = "claude" if provider == "anthropic" else provider
    if provider not in _PROVEDORES:
        provider = "gemini"
    header, chave_config, nome_chamar, modelos = _PROVEDORES[provider]
    api_key, origem = resolver_credencial(request.headers.get(header), _ler_configuracao(chave_config), provider)
    if not api_key:
        return None, "Sem chave de IA: informe a sua na configuração de IA ou peça ao administrador a chave padrão."
    if origem == "padrao":
        _checar_rate_limit(request)
    from services import validacao_ia

    async def chamar(texto):
        return await getattr(validacao_ia, nome_chamar)(api_key, list(modelos), texto, 0.0)
    return chamar, None


@router.post("/ia/sugerir")
async def sugerir_com_ia(request: Request, projeto: str):
    """Pede à IA ajustes que reproduzam as correções repetidas. Só depois de `IA_MINIMO_OBRAS` obras corrigidas. A IA recebe apenas
    códigos de ativo e contagens. Cada ação é validada pelo motor; o resultado entra como proposta PENDENTE (nunca é aplicado)."""
    uid = _user(request).get("user_id")
    registros = aprendizado_repo.listar_registros(uid, projeto)
    if len(registros) < aprendizado.IA_MINIMO_OBRAS:
        raise HTTPException(status_code=400, detail=f"A análise por IA é liberada com {aprendizado.IA_MINIMO_OBRAS} obras corrigidas (você tem {len(registros)}).")
    grupos = regras_efetivas(projeto)[1]
    padroes = aprendizado.detectar_padroes(registros, grupos, minimo=2, minimo_execucoes=2, consistencia_minima=0.5)
    if not padroes:
        return {"status": "sem_padroes", "mensagem": "Ainda não há correções repetidas para a IA analisar.", "novas": 0, "descartadas": []}
    chamar, indisponivel = _chamador(request)
    if indisponivel:
        return {"status": "sem_chave", "mensagem": indisponivel, "novas": 0, "descartadas": []}
    try:
        resposta = await chamar(PROMPT_IA % aprendizado.resumo_para_ia(padroes, registros))
        itens, _ = interpretar_acoes(resposta)
    except ValueError as e:
        return {"status": "erro", "mensagem": f"A IA não respondeu no formato esperado ({e}).", "novas": 0, "descartadas": []}
    except Exception as e:  # noqa: BLE001 — falha do provedor volta como status, não como 500
        return {"status": "erro", "mensagem": f"Falha ao consultar a IA: {type(e).__name__}.", "novas": 0, "descartadas": []}
    novos, descartadas = [], []
    for item in itens[:20]:
        acao = item["acao"]
        erros = validar_acoes([acao], grupos)
        if "operacao_nova" in acao:
            erros.append("a IA não pode mudar a operação da linha")
        if erros:
            descartadas.append({"acao": acao.get("acao"), "motivo": "; ".join(erros)[:200]})
            continue
        chave = "ia|" + json.dumps(acao, sort_keys=True, ensure_ascii=False)
        sufixo = hashlib.sha1(chave.encode("utf-8")).hexdigest()[:10]
        frase = descrever_acao(acao)
        motivo = item["motivo"] or "Sugerido pela IA a partir das suas correções."
        receita = aprendizado.receita_da_acao(acao, f"(IA) {frase}"[:100], f"{motivo}", sufixo)
        novos.append({"chave": chave, "tipo": "ia", "tabela": acao.get("tabela") or "ambos", "descricao": f"IA: {frase}. {motivo}"[:500],
                      "ocorrencias": 0, "execucoes": len(registros), "consistencia": 0.0, "receita": receita})
    n = aprendizado_repo.guardar_propostas(uid, projeto, novos, origem="ia")
    return {"status": "ok", "mensagem": f"{n} proposta(s) nova(s) da IA aguardando a sua aprovação.", "novas": n, "descartadas": descartadas}
