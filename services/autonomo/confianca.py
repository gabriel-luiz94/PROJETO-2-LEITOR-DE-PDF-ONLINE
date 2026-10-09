"""
services/autonomo/confianca.py — Nível de confiança do modo autônomo (TASK-059, etapa 3).

Medida: das últimas N obras do autônomo que o USUÁRIO REVISOU (abriu, corrigiu se precisou e salvou), quantas ele deixou SEM nenhuma correção.
É calculada por usuário (dono das obras do autônomo), projeto e TIPO de arquivo (extensão: dxf, pdf, dwg).
  • observando  — menos de N obras revisadas desse tipo: ainda não dá para confiar;
  • revisar     — taxa de acerto abaixo do mínimo (padrão 90%);
  • confiável   — taxa de acerto ≥ mínimo com N revisadas.
A janela RECOMEÇA quando o usuário aprova uma proposta de ajuste/regra (o comportamento mudou: precisa ser validado de novo).

Único efeito sobre o autônomo: com o projeto marcado para isso, nível confiável e sem anomalia (muitas exclusões de uma vez), as exclusões que hoje
sempre esperam confirmação passam a ser CONFIRMADAS sozinhas — e a execução registra isso no relatório (dá para reverter). Desligado por padrão.
"""
import os

from services.autonomo import aprendizado_repo

JANELA_PADRAO = 10
TAXA_PADRAO = 0.9
LIMITE_ANOMALIA = 0.25     # exclusões pendentes acima de 25% das linhas das tabelas = algo estranho: pede confirmação mesmo com confiança
MINIMO_ANOMALIA = 5        # ...mas até 5 exclusões sempre é aceitável (tabelas pequenas)
ROTULOS = {"observando": "Em observação", "revisar": "Precisa de revisão", "confiavel": "Confiável"}


def tipo_do_arquivo(nome) -> str:
    return os.path.splitext(str(nome or ""))[1].lower().lstrip(".") or "(sem extensão)"


def avaliar(revisoes: list, janela: int = JANELA_PADRAO, taxa_minima: float = TAXA_PADRAO) -> dict:
    """`revisoes` = [{arquivo, atualizado_em, correcoes}] (qualquer ordem). Considera as `janela` mais recentes."""
    ult = sorted(revisoes, key=lambda r: r["atualizado_em"], reverse=True)[:janela]
    acertos = sum(1 for r in ult if r["correcoes"] == 0)
    taxa = acertos / len(ult) if ult else None
    if len(ult) < janela:
        nivel = "observando"
    elif taxa >= taxa_minima:
        nivel = "confiavel"
    else:
        nivel = "revisar"
    return {"nivel": nivel, "rotulo": ROTULOS[nivel], "revisadas": len(ult), "janela": janela, "acertos": acertos,
            "taxa": None if taxa is None else round(taxa, 3), "taxa_minima": taxa_minima,
            "ultimas": [{"arquivo": r["arquivo"], "atualizado_em": r["atualizado_em"], "correcoes": r["correcoes"]} for r in ult]}


def estados(user_id, projeto, cfg: dict) -> dict:
    """{tipo: avaliação} das revisões do autônomo desde a última aprovação de proposta."""
    janela = int(cfg.get("confianca_janela") or JANELA_PADRAO)
    taxa = float(cfg.get("confianca_taxa") or TAXA_PADRAO)
    desde = aprendizado_repo.ultima_aprovacao(user_id, projeto) or ""
    por_tipo = {}
    for r in aprendizado_repo.revisoes_autonomo(user_id, projeto):
        if r["atualizado_em"] >= desde:
            por_tipo.setdefault(tipo_do_arquivo(r["arquivo"]), []).append(r)
    return {t: avaliar(rs, janela, taxa) for t, rs in sorted(por_tipo.items())}


def estado(user_id, projeto, tipo, cfg: dict) -> dict:
    return estados(user_id, projeto, cfg).get(tipo) or avaliar([], int(cfg.get("confianca_janela") or JANELA_PADRAO), float(cfg.get("confianca_taxa") or TAXA_PADRAO))


def decidir_autoconfirmacao(user_id, projeto, arquivo, n_linhas: int, n_pendentes: int, cfg: dict) -> dict:
    """{auto: bool, motivo: str, estado: {...}} — confirmar sozinho as exclusões deste arquivo? Nunca levanta (na dúvida, não confirma)."""
    try:
        est = estado(user_id, projeto, tipo_do_arquivo(arquivo), cfg)
    except Exception as e:  # noqa: BLE001
        return {"auto": False, "motivo": f"confiança indisponível ({type(e).__name__})", "estado": None}
    resumo = {k: est[k] for k in ("nivel", "revisadas", "janela", "acertos", "taxa", "taxa_minima")}
    if projeto not in (cfg.get("autoconfirmar_projetos") or []):
        return {"auto": False, "motivo": "confirmação automática desligada para este projeto", "estado": resumo}
    if est["nivel"] != "confiavel":
        return {"auto": False, "motivo": f"confiança insuficiente ({est['rotulo'].lower()})", "estado": resumo}
    if n_pendentes > max(MINIMO_ANOMALIA, LIMITE_ANOMALIA * max(n_linhas, 1)):
        return {"auto": False, "motivo": f"{n_pendentes} exclusões de {n_linhas} linhas parece anormal", "estado": resumo}
    return {"auto": True, "motivo": "confiável e dentro do esperado", "estado": resumo}
