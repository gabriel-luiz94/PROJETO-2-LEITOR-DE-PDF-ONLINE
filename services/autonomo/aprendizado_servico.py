"""
services/autonomo/aprendizado_servico.py — Orquestração do aprendizado (TASK-059): junta as regras do projeto, o motor do leitor e o armazenamento.
Os imports de `routers.*` são tardios (evitam ciclo e deixam este módulo barato de carregar).
"""
import logging
import time

from services.autonomo import aprendizado_repo

logger = logging.getLogger(__name__)
INTERVALO_AUTOMATICO_S = 20     # a mineração de regras do leitor (simulações no motor) roda no máximo a cada 20 s por usuário/projeto quando disparada por um "salvar"
_ultima = {}


def simulador():
    """`simular(entrada, proc, cls)` com o MESMO motor da tela (QuickJS, processo separado). Levanta se o motor não está disponível."""
    from services.autonomo.leitor_js import obter_leitor
    leitor = obter_leitor()
    return lambda entrada, proc, cls: leitor.processar_lote(entrada, proc, cls, "extracao")


def reanalisar_tudo(user_id, projeto, forcar=True) -> dict:
    """Ajustes (estatística) + regras do leitor (simulação). A parte do leitor é opcional: sem o motor (quickjs) ou sem regras cadastradas, só avisa."""
    from routers import regras_leitor as rl
    from routers.validacao_regras import regras_efetivas
    saida = aprendizado_repo.reanalisar(user_id, projeto, regras_efetivas(projeto)[1])
    agora = time.monotonic()
    if not forcar and agora - _ultima.get((user_id, projeto), -1e9) < INTERVALO_AUTOMATICO_S:
        saida["leitor"] = {"aviso": "Análise das regras do leitor adiada (use 'Analisar agora')."}
        return saida
    _ultima[(user_id, projeto)] = agora
    try:
        proc, cls = rl._get_regras("processamento", projeto), rl._get_regras("classificacao", projeto)
        if not proc and not cls:
            saida["leitor"] = {"aviso": "O projeto não tem regras do leitor cadastradas."}
        else:
            saida["leitor"] = aprendizado_repo.reanalisar_leitor(user_id, projeto, simulador(), proc, cls,
                                                                 lambda tabela, regra: rl.validar_regras(tabela, [regra]))
    except Exception as e:  # noqa: BLE001 — sem o motor (ou com regra travada) o aprendizado de ajustes segue valendo
        logger.warning("Regras do leitor não analisadas: %s", type(e).__name__)
        saida["leitor"] = {"aviso": f"Regras do leitor não analisadas ({type(e).__name__})."}
    return saida
