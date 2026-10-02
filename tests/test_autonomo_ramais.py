"""TASK-031 fase B — ativos gerados pelos RAMAIS: Python IGUAL ao modal RAMAIS do JS real (decisão: os ramais entram no orçamento autônomo)."""
import random

import pytest

from services.autonomo import ramais as rm
from tests import oraculo_tela as o

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado (oráculo de paridade)")

TEXTOS = ["TROCAR 3 RS M AC", "trocar 2 rs m am", "TROCAR RS T AM", "TROCAR 4 RS M AA CP-REDE", "TROCAR 1 RS T AA CP REDE", "TROCAR 5 RS MAC",
          "TROCAR 2 RS TAM CPREDE", "TROCAR 0 RS M AC", "TROCAR 2 RS XYZ", "TROCAR 2 RS XYZ CP-REDE", "RS M AC", "IP 2X", "REC. CALÇADA", "",
          "TROCAR   7   RS   MAA", "  TROCAR 2 RS TAA  ", "TROCAR 12 RS M AM", "TROCAR 1 RS M AC TROCAR 2 RS T AM", "TROCAR 3 RS M AC"]


@pytest.mark.parametrize("semente", range(1, 9))
def test_ramais_iguais_ao_js(semente):
    rnd = random.Random(semente)
    for _ in range(15):
        textos = [rnd.choice(TEXTOS) for _ in range(rnd.randint(0, 8))]
        outros = [{"entidade": "ESTRUTURA", "operacao": "I", "ativo": "1-U4"}] * rnd.randint(0, 2)
        esperado = o.ramais_para_outros(textos, [dict(x) for x in outros])
        assert outros + rm.linhas_ramais(textos) == esperado


def test_exemplo_conhecido():
    linhas = rm.linhas_ramais(["TROCAR 3 RS M AC", "TROCAR 2 RS T AA CP-REDE"])
    assert [(l["operacao"], l["ativo"]) for l in linhas] == [("I", "60-MAC"), ("I", "40-TAM"), ("I", "2-CPREDE"), ("R", "45-MAC"), ("R", "4-MAA")]
    assert all(l["entidade"] == "0" and l["texto"] == "RAMAIS (GERADO)" for l in linhas)


def test_sem_troca_nao_gera_nada():
    assert rm.linhas_ramais(["IP 2X", "REC. CALÇADA", "", None, "TROCAR 0 RS M AC"]) == []
