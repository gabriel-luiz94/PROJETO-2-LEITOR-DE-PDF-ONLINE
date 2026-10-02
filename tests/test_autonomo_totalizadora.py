"""TASK-031 fase B — Totalizadora e payload do orçamento em Python IGUAIS aos do JS real (syncTotalizadora / obterPayloadCalculo)."""
import math
import random

import pytest

from services.autonomo import totalizadora as tt
from tests import oraculo_tela as o
from tests.test_autonomo_aplicar import ATIVOS_CABOS, ATIVOS_OUTROS

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado (oráculo de paridade)")

CABOS_EXTRA = ["CAA2 ABC 35,5 m", "CA 4 AB 10 | CA4", "12,5 m | CAA2", "P 50", "CAZ 3 ABC 1.5 m", "CAA2 ABC . m", "XX", "P 50 m", "  CU 16 m  "]
OUTROS_EXTRA = ["2-U4 *1-CFU -3-X 4X5 2x-Y 1.5-ROCO", "  1-A   2-B  ", "*2-TR3 1-PR15", "DT11/300", "5-", "-U4", "3Xabc", "1-ÉÇÃO"]
REGRAS = [
    [],
    [tt.REGRA_PADRAO],
    [{"origem": "CABOS", "op_de": "", "ativo_de": "CAA%", "acao": "SUBST", "op_para": "", "ativo_para": "CAA2X", "fator": 1.1, "arredondamento": "PARA CIMA", "val_min": 40, "val_max": ""}],
    [{"origem": "OUTROS", "op_de": "I", "ativo_de": "U[0-9]", "acao": "ADICAO", "op_para": "*i", "ativo_para": "ESCORA", "fator": "2", "arredondamento": "INTEIRO", "val_min": "", "val_max": 6},
     {"origem": "", "op_de": "", "ativo_de": "%", "acao": "SUBST", "op_para": "", "ativo_para": "", "fator": "abc", "arredondamento": "PARA BAIXO", "val_min": "x", "val_max": ""},
     {"origem": "cabos", "op_de": "i", "ativo_de": "(P|CU)\\d*", "acao": "ADICAO", "op_para": "", "ativo_para": "", "fator": 0.333, "arredondamento": "NORMAL", "val_min": 0.5, "val_max": 20},
     {"origem": "OUTROS", "op_de": "", "ativo_de": "([", "acao": "SUBST", "op_para": "", "ativo_para": "Z", "fator": 1, "arredondamento": "NORMAL", "val_min": "", "val_max": ""}],
]
BASE = [{"ativo": "U4", "desc_ativo": "Estrutura U4", "projeto": ""}, {"ativo": "CAA2", "codigo": "C1", "componente": "Cabo", "projeto": "P1"},
        {"ativo": "CFU", "desc_codigo": "Chave", "projeto": "OUTRO"}, {"ativo": "TR3", "codigo": "TR3", "desc_ativo": "", "componente": "", "desc_codigo": "", "projeto": " "}]


def tabelas(rnd):
    cabos = [{"entidade": "CABO", "operacao": rnd.choice(["I", "R", "M", "*I"]), "ativo": rnd.choice(ATIVOS_CABOS + CABOS_EXTRA),
              **({"qtdAtivos": rnd.choice([1, 2, 3, "-2", "4", 0, "x"])} if rnd.random() < 0.8 else {})} for _ in range(rnd.randint(0, 7))]
    outros = [{"entidade": rnd.choice(["ESTRUTURA", "POSTE", "CHAVE"]), "operacao": rnd.choice(["I", "R", "M", "*R", ""]),
               "ativo": rnd.choice(ATIVOS_OUTROS + OUTROS_EXTRA)} for _ in range(rnd.randint(0, 6))]
    return cabos, outros


@pytest.mark.parametrize("semente", range(1, 11))
def test_totalizadora_e_payload_iguais_ao_js(semente):
    rnd = random.Random(semente)
    linhas = 0
    for _ in range(12):
        cabos, outros = tabelas(rnd)
        regras = rnd.choice(REGRAS)
        proj = rnd.choice(["P1", "p1", "", "OUTRO"])
        js = o.totalizadora_e_payload(cabos, outros, regras, BASE, proj)
        tot = tt.montar_totalizadora(cabos, outros, regras, BASE, proj)
        assert tt.normalizar_para_json(tot) == js["totalizadora"]
        assert tt.payload_calculo(tot) == js["payload"]
        linhas += len(tot)
    assert linhas > 40


@pytest.mark.parametrize("n,esperado", [(5.0, "5"), (35.5, "35.5"), (0.1 + 0.2, "0.30000000000000004"), (1e21, "1e+21"), (1e-7, "1e-7"), (0.000001, "0.000001"),
                                        (123456789012345680000.0, "123456789012345680000"), (-2.5, "-2.5"), (float("nan"), "NaN"), (36.75, "36.75")])
def test_numero_para_texto_como_o_js(n, esperado):
    assert tt.js_numero_para_texto(n) == esperado


@pytest.mark.parametrize("x,esperado", [(0.5, 1), (1.5, 2), (2.5, 3), (-0.5, -0.0), (-1.5, -1), (2.4999, 2), (0.49999999999999994, 0)])
def test_arredondamento_do_js_nao_e_o_do_python(x, esperado):
    assert tt.js_round(x) == esperado
    assert tt.js_round(2.5) == 3 and round(2.5) == 2   # o round() do Python (banqueiro) daria 2


@pytest.mark.parametrize("v,esperado", [("12,5", 12.0), ("1.2.3", 1.2), (".5", 0.5), ("abc", None), ("", None), (" 7x", 7.0), (3, 3.0), (None, None), ("-4.5e1", -45.0),
                                        ("1e", 1.0), (True, None)])
def test_parse_float_como_o_js(v, esperado):
    r = tt.js_parse_float(v)
    assert (esperado is None and math.isnan(r)) or r == esperado


def test_exemplo_conhecido():
    tot = tt.montar_totalizadora([{"entidade": "CABO", "operacao": "I", "ativo": "CAA2 ABC 35 m", "qtdAtivos": 3}],
                                 [{"entidade": "ESTRUTURA", "operacao": "I", "ativo": "DT11/300 2-U4 *1-CFU"}], [], BASE, "P1")
    assert [(r["origem"], r["ativo"], r["qtd"]) for r in tot] == [("CABOS", "CAA2", 35.0)] * 3 + [("OUTROS", "DT11/300", 1.0), ("OUTROS", "U4", 2.0), ("OUTROS", "CFU", -1.0)]
    assert tt.payload_calculo(tot)["outros"][-1]["ativo"] == "*1-CFU"


def regra(arr, fator, **kw):
    return {"origem": "", "op_de": "", "ativo_de": "", "acao": "SUBST", "op_para": "", "ativo_para": "", "fator": fator, "arredondamento": arr,
            "val_min": kw.get("min", ""), "val_max": kw.get("max", "")}


@pytest.mark.parametrize("arr,fator", [("PARA CIMA", 1.1), ("PARA BAIXO", 1.1), ("INTEIRO", 1.5), ("INTEIRO", 2.5), ("NORMAL", 1.005), ("", 1.0), ("PARA CIMA", 0)])
def test_arredondamentos_e_fator_zero_iguais_ao_js(arr, fator):
    cabos = [{"entidade": "CABO", "operacao": "I", "ativo": "CAA2 ABC 35,3 m"}]
    outros = [{"entidade": "ESTRUTURA", "operacao": "I", "ativo": "2-U4 1-CFU 3-X"}]
    js = o.totalizadora_e_payload(cabos, outros, [regra(arr, fator)], BASE, "P1")
    tot = tt.montar_totalizadora(cabos, outros, [regra(arr, fator)], BASE, "P1")
    assert tt.normalizar_para_json(tot) == js["totalizadora"]
    assert tt.payload_calculo(tot) == js["payload"]


def test_quantidade_zero_nao_vai_para_o_orcamento():
    tot = tt.montar_totalizadora([], [{"entidade": "ESTRUTURA", "operacao": "I", "ativo": "2-U4 1-CFU"}], [regra("NORMAL", 0)], [], "P1")
    assert all(r["qtd"] == 0 for r in tot) and tt.payload_calculo(tot) == {"cabos": [], "outros": []}
