"""TASK-031 fase B — Totalizadora e payload do orçamento em Python IGUAIS aos do JS real (syncTotalizadora / obterPayloadCalculo)."""
import math
import random

import pytest

from services.autonomo import totalizadora as tt
from tests import oraculo_tela as o
from tests.test_autonomo_aplicar import ATIVOS_CABOS, ATIVOS_OUTROS

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado (oráculo de paridade)")

CABOS_EXTRA = ["CAA2 ABC 35,5 m", "CA 4 AB 10 | CA4", "12,5 m | CAA2", "P 50", "CAZ 3 ABC 1.5 m", "CAA2 ABC . m", "XX", "P 50 m", "  CU 16 m  "]
OUTROS_EXTRA = ["2-U4 *1-CFU -3-X 4X5 2x-Y 1.5-ROCO", "  1-A   2-B  ", "*2-TR3 1-PR15", "DT11/300", "5-", "-U4", "3Xabc", "1-ÉÇÃO", "0,7-M335 *0,5-TR3"]
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


def test_quantidade_decimal_com_virgula_em_outros():
    """Pedido real do usuário: "0,7-M335" não aparecia certo na Totalizadora (TASK-050) — o regex de
    syncTotalizadora/montar_totalizadora só reconhecia "." como decimal, então a linha inteira caía
    no fallback e virava um ativo chamado "0,7-M335" com quantidade 1."""
    tot = tt.montar_totalizadora([], [{"entidade": "ESTRUTURA", "operacao": "I", "ativo": "0,7-M335 *0,5-X 2-Y"}], [], BASE, "P1")
    assert [(r["ativo"], r["qtd"]) for r in tot] == [("M335", 0.7), ("X", -0.5), ("Y", 2.0)]


def test_quantidade_decimal_com_virgula_igual_ao_js():
    cabos = []
    outros = [{"entidade": "ESTRUTURA", "operacao": "I", "ativo": "0,7-M335 *0,5-TR3 2-U4"}]
    js = o.totalizadora_e_payload(cabos, outros, [], BASE, "P1")
    tot = tt.montar_totalizadora(cabos, outros, [], BASE, "P1")
    assert tt.normalizar_para_json(tot) == js["totalizadora"]
    assert tt.payload_calculo(tot) == js["payload"]


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


# ── TASK-079: ativos com parâmetros (N1(XLPE,DT)) — o espelho em Python continua igual ao JS real ──
def _linha(ativo, comp, cod, proj="PERNAMBUCO", desc="D"):
    return {"ativo": ativo, "desc_ativo": desc if ativo else "", "componente": comp, "projeto": proj, "codigo": cod, "desc_codigo": "d"}


BASE_PARAM = [_linha("N1=N1(cabo=XLPE, poste=DT)", "N1", "c1", desc="ESTRUTURA N1"), _linha("", "N1", "c2"),
              _linha("N1(cabo=XLPE)", "N1.XLPE", "x1", desc="N1 XLPE"), _linha("N1(cabo=ALUM)", "N1.ALUM", "a1"),
              _linha("N1(poste=DT)", "N1.DT", "d1"), _linha("N1(poste=CC)", "N1.CC", "k1"), _linha("P50", "P", "p1", proj="", desc="POSTE 50")]
OUTROS_PARAM = ["3-N1", "2-N1(ALUM, CC)", "1-N1(XLIPE)", "1-N1(poste=CC,cabo=ALUM)", "2-N1(XLPE) 1-P50", "*1-N1(alum,dt) 4-P50(X)", "N1(ALUM,CC,XX)", "5-N1()", "2-N1(A B)", "1-n1(xlpe,dt)"]
REGRAS_PARAM = [[], [{"origem": "OUTROS", "op_de": "", "ativo_de": "N1", "acao": "ADICAO", "op_para": "", "ativo_para": "ESCORA", "fator": 2, "arredondamento": "NORMAL", "val_min": "", "val_max": ""}],
                [{"origem": "OUTROS", "op_de": "", "ativo_de": "N1%", "acao": "SUBST", "op_para": "", "ativo_para": "N1X", "fator": 1, "arredondamento": "NORMAL", "val_min": "", "val_max": ""}]]


@pytest.mark.parametrize("semente", range(1, 7))
def test_totalizadora_com_parametros_igual_ao_js(semente):
    rnd = random.Random(1000 + semente)
    visto = 0
    for _ in range(10):
        outros = [{"entidade": "0", "operacao": rnd.choice(["I", "*I", "R"]), "ativo": " ".join(rnd.sample(OUTROS_PARAM, rnd.randint(1, 3)))} for _ in range(rnd.randint(1, 4))]
        regras = rnd.choice(REGRAS_PARAM)
        proj = rnd.choice(["PERNAMBUCO", "pernambuco", "", "OUTRO"])
        js = o.totalizadora_e_payload([], outros, regras, BASE_PARAM, proj)
        tot = tt.montar_totalizadora([], outros, regras, BASE_PARAM, proj)
        assert tt.normalizar_para_json(tot) == js["totalizadora"], (outros, regras, proj)
        assert tt.payload_calculo(tot) == js["payload"]
        visto += sum(1 for r in tot if "(" in r["ativo"])
    assert visto > 5


def test_totalizadora_normaliza_e_descreve_pelo_nome_base():
    tot = tt.montar_totalizadora([], [{"entidade": "0", "operacao": "I", "ativo": "3-N1 2-N1(ALUM, CC) 1-N1(XLIPE)"}], [], BASE_PARAM, "PERNAMBUCO")
    assert [r["ativo"] for r in tot] == ["N1(XLPE,DT)", "N1(ALUM,CC)", "N1(XLIPE)"]
    assert tot[0]["desc"] == "ESTRUTURA N1" and not tot[0]["naoEncontrado"]
    assert tot[2]["naoEncontrado"] and "XLIPE" in tot[2]["desc"]


def test_vinculo_e_classificacao_ignoram_os_parametros():
    from services.autonomo import vinculacao as vc
    assert vc.codigo_estrutura("1-N1(XLPE,DT)") == "N1" and vc.codigo_estrutura("1-U4") == "U4" and vc.tipo_estrutura("1-N1(XLPE,DT)") == "1"


# ── TASK-080: operadores (%, |, !=) nos grupos — espelho em Python igual ao JS real ──
BASE_OPERADORES = [_linha("N1=N1(cabo=XLPE, poste=DT)", "N1", "c1", desc="ESTRUTURA N1"), _linha("N1(cabo=ALUM|COBRE)", "N1.NU", "n1"),
                   _linha("N1(poste=D%)", "N1.D", "p1"), _linha("N1(cabo!=XLPE, poste=C%)", "N1.NC", "q1"), _linha("N1(cabo=XLPE)", "N1.X", "x1")]
OUTROS_OPERADORES = ["3-N1", "2-N1(ALUM,D12)", "1-N1(COBRE,CC)", "1-N1(XLPE,XX)", "1-N1(ALUM,DT) 2-N1(poste=C1,cabo=ALUM)", "1-N1(A|B)", "1-N1(ALUM,D%)", "2-N1(xlpe)"]


@pytest.mark.parametrize("semente", range(1, 5))
def test_totalizadora_com_operadores_igual_ao_js(semente):
    rnd = random.Random(2000 + semente)
    for _ in range(8):
        outros = [{"entidade": "0", "operacao": rnd.choice(["I", "*I"]), "ativo": " ".join(rnd.sample(OUTROS_OPERADORES, rnd.randint(1, 3)))} for _ in range(rnd.randint(1, 3))]
        proj = rnd.choice(["PERNAMBUCO", ""])
        js = o.totalizadora_e_payload([], outros, [], BASE_OPERADORES, proj)
        tot = tt.montar_totalizadora([], outros, [], BASE_OPERADORES, proj)
        assert tt.normalizar_para_json(tot) == js["totalizadora"], (outros, proj)
