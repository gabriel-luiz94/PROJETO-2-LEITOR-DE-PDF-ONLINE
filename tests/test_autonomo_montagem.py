"""TASK-031 fase A — a montagem de Cabos/Outros em Python é IGUAL à da tela (resumo.js real executado no Node)."""
import json
import random

import pytest

from services.autonomo import montagem as mt
from tests import oraculo_tela as o
from tests.test_autonomo_leitor import CLS, PROC, itens_aleatorios

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado (oráculo de paridade)")

PEDACOS = ["CAA", "CAA2", "CAA 2", "CA", "CA 4", "CA4", "CU", "CU 16", "CAZ", "CAZ 3", "P", "P 50", "P50", "M", "M12", "M 1", "ABC", "35", "35,5",
           "3.5", "m", "M", "A", "AB", "ABC", "ABCD", "R", "S", "T", "N", "RS", "-", "/", "x", "1x", "10", "", " ", "ÇÃO", "ß", " ", "﻿",
           "CAA2 A 10 m", "CAA 2 ABC 35 m", "CA 4 AB 10", "P 50 m", "1x35 m"]


def textos_ativos(n, semente):
    rnd = random.Random(semente)
    saida = []
    for _ in range(n):
        partes = [rnd.choice(PEDACOS) for _ in range(rnd.randint(0, 5))]
        sep = rnd.choice([" ", " ", "  ", "\t", "  "])
        saida.append(sep.join(partes) + rnd.choice(["", "", " ", " m", "m", " M", "\n"]))
    return saida


def test_calculo_de_qtd_igual_ao_js_em_milhares_de_textos():
    ativos = textos_ativos(1500, 7) + ["CAA2 A 10 m", "CAA2 AB 10 m", "CA4 A 10 m", "CA 4 AB 10", "P 50", "CAZ", "M12", "m", "", "   ", None]
    casos = [{"ativo": a, "fase": f} for a in ativos for f in (None, "ABC")]
    js = o.funcoes_qtd(casos)
    for c, j in zip(casos, js):
        assert mt.calcular_qtd_ativos(c["ativo"], c["fase"]) == j["qtd"], c
        assert mt.extrair_fase(c["ativo"]) == j["fase"], c
        assert mt.linha_standalone(c["ativo"]) == j["standalone"], c


@pytest.mark.parametrize("semente", [1, 2, 3])
def test_recalculo_da_tabela_igual_ao_js(semente):
    rnd = random.Random(semente)
    for _ in range(40):
        linhas = [{"ativo": a, **({"qtdAtivos": rnd.choice([-3, "-2", 4, "5", 0])} if rnd.random() < 0.3 else {})}
                  for a in textos_ativos(rnd.randint(0, 12), rnd.randint(0, 10**6)) + ["P 50", "CAA2", "3 4 5"]]
        esperado = o.recalcular_lista(json.loads(json.dumps(linhas)))
        obtido = json.loads(json.dumps(linhas))
        mt.recalcular_qtd_ativos(obtido)
        assert obtido == esperado


@pytest.fixture(scope="module")
def leitor():
    from services.autonomo.leitor_js import LeitorJS
    lt = LeitorJS()
    yield lt
    lt.fechar()


@pytest.mark.parametrize("semente", [11, 12, 13, 14])
def test_tabelas_montadas_iguais_as_da_tela(leitor, semente):
    itens = itens_aleatorios(250, semente)
    exportados = o.passo_leitor(itens, PROC, CLS)          # script.js (JS real)
    esperado = o.passo_resumo(exportados, PROC, CLS)       # resumo.js (JS real)
    obtido = mt.montar_tabelas(itens, PROC, CLS, leitor)   # Python + mesmo motor via QuickJS
    assert obtido["cabos"] == esperado["cabos"]
    assert obtido["outros"] == esperado["outros"]
    assert obtido["ramais"] == esperado["ramais"]


def test_a_amostra_cobre_cabos_outros_e_ramais(leitor):
    """Garante que a comparação acima não é vazia: há linhas em todas as tabelas."""
    t = mt.montar_tabelas(itens_aleatorios(600, 21), PROC, CLS, leitor)
    assert len(t["cabos"]) > 5 and len(t["outros"]) > 5 and len(t["ramais"]) > 0
    assert all("qtdAtivos" in c for c in t["cabos"])


def test_sem_itens_gera_tabelas_vazias(leitor):
    assert mt.montar_tabelas([], PROC, CLS, leitor) == {"cabos": [], "outros": [], "ramais": []}


def test_exemplos_conhecidos_de_quantidade():
    q = mt.calcular_qtd_ativos
    assert q("CAA2 A 10 m") == 2          # CAA2 com fase de 1 letra: +1
    assert q("CAA2 ABC 35 m") == 3
    assert q("CA4 AB 10 m") == 3          # CA4 sempre +1
    assert q("CAZ 3 ABC 10 m") == 1 and q("M12") == 1
    assert q("P50") == 1                  # standalone sem fase herdada
    assert q("P50", "ABC") == 3           # standalone herda a fase da linha seguinte
    assert q("") == 0 and q(None) == 0


def test_linha_negativa_continua_negativa():
    linhas = [{"ativo": "CAA2 ABC 10 m", "qtdAtivos": "-9"}, {"ativo": "P 50", "qtdAtivos": 5}]
    mt.recalcular_qtd_ativos(linhas)
    assert linhas[0]["qtdAtivos"] == "-3" and linhas[1]["qtdAtivos"] == 1


class LeitorFalso:
    """Leitor de mentira para testar a LÓGICA de decisão da montagem (quem é confiado, quem roda o motor de novo)."""
    def __init__(self, passo1, passo2):
        self.passo1, self.passo2, self.chamadas = passo1, passo2, []

    def processar_lote(self, itens, proc, cls, modo="extracao", indices=None):
        self.chamadas.append((modo, indices))
        return self.passo1 if modo == "extracao" else [self.passo2[i] for i in indices]


def r(entidade, operacao, ativo):
    return {"entidade": entidade, "operacao": operacao, "ativo": ativo}


def test_resumo_confia_no_que_o_leitor_ja_definiu_e_refaz_o_resto():
    itens = [{"pagina": 1, "texto": t, "cor": "#f00", "layer": ""} for t in "ABCD"]
    passo1 = [r("CABO", "I", "CAA2 ABC 10 m"),    # confiado
              r("0", "M", "SOBRA"),               # entidade '0' => refaz mesmo com ativo
              r("ESTRUTURA", "I", ""),            # ativo vazio => refaz
              r("ESTRUTURA", "", "1-U4")]         # operação vazia => 'M'
    lt = LeitorFalso(passo1, {1: r("POSTE", "I", "DT11/300"), 2: r("RAMAIS", "R", "RS")})
    t = mt.montar_tabelas(itens, [], [], lt)
    assert lt.chamadas == [("extracao", None), ("resumo", [1, 2])]
    assert t["cabos"] == [{"entidade": "CABO", "operacao": "I", "ativo": "CAA2 ABC 10 m", "qtdAtivos": 3}]
    assert t["outros"] == [{"entidade": "POSTE", "operacao": "I", "ativo": "DT11/300"}, {"entidade": "ESTRUTURA", "operacao": "M", "ativo": "1-U4"}]
    assert t["ramais"] == [{"entidade": "RAMAIS", "texto": "C", "pagina": 1}]


def test_apoio_so_vai_para_ramais_com_recalcada_ou_base():
    itens = [{"pagina": 3, "texto": t, "cor": "#f00", "layer": ""} for t in ("REC. CALCADA 2X", "CONC BASE", "OUTRA COISA", "REC. CALÇADA")]
    lt = LeitorFalso([r("APOIO", "I", "1-RECAL")] * 4, {})
    t = mt.montar_tabelas(itens, [], [], lt)
    assert [x["texto"] for x in t["ramais"]] == ["REC. CALCADA 2X", "CONC BASE", "REC. CALÇADA"]
    assert len(t["outros"]) == 4   # APOIO também é 'outros'
