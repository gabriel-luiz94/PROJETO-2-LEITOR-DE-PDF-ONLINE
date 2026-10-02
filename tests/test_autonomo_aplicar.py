"""TASK-031 fase B — o aplicador de diff do backend é IGUAL ao da tela (aplicarOperacoesAjuste real, no Node)."""
import copy
import json
import random

import pytest

from services.ajustes_planilhas import ajustar
from services.autonomo import aplicar as ap
from services.autonomo.leitor_js import LeitorJS
from tests import oraculo_tela as o
from tests.test_autonomo_leitor import CLS

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado (oráculo de paridade)")

ATIVOS_OUTROS = ["DT11/300 1-U4", "DT11/300 1-CFU 1-SUP-L", "1-AA   1-BB", "", "DT10/200 1-U3 1-U3", "1-CFUU", "CV10/300 1-N1 2-ESCORA",
                 "1-TR3 1-PR15", "1-AF", "DT11/300 1-CFU", "1-SUPL 1-CFU", "1-RA2", "1-BASE 1-RECAL"]
ATIVOS_CABOS = ["CAA2 ABC 35 m", "P 50", "CA 4 AB 10 m", "CAA2 A 5 m", "", "CU 16", "P 70 m"]
GRUPOS = {}

ACOES = [
    {"acao": "substituir", "tabela": "outros", "de": "SUP-L", "para": "SUPL"},
    {"acao": "substituir", "tabela": "outros", "de": "1-CFUU", "para": "1-CFU", "modo": "item"},
    {"acao": "normalizar", "tabela": "outros", "regras": ["espacos", "poste"]},
    {"acao": "ordenar", "tabela": "outros", "por": [{"coluna": "operacao", "valores": ["I", "*I", "R", "*R", "M", "*M"]}, {"coluna": "ativo"}]},
    {"acao": "ordenar", "tabela": "cabos", "por": [{"coluna": "ativo", "ordem": "desc"}]},
    {"acao": "excluir_linhas", "tabela": "ambos", "onde": {"vazias": True}},
    {"acao": "excluir_linhas", "tabela": "outros", "onde": {"duplicadas": True}},
    {"acao": "adicionar_ativo", "tabela": "outros", "ativo": "SUPL", "qtd": 1, "quando": {"todos": [{"tem": "CFU"}, {"nao": {"tem": "SUPL"}}]}},
    {"acao": "adicionar_ativo", "tabela": "outros", "ativo": "ESCORA", "qtd": 2, "se_ja_existe": "ignorar",
     "quando": {"todos": [{"algum": [{"tem": "U4"}, {"tem": "U3"}]}, {"texto": "^\\s*(\\d+-)?DT"}]}},
    {"acao": "remover_ativo", "tabela": "outros", "ativo": "PR15"},
    {"acao": "mesclar_duplicadas", "tabela": "outros"},
    {"acao": "adicionar_linha", "tabela": "outros", "valores": {"operacao": "I", "ativo": "1-RA2", "entidade": "0"}, "posicao": "fim", "apenas_se_nao_existir": True},
    {"acao": "adicionar_linha", "tabela": "outros", "valores": {"operacao": "I", "ativo": "1-CFU"}, "posicao": "inicio"},
    {"acao": "adicionar_linha", "tabela": "cabos", "valores": {"operacao": "I", "ativo": "CAA2 ABC 10 m", "entidade": "CABO"}, "posicao": "fim"},
]


def tabelas_aleatorias(rnd):
    outros = [{"entidade": rnd.choice(["0", "ESTRUTURA", "POSTE", "CHAVE"]), "operacao": rnd.choice(["I", "R", "M", "*I"]), "ativo": rnd.choice(ATIVOS_OUTROS)}
              for _ in range(rnd.randint(0, 9))]
    cabos = [{"entidade": rnd.choice(["CABO", "0"]), "operacao": rnd.choice(["I", "R", "M"]), "ativo": rnd.choice(ATIVOS_CABOS)}
             for _ in range(rnd.randint(0, 6))]
    return cabos, outros


@pytest.fixture(scope="module")
def leitor():
    lt = LeitorJS()
    yield lt
    lt.fechar()


def igual_ao_js(leitor, cabos, outros, operacoes, escolhidas):
    esperado = o.aplicar_diff(copy.deepcopy(cabos), copy.deepcopy(outros), operacoes, None if escolhidas is None else sorted(escolhidas), CLS)
    obtido = ap.aplicar_operacoes(cabos, outros, operacoes, escolhidas, CLS, leitor)
    assert obtido["cabos"] == esperado["cabos"]
    assert obtido["outros"] == esperado["outros"]
    assert (obtido["aplicadas"], obtido["ignoradas"]) == (esperado["aplicadas"], esperado["ignoradas"])
    return obtido


@pytest.mark.parametrize("semente", range(1, 9))
def test_aplicar_igual_ao_js_com_acoes_aleatorias(leitor, semente):
    rnd = random.Random(semente)
    total_ops = 0
    for _ in range(25):
        cabos, outros = tabelas_aleatorias(rnd)
        acoes = rnd.sample(ACOES, rnd.randint(1, 5))
        diff = ajustar(acoes, cabos, outros, GRUPOS)
        ops = diff["operacoes"]
        total_ops += len(ops)
        igual_ao_js(leitor, cabos, outros, ops, None)                                  # tudo
        if ops:
            escolhidas = {i for i in range(len(ops)) if rnd.random() < 0.5}
            igual_ao_js(leitor, cabos, outros, ops, escolhidas)                        # subconjunto
    assert total_ops > 20     # a comparação não é vazia


def test_linha_alterada_desde_a_previa_e_ignorada(leitor):
    cabos, outros = [], [{"entidade": "ESTRUTURA", "operacao": "I", "ativo": "1-U4"}]
    diff = ajustar([{"acao": "substituir", "tabela": "outros", "de": "U4", "para": "U3"}], cabos, outros, GRUPOS)
    mudou = [{"entidade": "ESTRUTURA", "operacao": "I", "ativo": "1-OUTRA"}]     # a linha mudou depois da prévia
    r = igual_ao_js(leitor, cabos, mudou, diff["operacoes"], None)
    assert r["ignoradas"] == 1 and r["aplicadas"] == 0 and r["outros"] == mudou


def test_aplicar_tudo_chega_no_resultado_do_motor(leitor):
    """O diff aplicado reproduz as tabelas finais que o próprio `ajustar` calculou (campos visíveis)."""
    rnd = random.Random(99)
    for _ in range(30):
        cabos, outros = tabelas_aleatorias(rnd)
        diff = ajustar(rnd.sample(ACOES, 3), cabos, outros, GRUPOS)
        r = ap.aplicar_operacoes(cabos, outros, diff["operacoes"], None, CLS, leitor)
        assert [(x["operacao"], x["ativo"]) for x in r["outros"]] == [(x["operacao"], x["ativo"]) for x in diff["outros"]]
        assert [(x["operacao"], x["ativo"]) for x in r["cabos"]] == [(x["operacao"], x["ativo"]) for x in diff["cabos"]]


def test_nao_altera_os_argumentos(leitor):
    cabos, outros = [{"entidade": "CABO", "operacao": "I", "ativo": "P 50"}], [{"entidade": "0", "operacao": "I", "ativo": "SUP-L"}]
    antes = json.dumps([cabos, outros])
    diff = ajustar([{"acao": "substituir", "tabela": "outros", "de": "SUP-L", "para": "SUPL"}], cabos, outros, GRUPOS)
    ap.aplicar_operacoes(cabos, outros, diff["operacoes"], None, CLS, leitor)
    assert json.dumps([cabos, outros]) == antes


def test_casos_de_borda_montados_a_mao_iguais_ao_js(leitor):
    outros = [{"entidade": "ESTRUTURA", "operacao": "I", "ativo": "1-B"}, {"entidade": "ESTRUTURA", "operacao": "R", "ativo": "1-A"},
              {"entidade": "ESTRUTURA", "operacao": "M", "ativo": "1-C"}]
    vis = lambda op, at: {"operacao": op, "ativo": at}  # noqa: E731
    # 1) a operação da linha mudou desde a prévia ('antes' diz I, a linha é R) => ignorada
    r = igual_ao_js(leitor, [], outros, [{"op": "editar", "tabela": "outros", "linha_id": "OUTROS-1", "antes": vis("I", "1-A"), "depois": vis("I", "1-Z")}], None)
    assert r["ignoradas"] == 1
    # 2) inserir 'depois de' uma linha que não existe => vai para o fim; sem 'depois_de' => início
    r = igual_ao_js(leitor, [], outros, [
        {"op": "inserir", "tabela": "outros", "linha_id": "OUTROS-N1", "depois": vis("I", "1-FIM"), "depois_de": "OUTROS-99"},
        {"op": "inserir", "tabela": "outros", "linha_id": "OUTROS-N2", "depois": vis("I", "1-INI"), "depois_de": None}], None)
    assert [x["ativo"] for x in r["outros"]] == ["1-INI", "1-B", "1-A", "1-C", "1-FIM"]
    # 3) reordenar com linha nova (fora da ordem) no meio: ela acompanha a linha anterior
    r = igual_ao_js(leitor, [], outros, [
        {"op": "inserir", "tabela": "outros", "linha_id": "OUTROS-N1", "depois": vis("I", "1-NOVA"), "depois_de": "OUTROS-0"},
        {"op": "mover", "tabela": "outros", "ordem_antes": ["OUTROS-0", "OUTROS-1", "OUTROS-2"], "ordem_depois": ["OUTROS-2", "OUTROS-0", "OUTROS-1"]}], None)
    assert [x["ativo"] for x in r["outros"]] == ["1-C", "1-B", "1-NOVA", "1-A"]
    # 4) excluir linha que já não existe (ids deslocados) => ignorada
    r = igual_ao_js(leitor, [], outros, [{"op": "excluir", "tabela": "outros", "linha_id": "OUTROS-7", "antes": vis("I", "1-B")}], None)
    assert r["ignoradas"] == 1 and len(r["outros"]) == 3
