"""TASK-069 — ordenação por coluna (static/grade_tabela.js): natural, estável, ordem lógica das operações, vínculos remapeados."""
import json
import os
import random
import subprocess

import pytest

import config
from tests import oraculo_tela as o

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado")
GR = os.path.join(config.STATIC_DIR, "grade_tabela.js")


def js(expr, **entrada):
    codigo = f"const G=require({json.dumps(GR)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({expr}));"
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def ordenar(linhas, criterios):
    ordem = js("G.ordenarLinhas(e.l, e.c)", l=linhas, c=criterios)
    return [linhas[i] for i in ordem], ordem


def L(ativo="", op="I", ent="0", **kw):
    return {"entidade": ent, "operacao": op, "ativo": ativo, **kw}


def test_ativo_em_ordem_natural_e_vazio_por_ultimo():
    linhas = [L("10-U4"), L("2-U4"), L(""), L("1-N1"), L("2-B1")]
    r, _ = ordenar(linhas, [{"campo": "ativo", "dir": "asc"}])
    assert [x["ativo"] for x in r] == ["1-N1", "2-B1", "2-U4", "10-U4", ""]
    r, _ = ordenar(linhas, [{"campo": "ativo", "dir": "desc"}])
    assert [x["ativo"] for x in r] == ["10-U4", "2-U4", "2-B1", "1-N1", ""]          # vazio continua por último


def test_operacao_em_ordem_logica():
    linhas = [L(op=x) for x in ["*M", "M", "*R", "R", "*I", "I"]]
    r, _ = ordenar(linhas, [{"campo": "operacao", "dir": "asc"}])
    assert [x["operacao"] for x in r] == ["I", "*I", "R", "*R", "M", "*M"]


def test_entidade_alfabetica_com_zero_por_ultimo():
    linhas = [L(ent=e) for e in ["POSTE", "0", "APOIO", "ESTRUTURA", "0", "CHAVE"]]
    r, _ = ordenar(linhas, [{"campo": "entidade", "dir": "asc"}])
    assert [x["entidade"] for x in r] == ["APOIO", "CHAVE", "ESTRUTURA", "POSTE", "0", "0"]


def test_estavel_e_com_criterio_secundario():
    linhas = [L("A", "R", "X1"), L("B", "I", "X2"), L("C", "R", "X3"), L("D", "I", "X4")]
    r, _ = ordenar(linhas, [{"campo": "operacao", "dir": "asc"}])
    assert [x["ativo"] for x in r] == ["B", "D", "A", "C"]                  # empates mantêm a ordem original
    r, _ = ordenar(linhas, [{"campo": "operacao", "dir": "asc"}, {"campo": "ativo", "dir": "desc"}])
    assert [x["ativo"] for x in r] == ["D", "B", "C", "A"]


def test_qtd_ativos_numerica_com_negativos_e_strings():
    linhas = [{"qtdAtivos": "-5"}, {"qtdAtivos": 10}, {"qtdAtivos": 2}, {"qtdAtivos": ""}]
    r, _ = ordenar(linhas, [{"campo": "qtdAtivos", "dir": "asc"}])
    assert [x["qtdAtivos"] for x in r] == ["-5", 2, 10, ""]


def test_esta_ordenado():
    cr = [{"campo": "ativo", "dir": "asc"}]
    assert js("G.estaOrdenado(e.l, e.c)", l=[L("1-A"), L("2-A"), L("")], c=cr) is True
    assert js("G.estaOrdenado(e.l, e.c)", l=[L("2-A"), L("1-A")], c=cr) is False
    assert js("G.estaOrdenado(e.l, e.c)", l=[L("1-A")], c=[]) is False


@pytest.mark.parametrize("semente", range(5))
def test_vinculos_remapeados_continuam_apontando_para_a_mesma_estrutura(semente):
    rnd = random.Random(semente)
    outros = [L(f"{rnd.randint(1, 30)}-{rnd.choice(['U4', 'N1', 'B2', 'CE3'])}", rnd.choice(["I", "R", "M"]), rnd.choice(["ESTRUTURA", "APOIO", "0"])) for _ in range(40)]
    cabos = [{"ativo": f"CAA {i}", "vinculoEstruturas": rnd.sample(range(40), 2)} for i in range(15)]
    alvos = [[outros[i]["ativo"] + "|" + str(i) for i in c["vinculoEstruturas"]] for c in cabos]   # identidade da estrutura apontada
    for i, x in enumerate(outros):
        x["id"] = i
    crit = [{"campo": rnd.choice(["ativo", "operacao", "entidade"]), "dir": rnd.choice(["asc", "desc"])}]
    ordem = js("G.ordenarLinhas(e.l, e.c)", l=outros, c=crit)
    mapa = js("G.mapaDeOrdem(e.o)", o=ordem)
    novos = [outros[i] for i in ordem]
    for c, antes in zip(cabos, alvos):
        novos_idx = [mapa[i] for i in c["vinculoEstruturas"]]
        # a estrutura apontada depois é exatamente a mesma linha de antes
        assert [novos[k]["id"] for k in novos_idx] == c["vinculoEstruturas"]
