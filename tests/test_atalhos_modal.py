"""TASK-068 — partes puras de static/atalhos_modal.js e do colar dos modais (static/grade_tabela.js)."""
import json
import os
import subprocess

import pytest

import config
from tests import oraculo_tela as o

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado")
AT = os.path.join(config.STATIC_DIR, "atalhos_modal.js")
GR = os.path.join(config.STATIC_DIR, "grade_tabela.js")


def js(expr, **entrada):
    codigo = f"const A=require({json.dumps(AT)});const G=require({json.dumps(GR)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({expr}));"
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def test_fabrica_de_cabos_e_postes():
    m = js("A.mesclar('cabos', {})")
    assert [x["valor"] for x in m["grupos"]["ativos"]][:3] == ["CAA 2", "CAA 1/0", "CAA 4/0"]
    assert [x["valor"] for x in m["grupos"]["fases"]] == ["ABC", "A", "B", "C", "AC"] and len(m["grupos"]["comps"]) == 7
    p = js("A.mesclar('postes', {})")
    assert "DT11/300" in [x["valor"] for x in p["grupos"]["postes"]] and len(p["grupos"]["mt"]) == 16 and len(p["grupos"]["bt"]) == 6


def test_mescla_sem_duplicar_e_marca_a_origem():
    fontes = {"projeto": {"itens": [{"id": "p1", "grupo": "ativos", "valor": "caa 2", "comp": 30}, {"id": "p2", "grupo": "ativos", "valor": "P 70", "comp": 80}], "favoritos": []},
              "local": {"itens": [{"id": "l1", "grupo": "ativos", "valor": "P 70"}, {"id": "l2", "grupo": "ativos", "valor": "XYZ"}], "favoritos": []}}
    m = js("A.mesclar('cabos', e.f)", f=fontes)
    ativos = m["grupos"]["ativos"]
    valores = [x["valor"] for x in ativos]
    assert valores.count("CAA 2") == 1 and "caa 2" not in valores            # a de fábrica vence a repetida do projeto
    assert next(x for x in ativos if x["valor"] == "P 70")["origem"] == "projeto"
    assert next(x for x in ativos if x["valor"] == "XYZ")["origem"] == "local"


def test_ocultos_somem_e_voltam_no_modo_edicao():
    fontes = {"ocultos": ["f:ativos:CAA 2"], "local": {"itens": [], "favoritos": []}}
    assert "CAA 2" not in [x["valor"] for x in js("A.mesclar('cabos', e.f)", f=fontes)["grupos"]["ativos"]]
    m = js("A.mesclar('cabos', e.f, true)", f=fontes)
    assert next(x for x in m["grupos"]["ativos"] if x["valor"] == "CAA 2")["oculto"] is True and m["ocultosCount"] == 1


def test_favoritos_vem_de_projeto_e_local():
    fontes = {"projeto": {"itens": [], "favoritos": [{"id": "f1", "linha": {"op": "R", "ativo": "CAA 2", "fase": "ABC", "comp": 80}}]},
              "local": {"itens": [], "favoritos": [{"id": "f2", "linha": {"op": "I", "ativo": "P 50", "fase": "A", "comp": 40}, "nome": "Meu"}]}}
    m = js("A.mesclar('cabos', e.f)", f=fontes)
    assert [x["origem"] for x in m["favoritos"]] == ["projeto", "local"]
    assert js("A.rotuloFavorito('cabos', e.f)", f=fontes["projeto"]["favoritos"][0]) == "R · CAA 2 · ABC · 80 m"
    assert js("A.rotuloFavorito('cabos', e.f)", f=fontes["local"]["favoritos"][0]) == "Meu"


def test_comprimento_padrao():
    assert js("[A.compPadrao('M3x1x35+35'), A.compPadrao('Multiplex 70'), A.compPadrao('CAA 2'), A.compPadrao('mtx'), A.compPadrao('')]") == [80, 40, 80, 40, 80]


def test_ids_unicos():
    assert len(set(js("Array.from({length:200},()=>A.novoId())"))) == 200


COLS = [{"campo": "op", "tipo": "op"}, {"campo": "ativo", "tipo": "ativo"}, {"campo": "fase", "tipo": "fase"}, {"campo": "comp", "tipo": "comp"}, {"campo": "qtd", "tipo": "qtd"}]


def colar(linhas, texto, linha=0, coluna=0, cols=COLS):
    return js("(()=>{const l=e.l;const r=G.aplicarColagemModal(l,e.t,{colunas:e.c,linha:e.i,coluna:e.j,nova:p=>({op:p?p.op:'I',ativo:'',fase:'ABC',comp:80,qtd:1})});return {l,r};})()",
              l=linhas, t=texto, i=linha, j=coluna, c=cols)


def L(op="I", ativo="", fase="ABC", comp=80, qtd=1):
    return {"op": op, "ativo": ativo, "fase": fase, "comp": comp, "qtd": qtd}


def test_colar_cinco_colunas_cria_linhas():
    r = colar([L()], "R\tCAA 2\tA\t40\t3\n*M\tP 50\tBC\t100,5\t1")
    a, b = r["l"]
    assert (a["op"], a["ativo"], a["fase"], a["comp"], a["qtd"]) == ("R", "CAA 2", "A", 40, 3)
    assert (b["op"], b["ativo"], b["fase"], b["comp"], b["qtd"]) == ("*M", "P 50", "ABC", 100.5, 1)     # BC não está na lista de fases: ignorada
    assert r["r"]["invalidos"] == 1 and r["r"]["criadas"] == 1 and a["compManual"] is True


def test_coluna_unica_com_ativo_montado_separa_tudo():
    r = colar([L("R")], "CAA 2 ABC 80 m\nM3x1x70+70 A 40 m\n35 ABC 10,5 m")
    assert [(x["ativo"], x["fase"], x["comp"]) for x in r["l"]] == [("CAA 2", "ABC", 80), ("M3x1x70+70", "A", 40), ("35", "ABC", 10.5)]
    assert r["l"][0]["op"] == "R" and r["r"]["montado"] is True


def test_operacao_mais_ativo_com_cursor_no_ativo():
    r = colar([L(), L()], "R\tP 50\n*M\tP 70", coluna=1)
    assert [(x["op"], x["ativo"]) for x in r["l"]] == [("R", "P 50"), ("*M", "P 70")]


def test_lista_de_ativos_colada_sobre_a_operacao_vai_para_o_ativo():
    r = colar([L()], "CAA 2\nCAA 4/0", coluna=0)
    assert [x["ativo"] for x in r["l"]] == ["CAA 2", "CAA 4/0"]


def test_colar_no_meio_sobrescreve():
    r = colar([L(ativo="A"), L(ativo="B"), L(ativo="C")], "X\nY", linha=1, coluna=1)
    assert [x["ativo"] for x in r["l"]] == ["A", "X", "Y"]


def test_intervalo_em_tsv():
    t = js("G.textoDoIntervaloModal(e.l, e.c, 0, 0, 1, 3)", l=[L("R", "CAA 2", "A", 40), L("I", "P 50", "ABC", 80)], c=COLS)
    assert t == "R\tCAA 2\tA\t40\nI\tP 50\tABC\t80"
