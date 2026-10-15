"""TASK-070 — lógica pura da visão SAP (static/orcamento_sap.js)."""
import json
import os
import subprocess

import pytest

import config
from tests import oraculo_tela as o

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado")
MOD = os.path.join(config.STATIC_DIR, "orcamento_sap.js")


def js(expr, **entrada):
    codigo = f"const S=require({json.dumps(MOD)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({expr}));"
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def L(op="I", mdo="MATERIAL", cod="1", desc="", qtd="1,00", filtro="", uid="u"):
    return {"uid": uid, "operacao": op, "mdo": mdo, "codigo": cod, "desc": desc, "qtd": qtd, "filtro": filtro}


def test_rotulos_m_e_n():
    assert js("['MATERIAL','material','M','MAO-DE-OBRA','MÃO-DE-OBRA','','-','N','S','OUTRO'].map(S.rotuloMdo)") == ["M", "M", "M", "N", "N", "N", "N", "N", "N", "OUTRO"]


def test_digitado_volta_para_o_valor_do_rec():
    assert js("['m','M','material','n','N','mao-de-obra','x','',' m '].map(S.mdoDeRotulo)") == ["MATERIAL", "MATERIAL", "MATERIAL", "MAO-DE-OBRA", "MAO-DE-OBRA", "MAO-DE-OBRA", None, None, "MATERIAL"]


def test_separa_odi_e_odd_so_r_e_remocao():
    linhas = [L("I", cod="a"), L("R", cod="b"), L("i", cod="c"), L("R", cod="d"), L("", cod="e")]
    r = js("S.separar(e.l)", l=linhas)
    assert [x["codigo"] for x in r["odi"]] == ["a", "c", "e"] and [x["codigo"] for x in r["odd"]] == ["b", "d"]


def test_ordenacao_texto_e_numerica_estavel():
    linhas = [L(cod="b", qtd="10,00"), L(cod="a", qtd="2,50"), L(cod="c", qtd="1.000,00"), L(cod="a", qtd="3,00")]
    assert [x["codigo"] for x in js("S.ordenar(e.l,'codigo',true)", l=linhas)] == ["a", "a", "b", "c"]
    assert [x["qtd"] for x in js("S.ordenar(e.l,'codigo',true)", l=linhas)][:2] == ["2,50", "3,00"]            # estável
    assert [x["qtd"] for x in js("S.ordenar(e.l,'qtd',true)", l=linhas)] == ["2,50", "3,00", "10,00", "1.000,00"]
    assert [x["qtd"] for x in js("S.ordenar(e.l,'qtd',false)", l=linhas)] == ["1.000,00", "10,00", "3,00", "2,50"]


def test_filtros_global_coluna_e_filtro_exato():
    l = L(mdo="MAO-DE-OBRA", cod="12345", desc="INSTALAR POSTE", filtro="Poste")
    f = lambda **k: js("S.passaFiltro(e.l, e.f)", l=l, f=k)
    assert f() is True and f(**{"global": "poste"}) is True and f(**{"global": "zzz"}) is False
    assert f(mdo="N") is True and f(mdo="mao") is True and f(mdo="M") is False                # M ≠ MAO-DE-OBRA
    assert f(codigo="234") is True and f(codigo="9") is False
    assert f(desc="instalar") is True and f(desc="remover") is False
    assert f(filtro="POSTE") is True and f(filtro="Pos") is False                            # igualdade exata


def test_tsv_sem_cabecalho_com_m_e_n():
    t = js("S.tsv(e.l)", l=[L(mdo="MATERIAL", cod="1", desc="A", qtd="2,00"), L(mdo="MAO-DE-OBRA", cod="2", desc="B", qtd="3,50")])
    assert t == "M\t1\tA\t2,00\nN\t2\tB\t3,50"


def test_para_excel_com_quantidade_numerica():
    r = js("S.paraExcel(e.l)", l=[L(mdo="MAO-DE-OBRA", cod="0007", desc="x", qtd="1.234,50")])
    assert r == [{"mdo": "N", "codigo": "0007", "desc": "x", "qtd": 1234.5}]


def test_valores_unicos():
    assert js("S.valoresUnicos(e.l,'filtro')", l=[L(filtro="B"), L(filtro="A"), L(filtro=""), L(filtro="B")]) == ["A", "B"]
