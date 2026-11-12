"""TASK-146 — pesquisa e ficha do ativo na Base de Orçamento (static/base_ficha_dados.js, parte pura) e as ligações na tela."""
import json
import os
import subprocess

import pytest

import config
from tests import oraculo_tela as o

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado")
MOD = os.path.join(config.STATIC_DIR, "base_ficha_dados.js")


def L(ativo, comp="", cod="", desc="", fi=1, mdo="MATERIAL", proj="PE", **k):
    r = {"ativo": ativo, "desc_ativo": "d", "componente": comp, "projeto": proj, "mdo": mdo, "codigo": cod, "desc_codigo": desc, "fator_i": fi, "fator_r": 0, "filtro": "F", "lv": "A", "arred": "NORMAL",
         "formula_i": "", "formula_r": ""}
    r.update(k)
    return r


BASE = [
    L("CHAVE_FACA=CHAVE_FACA(cabo=S04,cabo2=S04,con=S), SUH"), L("CHAVE_FACA", "CF", "500109", "CH SEC", 3), L("", "CF", "5040005", "GRAMPO", 1),
    L("CHAVE_FACA(cabo=%S336%)", "CF.C", "2420373", "BARRA", 3),
    L("CHAVE_FACA(cabo=%,cabo2=%,con=S|N):=3*TERM(cabo) + 3*TERM(cabo2)"), L("CHAVE_FACA(cabo=%,cabo2=%,con=S):=6*CON(cabo,cabo2)"),
    L("TERM=TERM(cabo)"), L("TERM(cabo=%S04%)", "T", "2418089", "TERMINAL", 1), L("KIT:=2*CHAVE_FACA(S04,S04)"),
    L("OUTRO", "O", "500109", "CH SEC DIFERENTE", 1), L("PB_ONLY", "P", "9", "x", 1, proj="PARAIBA"), L("COMUM", "C", "5040005", "GRAMPO", 1, proj=""),
]


def js(expr, projeto="PE"):
    codigo = (f"const F=require({json.dumps(MOD)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));const m=F.construir(e,{json.dumps(projeto)});"
              f"console.log(JSON.stringify({expr}));")
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(BASE), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def test_pesquisa_ativo_apelido_codigo_e_projeto():
    r = js("F.pesquisar(m,'suh').map(x=>[x.tipo,x.nome,x.via])")
    assert r == [["ativo", "CHAVE_FACA", "SUH"]]
    assert js("F.pesquisar(m,'chave')[0].nome") == "CHAVE_FACA"
    cod = js("F.pesquisar(m,'5040')[0]")
    assert cod["tipo"] == "codigo" and cod["codigo"] == "5040005" and sorted(cod["em"]) == ["CHAVE_FACA", "COMUM"]
    assert js("F.pesquisar(m,'PB_ONLY')") == []                                  # ativo de OUTRO projeto não aparece
    assert [x["nome"] for x in js("F.pesquisar(m,'PB_ONLY')", "PARAIBA")] == ["PB_ONLY"]
    assert js("F.pesquisar(m,'')") == []


def test_ficha_do_ativo():
    f = js("F.ficha(m,'suh')")                                                   # pelo apelido
    assert f["nome"] == "CHAVE_FACA" and f["apelidos"] == ["SUH"]
    assert [p["nome"] for p in f["parametros"]] == ["CABO", "CABO2", "CON"] and f["parametros"][0]["padrao"] == "S04"
    assert sorted(f["parametros"][0]["valores"]) == ["%", "%S336%"] and sorted(f["parametros"][2]["valores"]) == ["S", "S|N"]
    assert [l["codigo"] for l in f["comuns"]] == ["500109", "5040005"]           # a linha de ATIVO em branco herda o ativo de cima
    assert [c["nome"] for c in f["chama"]] == ["TERM", "TERM", "CON"] and [c["existe"] for c in f["chama"]] == [True, True, False]
    assert [c["ativo"] for c in f["chamadoPor"]] == ["KIT"]
    assert any(v["nivel"] == "erro" and "CON" in v["texto"] for v in f["verificacoes"])
    assert js("F.ficha(m,'TERM').chamadoPor.map(c=>c.ativo)") == ["CHAVE_FACA", "CHAVE_FACA"]
    assert js("F.ficha(m,'NAOEXISTE')") is None


def test_verificacao_de_descricao_divergente_e_uso_do_codigo():
    f = js("F.ficha(m,'OUTRO')")
    assert any("500109" in v["texto"] and "descrições diferentes" in v["texto"] for v in f["verificacoes"])
    u = js("F.usoDoCodigo(m,'500109')")
    assert sorted(x["ativo"] for x in u["usos"]) == ["CHAVE_FACA", "OUTRO"] and len(u["descricoes"]) == 2
    assert js("F.usoDoCodigo(m,'0000')") is None


def test_ligacoes_na_tela_da_base():
    html = open(os.path.join(config.STATIC_DIR, "orcamento.html"), encoding="utf-8").read()
    for trecho in ("/static/base_ficha_dados.js", "/static/base_ficha_ui.js", "BaseFicha.iniciar(", "BaseFicha.invalidar()", "BaseFicha.pesquisaRapida"):
        assert trecho in html, trecho
    app = open(os.path.join(config.BASE_DIR, "app.py"), encoding="utf-8").read()
    assert "base_ficha_dados.js" in app and "base_ficha_ui.js" in app
