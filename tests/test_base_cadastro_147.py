"""TASK-147 — cadastro rápido de ativos: geração das linhas, leitura do que se cola do Excel, validação local (JS) e avaliação do rascunho no servidor."""
import json
import os
import subprocess

import pytest

import config
from services.rascunho_ativo import avaliar
from tests import oraculo_tela as o
from tests.test_fator_formula_109 import adm, banco, client  # noqa: F401

MOD = os.path.join(config.STATIC_DIR, "base_cadastro_dados.js")
FD = os.path.join(config.STATIC_DIR, "base_ficha_dados.js")
node = pytest.mark.skipif(not o.NODE, reason="node não instalado")


def js(expr, **e):
    codigo = (f"const C=require({json.dumps(MOD)});const F=require({json.dumps(FD)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));"
              f"console.log(JSON.stringify({expr}));")
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(e), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def It(**k):
    d = {"condicao": "", "codigo": "", "desc_codigo": "", "mdo": "M", "fator_i": "1", "fator_r": "0", "lv": "A", "filtro": "", "arred": "NORMAL"}
    d.update(k)
    return d


RASC = {"nome": "CHAVE_FACA_OLEO", "apelidos": "SUO, CFO", "descricao": "CHAVE FACA A OLEO", "parametros": [{"nome": "cabo", "padrao": "S04"}], "itens": [
    It(codigo="500109", desc_codigo="CH SEC", fator_i="3", filtro="6.CHAVE"), It(condicao="cabo=%S336%", codigo="2420373", desc_codigo="BARRA", fator_i="3"),
    It(codigo="SDEMU3602II", desc_codigo="INST", mdo="N", fator_i="1", lv="LM"), It(condicao="cabo=%", codigo=":= 3*TERM(cabo)"), It(), It(codigo="7", fator_i="=NUM(cabo)/8")]}


@node
def test_gerar_linhas_no_formato_da_base():
    ls = js("C.gerarLinhas(e.r,'PE')", r=RASC)
    assert [l["ativo"] for l in ls] == ["CHAVE_FACA_OLEO=CHAVE_FACA_OLEO(cabo=S04), SUO, CFO", "CHAVE_FACA_OLEO", "CHAVE_FACA_OLEO(cabo=%S336%)", "CHAVE_FACA_OLEO", "CHAVE_FACA_OLEO(cabo=%):=3*TERM(cabo)", "CHAVE_FACA_OLEO"]
    assert ls[1]["componente"] == "CHAVE_FACA_OLEO" and ls[2]["componente"] == "CHAVE_FACA_OLEO.CABO=%S336%" and ls[3]["mdo"] == "MAO-DE-OBRA" and ls[3]["lv"] == "LM"
    assert ls[4]["codigo"] == "" and ls[4]["fator_i"] == 0 and ls[5]["formula_i"] == "=NUM(cabo)/8" and ls[0]["projeto"] == "PE" and ls[1]["fator_i"] == 3
    sem = js("C.gerarLinhas({nome:'N1',apelidos:'',parametros:[],itens:[e.i]},'PE').map(l=>l.ativo)", i=It(codigo="1"))
    assert sem == ["N1"]                                                                  # sem parâmetros nem apelidos: nenhuma linha de declaração
    assert js("C.gerarLinhas({nome:'9400D',apelidos:'DT9/400',parametros:[],itens:[e.i]},'PE').map(l=>l.ativo)", i=It(codigo="1")) == ["9400D, DT9/400", "9400D"]


@node
def test_colar_do_excel_e_combinacoes():
    it = js("C.lerColado(e.t)", t="CODIGO\tFATOR I\n500109\t3\n7511022\t1,5\t0\tLM\n")
    assert [(x["codigo"], x["fator_i"], x["lv"]) for x in it] == [("500109", "3", "A"), ("7511022", "1,5", "LM")]
    base = js("C.lerColado(e.t)", t="CHAVE_FACA(cabo=%S336%);d;c;PE;MATERIAL;;2420373;BARRA;3;0;6.CHAVE;A;NORMAL".replace(";", "\t"))
    assert base[0]["condicao"] == "cabo=%S336%" and base[0]["codigo"] == "2420373" and base[0]["filtro"] == "6.CHAVE"
    comb = js("C.lerCombinacoes(e.t,['cabo','cabo2'])", t="S4/0\tS4/0\t2400003\nC35\tC25\t2401000\t2")
    assert comb[0]["condicao"] == "cabo=%S4/0%,cabo2=%S4/0%" and comb[1]["fator_i"] == "2"


@node
def test_validacao_local():
    base = [{"ativo": "TERM=TERM(cabo=S04)", "projeto": "PE"}, {"ativo": "TERM(cabo=%S04%)", "projeto": "PE", "codigo": "2418089", "desc_codigo": "T", "mdo": "MATERIAL", "fator_i": 1, "filtro": "F"},
            {"ativo": "OUTRO, SUO", "projeto": "PE"}]
    v = js("C.validarLocal(e.r,F.construir(e.b,'PE'))", r=RASC, b=base)
    textos = [x["texto"] for x in v["erros"]]
    assert any('apelido "SUO"' in t for t in textos)                                      # apelido já usado por OUTRO
    assert not any("TERM" in t for t in textos)                                           # TERM existe
    assert any("7" in a["texto"] and "novo na base" not in a["texto"] or "código 7 novo" in a["texto"] for a in v["avisos"])
    ruim = js("C.validarLocal({nome:'A B',apelidos:'',parametros:[],itens:[e.i,e.j,e.k]},F.construir([],'PE'))", i=It(condicao="cabo=%X%", codigo="1"), j=It(codigo="2", fator_i="abc"), k=It(condicao="x=1"))
    t = " | ".join(x["texto"] for x in ruim["erros"])
    assert "não pode ter espaços" in t and "não declara parâmetros" in t and 'Fator I "abc"' in t and "informe o código" in t


def L(a, comp="", cod="", d="", fi=0):
    return {"ativo": a, "desc_ativo": "d", "componente": comp, "projeto": "PE", "mdo": "MATERIAL" if cod else "", "codigo": cod, "desc_codigo": d, "fator_i": fi, "fator_r": 0, "filtro": "", "origem": "", "lv": "A", "arred": "NORMAL",
            "formula_i": "", "formula_r": ""}


BASE = [L("TERM=TERM(cabo=S04)"), L("TERM(cabo=%S04%)", "T", "2418089", "TERM", 1), L("VELHO=VELHO(cabo=S04)"), L("VELHO(cabo=%S04%)", "V", "1", "x", 1)]


def test_avaliar_rascunho_so_devolve_problemas_novos_e_simula():
    novas = [L("X=X(cabo=S04), XX"), L("X", "X", "500109", "CH", 3), L("X(cabo=%S336%)", "", "2420373", "BARRA", 3), L("X(cabo=%):=3*TERM(cabo)")]
    r = avaliar(BASE, novas, "PE", [{"ativo": "X(S04)"}])
    assert r["erros"] == [] and r["componentes_gerados"] == 1 and [l["componente"] for l in r["linhas"]][2] == "X.CABO=%S336%"
    sim = {x["codigo"]: x["total"] for x in r["simulacao"][0]["resultado"]}
    assert sim == {"500109": 3.0, "2418089": 3.0} and r["simulacao"][0]["nao_encontrados"] == []
    ruim = avaliar(BASE, [L("Y=Y(cabo=S04)"), L("Y(foo=%A%)", "", "1", "d", 1)], "PE")
    assert any("FOO" in e for e in ruim["erros"])
    assert avaliar(BASE, [{"ativo": "Z", "lv": "XX", "codigo": "1"}], "PE")["erros"][0].startswith("Linha nova 1")


def test_avaliar_troca_as_linhas_antigas_do_mesmo_ativo_no_projeto():
    r = avaliar(BASE, [L("VELHO=VELHO(cabo=S04)"), L("VELHO(cabo=%S04%)", "V2", "2", "novo", 5)], "PE", [{"ativo": "VELHO(S04)"}])
    assert {x["codigo"]: x["total"] for x in r["simulacao"][0]["resultado"]} == {"2": 5.0}       # a definição antiga (código 1) não entra


def test_rota_rascunho(client):
    r = client.post("/api/orcamento/rascunho", json={"linhas": [L("W", "W", "1", "d", 2)], "projeto": "PE", "simular": [{"ativo": "3-W"}]}, headers=adm())
    assert r.status_code == 200, r.text
    j = r.json()
    assert j["erros"] == [] and j["simulacao"][0]["resultado"][0]["total"] == 6.0
    assert client.post("/api/orcamento/rascunho", json={"linhas": []}).status_code in (401, 403)


def test_ligacoes_na_tela():
    html = open(os.path.join(config.STATIC_DIR, "orcamento.html"), encoding="utf-8").read()
    for t in ("/static/base_cadastro_dados.js", "/static/base_cadastro_ui.js", "BaseCadastro.abrir()", "BaseCadastro.iniciar("):
        assert t in html, t
    app = open(os.path.join(config.BASE_DIR, "app.py"), encoding="utf-8").read()
    assert "base_cadastro_dados.js" in app and "base_cadastro_ui.js" in app


@node
def test_rascunho_de_linhas_ida_e_volta_mantendo_o_que_a_linha_sem_codigo_ja_tinha():
    """TASK-148: editar um ativo existente (linhas → rascunho → linhas) não perde MDO/filtro/LV da declaração nem das chamadas."""
    rows = [{"ativo": "CF=CF(cabo=S04,cabo2=S04), SUH", "desc_ativo": "CHAVE", "mdo": "MATERIAL", "filtro": "6.CHAVE", "lv": "A", "arred": "NORMAL", "codigo": ""},
            {"ativo": "CF", "desc_ativo": "CHAVE", "mdo": "MATERIAL", "codigo": "500109", "desc_codigo": "CH", "fator_i": 3, "fator_r": 0, "lv": "A", "filtro": "6.CHAVE"},
            {"ativo": "CF(cabo=%S336%)", "mdo": "MAO-DE-OBRA", "codigo": "9", "desc_codigo": "M", "fator_i": 1, "formula_i": "=NUM(cabo)", "lv": "LM"},
            {"ativo": "CF(cabo=%,cabo2=%):=3*TERM(cabo)", "mdo": "MATERIAL", "filtro": "9.X", "lv": "LV", "codigo": ""}]
    ls = js("C.gerarLinhas(C.rascunhoDeLinhas(e.rows),'PE')", rows=rows)
    assert [l["ativo"] for l in ls] == [r["ativo"] for r in rows]
    assert (ls[0]["mdo"], ls[0]["filtro"]) == ("MATERIAL", "6.CHAVE") and (ls[3]["filtro"], ls[3]["lv"]) == ("9.X", "LV")
    assert ls[2]["formula_i"] == "=NUM(cabo)" and ls[2]["mdo"] == "MAO-DE-OBRA" and ls[2]["lv"] == "LM"
    assert js("C.rascunhoDeLinhas(e.rows).parametros.map(p=>p.nome+'='+p.padrao)", rows=rows) == ["cabo=S04", "cabo2=S04"]
