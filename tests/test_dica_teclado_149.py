"""TASK-149/150 — teclado na dica de parâmetros (Outros) e botão "⇄ Inverter" do layout do Resumo."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
MODULO = os.path.join(config.STATIC_DIR, "ativo_dica.js")
BASE = [{"ativo": "N1=N1(cabo=XLPE,poste=D)", "projeto": ""}, {"ativo": "N1(cabo=XLPE,poste=D)", "projeto": "", "componente": "a", "codigo": "1"},
        {"ativo": "N1(cabo=ALUM,poste=T)", "projeto": "", "componente": "b", "codigo": "2"}, {"ativo": "SEM_PARAM", "projeto": "", "componente": "c", "codigo": "3"}]


def passo(texto, cursor, valor):
    """O que a tela faz ao aceitar `valor`: {texto, cursor, proximos (valores do parâmetro seguinte)}."""
    codigo = (f"const D=require({json.dumps(MODULO)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));const ctx=D.construirContexto(e.b,'');"
              "const c=D.contextoDoParentese(e.t,e.c);const d=D.dica(c.nome,ctx);const p=D.planoDeInsercao(e.t,e.c,c,d,e.v);"
              "const n=e.t.slice(0,p.inicio)+p.inserir+e.t.slice(p.fim);const c2=D.contextoDoParentese(n,p.cursor);"
              "console.log(JSON.stringify({texto:n,cursor:p.cursor,valores:D.valoresDoAtual(c,d),proximos:c2?D.valoresDoAtual(c2,D.dica(c2.nome,ctx)):null}));")
    r = subprocess.run([NODE, "-e", codigo], input=json.dumps({"b": BASE, "t": texto, "c": cursor, "v": valor}), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_inserir_avanca_para_o_proximo_parametro_e_fecha_no_ultimo():
    a = passo("N1(", 3, "XLPE")
    assert (a["texto"], a["cursor"], a["valores"]) == ("N1(XLPE, ", 9, ["XLPE", "ALUM"]) and a["proximos"] == ["D", "T"]     # insere o valor e a ", " e já oferece o 2º parâmetro
    b = passo(a["texto"], a["cursor"], "D")
    assert (b["texto"], b["cursor"]) == ("N1(XLPE, D)", 11) and b["proximos"] is None                                          # último: fecha o ")" e a caixa some
    assert passo("1-N1(", 5, "ALUM")["texto"] == "1-N1(ALUM, "                                                                  # quantidade antes do nome


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_inserir_no_meio_do_texto_existente():
    assert passo("N1(XL,D)", 5, "XLPE") == {"texto": "N1(XLPE,D)", "cursor": 8, "valores": ["XLPE", "ALUM"], "proximos": ["D", "T"]}   # troca o valor e pula a "," que já existe
    assert passo("N1(XLPE,)", 8, "D")["texto"] == "N1(XLPE,D)" and passo("N1(XLPE,)", 8, "D")["cursor"] == 10                   # o ")" já existe: o cursor passa por cima
    assert passo("N1(cabo=", 8, "ALUM")["texto"] == "N1(cabo=ALUM, "                                                            # parâmetro digitado por nome


def test_teclado_ligado_na_tela_e_inverter():
    js = open(os.path.join(config.STATIC_DIR, "ativo_dica.js"), encoding="utf-8").read()
    for t in ("stopImmediatePropagation", "ArrowDown", "planoDeInsercao", "dp-val.sel"):
        assert t in js, t
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    assert 'id="btn-inverter-tabelas"' in html and "Inverter" in html
    lay = open(os.path.join(config.STATIC_DIR, "resumo_layout.js"), encoding="utf-8").read()
    assert "inverter" in lay and "invertidas" in lay and "button[data-layout]" in lay
    css = open(os.path.join(config.STATIC_DIR, "resumo_layout.css"), encoding="utf-8").read()
    assert ".tables-row.invertidas #section-outros { order: 1; }" in css


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_mascaras_simples_viram_valores_selecionaveis():
    base = [{"ativo": "TERM=TERM(cabo=S04)", "projeto": ""}, {"ativo": "TERM(cabo=%S04%)", "projeto": "", "componente": "a", "codigo": "1"}, {"ativo": "TERM(cabo=%S336%)", "projeto": "", "componente": "b", "codigo": "2"},
            {"ativo": "TERM(cabo=!%A35%)", "projeto": "", "componente": "c", "codigo": "3"}]
    codigo = (f"const D=require({json.dumps(MODULO)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));const d=D.dica('TERM',D.construirContexto(e,''));console.log(JSON.stringify(d.params[0].valores));")
    r = subprocess.run([NODE, "-e", codigo], input=json.dumps(base), capture_output=True, text=True, timeout=30)
    assert json.loads(r.stdout) == ["S04", "S336"]                                           # o padrão e as máscaras %X%; a negação (!) não vira valor
