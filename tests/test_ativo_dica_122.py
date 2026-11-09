"""TASK-122: dica de parâmetros ao digitar "(" na tabela Outros (static/ativo_dica.js) — parte pura em Node + ligação na tela."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
pytestmark = pytest.mark.skipif(not NODE, reason="node indisponível")
MODULO = os.path.join(config.STATIC_DIR, "ativo_dica.js")


def _js(corpo, entrada=None):
    codigo = f"const D=require({json.dumps(MODULO)});const e=JSON.parse(require('fs').readFileSync(0,'utf8')||'null');console.log(JSON.stringify({corpo}));"
    r = subprocess.run([NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def test_contexto_do_parentese():
    casos = [
        ("1-IT_2(", None), ("1-IT_2(D,", None), ("9400D, N1(XL", None), ("2 IT-2(poste=", None), ("[CAZ7,94](", None), ("1-N1(A) 2-N4(", None),
        ("1-N1(A)", None), ("1-N1", None), ("*3-N1(CC, A", None), ("N1(a(", None), ("", None), ("(", None),
    ]
    r = _js("e.map(t => D.contextoDoParentese(t[0], t[1]))", casos)
    g = lambda i: {k: r[i][k] for k in ("nome", "indice", "nomeParam", "parcial")} if r[i] else None
    assert g(0) == {"nome": "IT_2", "indice": 0, "nomeParam": None, "parcial": ""}
    assert g(1) == {"nome": "IT_2", "indice": 1, "nomeParam": None, "parcial": ""}
    assert g(2) == {"nome": "N1", "indice": 0, "nomeParam": None, "parcial": "XL"} and r[2]["inicioParcial"] == len("9400D, N1(")
    assert g(3) == {"nome": "IT-2", "indice": 0, "nomeParam": "POSTE", "parcial": ""}          # apelido com hífen; nome do parâmetro digitado
    assert g(4)["nome"] == "CAZ7,94"                                                          # literal entre colchetes
    assert g(5)["nome"] == "N4" and g(6) is None and g(7) is None                              # o parêntese anterior já fechou; sem "(" não há dica
    assert g(8) == {"nome": "N1", "indice": 1, "nomeParam": None, "parcial": "A"}              # parcial sem o espaço; "*" e quantidade saem do nome
    assert g(9)["nome"] == "A" and g(10) is None and g(11)["nome"] == ""


BASE = [
    {"ativo": "IT_2=IT_2(poste=D,tipocabo=A,cabobt=35), IT-2", "desc_ativo": "IT-2", "componente": "", "codigo": "", "projeto": "PERNAMBUCO", "origem": ""},
    {"ativo": "IT_2", "desc_ativo": "IT-2", "componente": "IT_2", "codigo": "3417025", "projeto": "PERNAMBUCO", "origem": ""},
    {"ativo": "IT_2(poste=%D)", "desc_ativo": "FLABIT", "componente": "IT_2.D", "codigo": "3480305", "projeto": "PERNAMBUCO", "origem": ""},
    {"ativo": "IT_2(poste=R)", "desc_ativo": "FLABIT", "componente": "IT_2.R", "codigo": "3480270", "projeto": "PERNAMBUCO", "origem": ""},
    {"ativo": "CFU", "desc_ativo": "CHAVE", "componente": "CFU", "codigo": "200", "projeto": "PERNAMBUCO", "origem": "OUTROS"},
    {"ativo": "P50", "desc_ativo": "CABO", "componente": "C1", "codigo": "100", "projeto": "PERNAMBUCO", "origem": "CABOS"},
    {"ativo": "CERCA=CERCA(descr=5F)", "desc_ativo": "", "componente": "", "codigo": "", "projeto": "PERNAMBUCO", "origem": ""},
    {"ativo": "CERCA", "desc_ativo": "", "componente": "CERCA", "codigo": "k", "projeto": "PERNAMBUCO", "origem": "", "formula_i": "=NUM(fios)/8", "formula_r": ""},
]


def test_dica_parametros_sem_parametros_e_nao_encontrado():
    nomes = ["IT_2", "it-2", "CFU", "3417025", "FLABIT", "P50", "XYZ", "IT_2.D", "CERCA"]
    r = _js("(()=>{const c=D.construirContexto(e.base,'PERNAMBUCO');return e.nomes.map(n=>D.dica(n,c));})()", {"base": BASE, "nomes": nomes})
    it2 = r[0]
    assert it2["tipo"] == "parametros" and it2["ativo"] == "IT_2" and [p["nome"] for p in it2["params"]] == ["POSTE", "TIPOCABO", "CABOBT"]       # na ordem da declaração
    poste = it2["params"][0]
    assert poste["padrao"] == "D" and "D" in poste["valores"] and "R" in poste["valores"] and "%D" in poste["mascaras"] and poste["aberto"] is False
    assert r[1] == dict(it2)                                                       # o apelido IT-2 mostra os parâmetros do IT_2
    assert r[2]["tipo"] == "sem_parametros" and r[2]["mensagem"] == "CFU: ativo sem parâmetros declarados"
    assert r[3]["tipo"] == "sem_parametros"                                         # código também é nome válido (sem parâmetros)
    assert r[4]["tipo"] == "nao_encontrado" and r[4]["mensagem"] == "FLABIT: ativo não encontrado na base"      # só existe como DESCRIÇÃO: não vale
    assert r[5]["tipo"] == "nao_encontrado"                                         # P50 só existe para CABOS
    assert r[6]["tipo"] == "nao_encontrado" and r[7]["tipo"] == "sem_parametros"   # componente é nome válido
    cerca = r[8]
    assert cerca["tipo"] == "parametros" and [(p["nome"], p["porNome"]) for p in cerca["params"]] == [("DESCR", False), ("FIOS", True)]       # "fios" só na fórmula: por nome


def test_dica_respeita_o_projeto():
    r = _js("(()=>{const c=D.construirContexto(e.base,'RONDONIA');return D.dica('IT_2',c);})()", {"base": BASE})
    assert r["tipo"] == "nao_encontrado"                                           # a base é de PERNAMBUCO


def test_ligacao_na_tela():
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    assert html.index("/static/ativo_parametros.js") < html.index("/static/ativo_dica.js") < html.index("/static/resumo.js")
    js = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    assert "window.AtivoDica.iniciar" in js and "carregarBase" in js
    dica = open(MODULO, encoding="utf-8").read()
    assert "#body-outros .inp-ativo, #body-outros-vazio .inp-ativo-vazio" in dica          # só a tabela Outros
