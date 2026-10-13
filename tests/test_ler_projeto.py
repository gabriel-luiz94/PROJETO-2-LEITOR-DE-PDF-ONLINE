"""TASK-062 — static/ler_projeto.js (botão "Ler projeto" do Resumo) dá o MESMO resultado que Leitor -> Processar -> Resumo (JS real no Node)."""
import json
import os
import subprocess

import pytest

import config
from tests import oraculo_tela as o
from tests.test_autonomo_leitor import CLS, PROC, itens_aleatorios

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado (oráculo de paridade)")

MODULO = os.path.join(config.STATIC_DIR, "ler_projeto.js") if hasattr(config, "STATIC_DIR") else os.path.join(os.path.dirname(o.MOTOR), "ler_projeto.js")


def ler_projeto(itens, proc, cls):
    codigo = f"""
const fs = require('fs');
const e = JSON.parse(fs.readFileSync(0, 'utf8'));
const LP = require({json.dumps(MODULO)});
const Motor = require({json.dumps(o.MOTOR)});
LP.processar(e.itens, e.proc, e.cls, Motor).then(r => console.log(JSON.stringify(r)));
"""
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps({"itens": itens, "proc": proc, "cls": cls}),
                       capture_output=True, text=True, timeout=120)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def tri(linhas):
    return [(l["entidade"], l["operacao"], l["ativo"]) for l in linhas]


@pytest.mark.parametrize("semente", [31, 32, 33, 34])
def test_igual_ao_caminho_leitor_processar_resumo(semente):
    itens = itens_aleatorios(250, semente)
    exportados = o.passo_leitor(itens, PROC, CLS)
    esperado = o.passo_resumo(exportados, PROC, CLS)
    obtido = ler_projeto(itens, PROC, CLS)
    assert [json.dumps(x, sort_keys=True) for x in obtido["exportados"]] == [json.dumps(x, sort_keys=True) for x in exportados]
    assert tri(obtido["cabos"]) == tri(esperado["cabos"])
    assert tri(obtido["outros"]) == tri(esperado["outros"])
    assert [{"entidade": r["entidade"], "texto": r["texto"], "pagina": r["pagina"]} for r in obtido["ramais"]] == esperado["ramais"]


def test_amostra_cobre_cabos_outros_e_ramais():
    r = ler_projeto(itens_aleatorios(600, 21), PROC, CLS)
    assert len(r["cabos"]) > 5 and len(r["outros"]) > 5 and len(r["ramais"]) > 0


def test_sem_itens_gera_tabelas_vazias():
    r = ler_projeto([], PROC, CLS)
    assert r == {"exportados": [], "cabos": [], "outros": [], "ramais": []}


def test_coordenada_passa_para_as_linhas():
    itens = [{"pagina": 1, "texto": "CAA 2", "cor": "#000000", "layer": "", "_x": 1.5, "_y": 2.5}]
    r = ler_projeto(itens, PROC, CLS)
    assert r["exportados"][0]["_x"] == 1.5 and r["exportados"][0]["_y"] == 2.5


def test_extensoes():
    codigo = f"const L=require({json.dumps(MODULO)});console.log(JSON.stringify(['a.PDF','b.dxf','c.dwg','d.txt','',null].map(L.extensaoValida)));"
    r = subprocess.run([o.NODE, "-e", codigo], capture_output=True, text=True)
    assert json.loads(r.stdout) == [True, True, True, False, False, False]
