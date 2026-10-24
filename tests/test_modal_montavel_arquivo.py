"""TASK-077 etapa 6 — exportar/importar modais montáveis (static/modal_montavel_arquivo.js, Node)."""
import json
import os
import subprocess

import pytest

import config
from tests import oraculo_tela as o

AR = os.path.join(config.STATIC_DIR, "modal_montavel_arquivo.js")
pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado")


def js(expr, **entrada):
    codigo = f"const A=require({json.dumps(AR)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));let r;try{{r={expr};}}catch(x){{r={{erro:x.message}}}}console.log(JSON.stringify(r));"
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


M = lambda i, n: {"id": i, "nome": n, "tipo": "montagem", "blocos": [{"tipo": "texto", "texto": "x"}]}


def test_exportar_e_ler_ida_e_volta():
    arq = js("A.exportar('229', e.m)", m=[M("a", "A")])
    assert arq["formato"] == "modais_montaveis" and arq["versao"] == 1 and arq["projeto"] == "229" and arq["exportado_em"]
    assert js("A.ler(JSON.stringify(e.a))", a=arq) == [M("a", "A")]


def test_ler_aceita_formatos_simples_e_recusa_outros():
    assert js("A.ler(JSON.stringify({versao:1, modais:[e.m]}))", m=M("a", "A")) == [M("a", "A")]
    assert js("A.ler(JSON.stringify([e.m]))", m=M("a", "A")) == [M("a", "A")]
    assert js("A.ler(JSON.stringify(e.m))", m=M("a", "A")) == [M("a", "A")]                   # um modal só
    assert "JSON válido" in js("A.ler('{nao é json')")["erro"]
    assert "outro cadastro" in js("A.ler(JSON.stringify({formato:'ajustes', ajustes:[]}))")["erro"]
    assert "não reconhecido" in js("A.ler(JSON.stringify({qualquer:1}))")["erro"]


def test_resumir_acrescentar_e_substituir():
    atuais = [M("a", "A"), M("b", "B")]
    novo_b = M("b", "B2"); novo_c = M("c", "C")
    ac = js("A.resumir(e.at, e.im, 'acrescentar')", at=atuais, im=[novo_b, novo_c])
    assert ac["novos"] == ["C"] and ac["atualizados"] == ["B2"] and ac["removidos"] == []
    assert [m["nome"] for m in ac["resultado"]] == ["A", "B2", "C"]                           # mantém a ordem, atualiza no lugar
    su = js("A.resumir(e.at, e.im, 'substituir')", at=atuais, im=[novo_b, novo_c])
    assert su["removidos"] == ["A"] and [m["id"] for m in su["resultado"]] == ["b", "c"]
    assert js("A.resumir([], [e.m], 'acrescentar')", m=M("z", "Z"))["novos"] == ["Z"]
