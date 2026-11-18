"""TASK-182 C — a Operação pode ser DIGITADA (com a lista ao lado) e é validada como uma célula do Excel: maiúsculas, sem espaços, só I, *I, R, *R, M, *M."""
import json
import os
import re
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
JS = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()


def test_uma_fonte_unica_da_lista_e_nenhum_select_de_operacao_sobrou():
    assert JS.count("const OPERACOES = [") == 1
    assert "criarCampoOperacao(" in JS and JS.count("criarCampoOperacao(") >= 3                  # tabelas, linha em branco e a definição
    assert "document.createElement('select');\n            selOp.className = 'sel-operacao'" not in JS
    assert 'class="tot-input" style="text-align: center;" onchange="updateTotRow(${index}, \'operacao\'' not in JS      # a Totalizadora também é campo de texto
    assert "validarOperacaoTot" in JS and "dl-operacoes" in JS


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_normalizar_operacao():
    m = re.search(r"const normalizarOperacao = [^\n]+", JS)
    assert m
    ops = re.search(r"const OPERACOES = (\[[^\]]*\]);", JS).group(1)
    codigo = f"const OPERACOES={ops};{m.group(0)}\nconst v=(t)=>{{const n=normalizarOperacao(t);return OPERACOES.includes(n)?n:null}};console.log(JSON.stringify(['*r',' i ','M','*M','X','','*','II','R I'].map(v)));"
    r = subprocess.run([NODE, "-e", codigo], capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    assert json.loads(r.stdout) == ["*R", "I", "M", "*M", None, None, None, None, None]


def test_mensagem_e_estilo_de_erro():
    assert "Operação inválida: use I, *I, R, *R, M ou *M" in JS
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    assert "op-invalida" in html and "inp-operacao" in html
