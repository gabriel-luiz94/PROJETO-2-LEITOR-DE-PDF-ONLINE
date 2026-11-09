"""TASK-125 — o modal de vinculação lista o que está na tabela Outros: linha ESTRUTURA/POSTE e também a ainda sem classificação ('0'/vazia) com ativo.
Entidade escolhida à mão (CHAVE, RAMAIS…), linha sem ativo e operação R/*R continuam de fora. Paridade JS × Python."""
import json
import shutil
import subprocess

import pytest

from services import vinculacao_regras as vr
from services.autonomo import vinculacao as av

NODE = shutil.which("node")
OUTROS = [
    {"entidade": "ESTRUTURA", "operacao": "I", "ativo": "9400D, N1"},
    {"entidade": "0", "operacao": "I", "ativo": "121000R, N1"},          # classificação não reconheceu
    {"entidade": "", "operacao": "M", "ativo": "N3"},
    {"entidade": "0", "operacao": "I", "ativo": ""},                      # sem ativo
    {"entidade": "CHAVE", "operacao": "I", "ativo": "1-CF1"},             # escolhida à mão
    {"entidade": "0", "operacao": "R", "ativo": "N1"},                    # retirada não entra
    {"entidade": "RAMAIS", "operacao": "I", "ativo": "RA1"},
]
ESPERADO = [0, 1, 2]


def test_python_lista_o_que_esta_na_tabela():
    assert [i for i, _ in av.estruturas_disponiveis(OUTROS)] == ESPERADO
    assert [i for i, r in enumerate(OUTROS) if vr._eh_linha_estrutura(r)] == ESPERADO


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_js_igual_ao_python():
    cod = """
const VR = require('./static/vinculacao_regras.js');
const o = JSON.parse(require('fs').readFileSync(0, 'utf8'));
console.log(JSON.stringify(VR.estruturasDisponiveis(o)));
"""
    r = subprocess.run([NODE, "-e", cod], input=json.dumps(OUTROS), capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    assert json.loads(r.stdout) == ESPERADO
