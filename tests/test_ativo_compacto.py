"""Ocultar parâmetros ao sair da linha: texto compacto do ativo (static/ativo_compacto.js) e o rótulo "manual" do balão."""
import json
import os
import subprocess

import pytest

import config
from tests import oraculo_tela as o

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado")


def js(codigo):
    r = subprocess.run([o.NODE, "-e", codigo], capture_output=True, text=True, timeout=30, cwd=config.STATIC_DIR)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout.strip().splitlines()[-1])


@pytest.mark.parametrize("entrada,saida", [
    ("RL, PR(D,X), ATRINT", "RL, PR, ATRINT"),
    ("CE_4(12600D,A35P,A35P)", "CE_4"),
    ("2-N3(XLPE) 1-T3", "2-N3 1-T3"),
    ("PR(D,(X))", "PR"),
    ("N1, N2", "N1, N2"),                       # sem parâmetros: igual
    ("PR(", "PR("),                             # ainda digitando: não mexe
    ("RL, PR(D", "RL, PR(D"),
    ("(x)", "(x)"),                             # parêntese sem nome antes: não é parâmetro
    ("", ""),
])
def test_compactar(entrada, saida):
    assert js(f"console.log(JSON.stringify(require('./ativo_compacto.js').compactar({json.dumps(entrada)})))") == saida


def test_balao_chama_de_manual_o_que_o_usuario_preencheu():
    assert js("console.log(JSON.stringify(require('./ativo_status.js').rotuloOrigem('pedido')))") == "manual"
    assert js("console.log(JSON.stringify(require('./ativo_status.js').rotuloOrigem('vinculo')))") == "vínculo"


def test_com_a_linha_em_foco_a_linha_fica_so_com_o_nome_e_o_parentese_abre_o_balao():
    """Pedido de 18/11: clicar na linha mantém só o nome; ao digitar "(" os parâmetros PREENCHIDOS aparecem no balão (não na linha)."""
    import os
    import config
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    js = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    assert "td.cel-ativo-comp.tem-params .ativo-compacto { display: flex; }" in html
    assert "tr:not(:focus-within) td.cel-ativo-comp.tem-params" not in html          # não depende mais de a linha ter foco
    assert "editando-par" not in html and "editando-par" not in js
    assert "window.mostrarBalaoParametros = el =>" in js and "window.esconderBalaoParametros" in js
    assert "(inpAt.selectionStart || 0) > i && window.mostrarBalaoParametros" in js
