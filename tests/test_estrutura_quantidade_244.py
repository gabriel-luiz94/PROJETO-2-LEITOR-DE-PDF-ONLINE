"""TASK-244: no Mapa/vínculo, «2 FLABIT» são DUAS estruturas. As cópias vão no FIM da lista da linha (as posições — e os vínculos — das
que já existiam não mudam); na lista do ponto elas aparecem juntas («FLABIT · 2ª»). Python e JS iguais."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services.anotacao_outros import codigos_estruturas
from services.autonomo.totalizadora import classico_totalizadora

NODE = shutil.which("node")
CASOS = {
    "12600D, CEJ2, 2 FLABIT, IT-R": ["CEJ2", "FLABIT", "IT-R", "FLABIT"],
    "2 FLABIT": ["FLABIT", "FLABIT"],
    "N3, 1,5 FLABIT": ["N3", "FLABIT"],                    # quantidade não inteira: uma
    "3 N3, 2 IT-R": ["N3", "IT-R", "N3", "N3", "IT-R"],
    "9400D, N1, N4": ["N1", "N4"],                         # sem quantidade: como sempre
}


def test_python():
    for texto, esperado in CASOS.items():
        assert codigos_estruturas(texto, classico_totalizadora, None) == esperado, texto
    assert len(codigos_estruturas("30 FLABIT", classico_totalizadora, None)) == 20            # limite por linha


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_js_igual_ao_python():
    js = ("const A = require('./static/anotacao_outros.js');"
          f"console.log(JSON.stringify({json.dumps(list(CASOS))}.map(t => A.codigosEstruturas(t, null, {{}}))));")
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=30, cwd=config.BASE_DIR)
    assert r.returncode == 0, r.stderr
    assert json.loads(r.stdout) == list(CASOS.values())


def test_posicoes_antigas_nao_mudam():
    """O vínculo guarda a posição da estrutura na linha: as que já existiam continuam nas mesmas posições."""
    antes = ["CEJ2", "FLABIT", "IT-R"]                                                        # como era (uma por item)
    depois = codigos_estruturas("12600D, CEJ2, 2 FLABIT, IT-R", classico_totalizadora, None)
    assert depois[:len(antes)] == antes


def test_lista_do_ponto_mostra_as_copias_juntas():
    src = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    assert "function estruturasNaTela(lista)" in src and "estruturasNaTela(estruturas).map(([ie, ordem])" in src
    assert "(ordem > 1 ? ' · ' + ordem + 'ª' : '')" in src
