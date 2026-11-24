"""TASK-243: pré-preenchimento do ativo do cabo nos formulários do Mapa — «3» → «3x», «A35P(» → «7,94)-13,8kV», «-» → «13,8kV»,
aprendendo a continuação com os cabos da obra. Conferido no Chromium digitando no «+ Cabo» (Tab aceita, Backspace apaga, digitar substitui)."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
EX = ["3xA35P(7,94)-13,8kV 40 m", "3xA35P(7,94)-13,8kV 25 m", "1xC16(16) 21 m", "3xC25-13,8kV 32 m", "3xA70P(9,5)-34,5kV 30 m"]


def _sugerir(casos, exemplos):
    js = ("const C = require('./static/cabo_autocompletar.js');"
          f"console.log(JSON.stringify({json.dumps(casos)}.map(t => C.sugerir(t, {json.dumps(exemplos)}))));")
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=30, cwd=config.BASE_DIR)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_sugestoes_pedidas():
    casos = ["3", "3xA34P(", "A35P(", "3xC25-", "3xA35P(7,94)-"]
    assert _sugerir(casos, []) == ["x", "7,94)-13,8kV", "7,94)-13,8kV", "13,8kV", "13,8kV"]


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_aprende_com_os_cabos_da_obra():
    casos = ["3xA35P(", "1xC16(", "3xA70P(", "2xC25(", "3x", "35", "PB35-", "PB35 -", "N3", ""]
    assert _sugerir(casos, EX) == ["7,94)-13,8kV", "16)", "9,5)-34,5kV", "", "", "", "", "", "", ""]
    # a tensão mais comum na obra
    assert _sugerir(["3xC25-"], ["X-34,5kV", "Y-34,5kV", "Z-13,8kV"]) == ["34,5kV"]


def test_ligado_nos_formularios_de_cabo_do_mapa():
    ui = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    assert "if (cabo) completarCabo(ativo);" in ui                                     # + Cabo e Traçar cabo
    assert ui.count("completarCabo(ativo);") == 2                                      # e Editar cabo
    assert html.index("cabo_autocompletar.js") < html.index("vinculacao_ui.js")
