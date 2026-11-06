"""TASK-116: Resumo da rede — o neutro não entra na Média e o cabo sem comprimento usa a linha de baixo (JS: anotacao_cabos.js + resumo.js)."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
_JS = """
const A = require(%s);
const t = JSON.parse(require('fs').readFileSync(0, 'utf8'));
console.log(JSON.stringify(t.map(x => A.comprimentoSozinho(x))));
"""


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_comprimento_sozinho():
    textos = ["45 m", "45", "7,5M", " 12.5 m ", "3xA35P 45 m", "CAA2 ABC", "", "m", "45 mm", "-45"]
    r = subprocess.run([NODE, "-e", _JS % json.dumps(os.path.join(config.STATIC_DIR, "anotacao_cabos.js"))], input=json.dumps(textos),
                       capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    assert json.loads(r.stdout) == [45, 45, 7.5, 12.5, None, None, None, None, None, None]


def test_resumo_js_usa_a_regra_nova():
    js = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    i = js.index("function calcularResumoRede")
    corpo = js[i:js.index("function formatarNumeroResumoRede")]
    assert "if (!it.neutro) resumo.redeMedia" in corpo                    # neutro fora da Média
    assert "AnotacaoCabos.comprimentoSozinho" in corpo and "linhasCabos[idx + 1]" in corpo          # comprimento na linha de baixo
    assert "extrairPrefixoCaboSemComprimento" in corpo
