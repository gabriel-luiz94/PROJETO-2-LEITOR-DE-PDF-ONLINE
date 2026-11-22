"""TASK-227 — nível MT/BT do cabo igual na lista e no mapa (o rótulo manda; o desenho só decide quando o rótulo não diz) e diagnóstico de cabos que não aparecem."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
pytestmark = pytest.mark.skipif(not NODE, reason="node indisponível")


def _node(corpo):
    js = "const VR = require(%s);\n%s" % (json.dumps(os.path.join(config.STATIC_DIR, "vinculacao_regras.js")), corpo)
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def test_nivel_do_cabo_ordem_linha_rotulo_dicionario_geometria_padrao():
    r = _node("""
const n = (c, dic) => { const x = VR.nivelDoCabo(c, dic); return [x.nivel, x.fonte, x.geo]; };
const dic = [{padrao: 'A35P', nivel: 'MT'}];
console.log(JSON.stringify([
  n({nivel: 'BT', textoOriginal: '3xA35P-13,8kV 10 m', _geoTipo: 'MT'}, []),         // nível digitado vence tudo
  n({textoOriginal: '3xA35P(7,94)-13,8kV 35 m', _geoTipo: 'BT'}, []),                // tensão no rótulo vence o desenho
  n({textoOriginal: '3xC16(16) 35 m', _geoTipo: 'MT'}, []),                          // rótulo não diz: manda o desenho
  n({textoOriginal: '3xC16(16) 35 m'}, []),                                          // nada diz: BT
  n({textoOriginal: 'PB35 20 m', _geoTipo: 'BT'}, []),
  n({ativo: '3xA35P-13,8kV 10 m'}, []),                                               // sem textoOriginal: usa o ativo
  n({textoOriginal: '3xC16(16) 35 m', _geoTipo: 'xx'}, [])                           // tipo de desenho inválido é ignorado
]));""")
    assert r == [["BT", "linha", "MT"], ["MT", "rótulo", "BT"], ["MT", "geometria", "MT"], ["BT", "padrão", ""], ["BT", "geometria", "BT"], ["MT", "rótulo", ""], ["BT", "padrão", ""]]


def test_lista_e_mapa_usam_a_mesma_conta_e_diagnostico_explica_o_que_falta():
    r = _node("""
const c = {textoOriginal: '3xA35P(7,94)-13,8kV 35 m', ativo: '3xA35P(7,94)-13,8kV 35 m', _geoTipo: 'BT'};
const analise = {cabos: {0: {vincula: false}}};
const mapa = VR.gruposVisuaisDoVao([0], [c], analise, [])[0].tipo, lista = VR.extraCabo(c, []).nivel;
const d1 = VR.diagnosticoDoCabo(c, {ids: ['a', 'b'], geo: [[0, 0], [1, 1]]}, []);
const d2 = VR.diagnosticoDoCabo({textoOriginal: '3xC16(16) 20 m', ativo: '3xC16(16) 20 m'}, {ids: [null, null], geo: null}, []);
const d3 = VR.diagnosticoDoCabo({textoOriginal: '3xC16(16) 20 m', _geoTipo: 'BT'}, {ids: ['a', null], geo: [[0, 0], [9, 9]]}, []);
const d4 = VR.diagnosticoDoCabo({textoOriginal: '3xC16(16) 20 m'}, {ids: ['a', null], geo: null}, []);
console.log(JSON.stringify([mapa, lista, d1.diverge, d1.motivo, d2.naoDesenhado, d2.motivo, d3.diverge, d3.motivo, d4.semGeometria, d4.naoDesenhado, d4.motivo]));""")
    assert r[0] == r[1] == "MT"                                                                    # antes o mapa dizia BT (pelo desenho) e a lista MT
    assert r[2] is True and r[3] == ""
    assert r[4] is True and "não aparece no mapa" in r[5]
    assert r[6] is False and "mais de 5 m de um poste" in r[7]
    assert r[8] is True and r[9] is False and "só a ponta ligada aparece" in r[10]


def test_mapa_mostra_alerta_filtro_e_aresta_tracejada():
    ui = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    corpo = ui[ui.index("function abaMapa(corpo)"):]
    assert "VR.diagnosticoDoCabo(" in corpo and "'A conferir (nível/geometria)'" in corpo and "' diverge'" in corpo
    assert "Use o filtro «A conferir»." in corpo and ".vz-mapa .vz-aresta.diverge{stroke-dasharray" in ui
