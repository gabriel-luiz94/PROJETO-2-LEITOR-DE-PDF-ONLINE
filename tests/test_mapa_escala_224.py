"""TASK-224 — mapa de vinculação: escala 1:500, recorte por janela (redes grandes), pan/zoom por quadro e ajuste manual em metros."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import obras_arquivo as oa

NODE = shutil.which("node")
S = config.STATIC_DIR
MOD = os.path.join(S, "mapa_escala.js")


def _node(corpo):
    js = "const E = require(%s);\n%s" % (json.dumps(MOD), corpo)
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_escala_500_em_pixels_e_escala_mostrada():
    r = _node("""
const k = E.pxPorMetro(500);
console.log(JSON.stringify([k, E.denominadorDe(k, 1), E.rotuloDaEscala(E.denominadorDe(k, 1)), E.rotuloDaEscala(E.denominadorDe(k, .5)), E.denominadorDe(0, 1), E.denominadorDe(null, 1),
  E.escalaValida('500'), E.escalaValida('xx'), E.escalaValida(undefined), E.rotuloDaEscala(null)]));""")
    assert abs(r[0] - 7.5591) < 1e-3                                       # 1 m a 1:500 = 2 mm = 7,559 px (96 dpi)
    assert abs(r[1] - 500) < 1e-6 and r[2] == "1:500" and r[3] == "1:1.000"
    assert r[4] is None and r[5] is None and r[6] == "500" and r[7] == "ajustar" and r[8] == "ajustar" and r[9] == ""


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_barra_de_escala_usa_comprimento_redondo():
    r = _node("""console.log(JSON.stringify([E.barraDeEscala(7.559), E.barraDeEscala(.5), E.barraDeEscala(.05), E.barraDeEscala(500), E.barraDeEscala(0), E.barraDeEscala(NaN)]));""")
    assert r[0]["metros"] == 10 and r[0]["rotulo"] == "10 m" and r[0]["px"] <= 140
    assert r[1]["metros"] == 200 and r[2]["rotulo"] == "2 km" and r[3]["metros"] == 1                   # nunca menos de 1 m
    assert r[4] is None and r[5] is None


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_recorte_mostra_so_o_que_cruza_a_janela_e_avisa_so_quem_mudou():
    r = _node("""
const itens = [{b: [0, 0, 10, 10], v: true}, {b: [500, 500, 510, 510], v: true}, {b: [90, 90, 120, 120], v: true}, {b: [-5, -5, 1, 1], v: false}];
const vista = [0, 0, 100, 50], area = E.retanguloDaVista(vista, .5), mudou = [];
const n1 = E.recortar(itens, area, (it, v) => mudou.push([itens.indexOf(it), v]));
const n2 = E.recortar(itens, area, () => mudou.push('de novo'));
console.log(JSON.stringify([area, n1, mudou, n2, itens.map(i => i.v), E.contem(area, E.retanguloDaVista(vista, 0)), E.contem(area, [-80, 0, 10, 10]),
  E.caixaDe([[1, 5], [4, -2], [3, 9]], 2), E.vistaCentrada(100, 50, 40, 20)]));""")
    assert r[0] == [-50, -25, 150, 75]
    assert r[1] == 3 and r[2] == [[1, False], [2, False], [3, True]]                                   # 1 e 2 ficam fora da janela; o 3 (antes escondido) entra
    assert r[3] == 0 and r[4] == [True, False, False, True]
    assert r[5] is True and r[6] is False and r[7] == [-1, -4, 6, 11] and r[8] == [80, 40, 40, 20]


def test_ajuste_manual_em_metros_vai_para_a_obra_com_limites():
    dados = {"cabos": [{"ativo": "MT1 40 m", "operacao": "I", "_mapaCabo500": [-12.5, 8]}, {"ativo": "MT2 40 m", "operacao": "I", "_mapaCabo500": [1e9, 0]}],
             "outros": [{"ativo": "9100D,N3", "entidade": "POSTE", "_mapaPoste500": [7.94, -3.97]}, {"ativo": "9200D,N3", "entidade": "POSTE", "_mapaPoste500": [False, 1]}]}
    exportada = oa.montar_exportacao({"nome": "Rede", "projeto": "081", "dados_json": dados})
    r = oa.validar_importacao(json.loads(json.dumps(exportada)), "081")["dados"]
    assert r["cabos"][0]["_mapaCabo500"] == [-12.5, 8] and r["outros"][0]["_mapaPoste500"] == [7.94, -3.97]
    assert "_mapaCabo500" not in r["cabos"][1] and "_mapaPoste500" not in r["outros"][1]               # absurdo ou não numérico: descartado


def test_ligacao_na_tela_e_rota_sem_cache():
    js = open(os.path.join(S, "vinculacao_ui.js"), encoding="utf-8").read()
    html = open(os.path.join(S, "resumo.html"), encoding="utf-8").read()
    res = open(os.path.join(S, "resumo.js"), encoding="utf-8").read()
    assert html.index("/static/mapa_escala.js") < html.index("/static/vinculacao_ui.js")
    assert "'_mapaPoste500', '_mapaCabo500'" in res                                                   # desfazer/salvar/duplicar conhecem o campo novo
    corpo = js[js.index("function abaMapa(corpo)"):]
    assert "ME.pxPorMetro(500)" in corpo and "'Escala do mapa'" in corpo and "vz-grafo-escala" in corpo
    assert "raiz.requestAnimationFrame(() => { quadroVista = 0; aplicarVistaJa(); })" in corpo           # pan/zoom: uma atualização por quadro
    assert "it.el.remove()" in corpo and "vz-camada-nos" in corpo                                      # recorte tira do DOM (display:none ainda custava layout)
    from fastapi.testclient import TestClient
    import app as appmod
    r = TestClient(appmod.app).get("/static/mapa_escala.js")
    assert r.status_code == 200 and "no-store" in r.headers["cache-control"]
