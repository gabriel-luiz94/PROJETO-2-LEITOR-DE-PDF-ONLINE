"""TASK-239: leitura do DXF no padrão da NOTA 9101578729 — anotação emoldurada com a guia saindo da borda da moldura, MT tracejada
na layer ANOTACOES correndo a 1 m da BT, marcas de "X" vermelhas na layer 0 riscando anotações; e o cartão do cabo com o nome do ponto."""
import os

import ezdxf
import pytest

import config
from services import dxf_geometria as dg
from services.dxf_service import extract_dxf_content

CFG_081 = {"layers_rede_mt": ["REDE_MT", "ANOTACOES", "=0"], "blocos_poste": ["poste", "pg", "KMLP"]}


def _desenho():
    doc = ezdxf.new()
    for nome, cor in (("ANOTACOES", 7), ("REDE_BT", 3), ("POSTE_INSTALANDO", 7), ("POSTE_EXISTENTE", 7), ("0", 7)):
        if nome not in doc.layers:
            doc.layers.add(nome, color=cor)
    msp = doc.modelspace()
    msp.add_lwpolyline([(0, 10), (40, 10)], dxfattribs={"layer": "ANOTACOES", "color": 1, "linetype": "DASHED"})   # MT nova (vermelha) na ANOTACOES
    msp.add_lwpolyline([(0, 9), (40, 9)], dxfattribs={"layer": "REDE_BT"})                                          # BT a 1 m
    # anotação de MT emoldurada, a 10 m da rede; a guia sai da borda de baixo da moldura e chega mais perto da BT do que da MT
    msp.add_lwpolyline([(19, 19), (32, 19), (32, 21.5), (19, 21.5)], close=True, dxfattribs={"layer": "POSTE_INSTALANDO"})
    msp.add_mtext("3xA35P(7,94)-13,8kV", dxfattribs={"layer": "POSTE_INSTALANDO", "char_height": 1.2, "insert": (25, 21)})
    msp.add_lwpolyline([(19.5, 18.6), (18, 9.3)], dxfattribs={"layer": "POSTE_EXISTENTE"})
    # anotação de BT riscada por um X vermelho na layer 0 (o X não é rede)
    msp.add_mtext("PB35 20 m", dxfattribs={"layer": "POSTE_INSTALANDO", "char_height": 1.2, "insert": (55, 30)})
    msp.add_lwpolyline([(53, 29), (63, 31)], dxfattribs={"layer": "0", "color": 1})
    msp.add_lwpolyline([(53, 31), (63, 29)], dxfattribs={"layer": "0", "color": 1})
    # MT de verdade na layer 0 (vermelha), um trecho só
    msp.add_lwpolyline([(40, 10), (70, 10)], dxfattribs={"layer": "0", "color": 1})
    return doc


def _linha(rows, texto):
    return next(r for r in rows if r["texto"] == texto)


def test_guia_pela_moldura_e_cabo_na_rede_do_seu_tipo():
    rows = extract_dxf_content(_desenho(), leitura=CFG_081)
    mt = _linha(rows, "3xA35P(7,94)-13,8kV")
    assert mt["_geoTipo"] == "MT" and mt["_geo"] == [[0, 10], [40, 10]]                 # pela guia da moldura, e MT mesmo com a BT mais perto da ponta


def test_sem_a_moldura_a_guia_nao_alcanca_o_texto():
    doc = _desenho()
    geo = dg.coletar_geometria(doc, CFG_081)
    geo["molduras"] = []
    rows = extract_dxf_content(doc, leitura=CFG_081)
    for r in rows:
        r.pop("_geo", None), r.pop("_geoTipo", None)
    dg.associar_textos(rows, geo)
    assert "_geo" not in _linha(rows, "3xA35P(7,94)-13,8kV")


def test_marca_x_nao_e_rede():
    geo = dg.coletar_geometria(_desenho(), CFG_081)
    camada0 = [r["pts"] for r in geo["redes"] if r["layer"] == "0"]
    assert camada0 == [[(40.0, 10.0), (70.0, 10.0)]]                                    # o X saiu; a MT da layer 0 ficou
    rows = extract_dxf_content(_desenho(), leitura=CFG_081)
    assert "_geo" not in _linha(rows, "PB35 20 m")                                      # o PB35 não vira MT por causa do X


def test_cabo_que_diz_a_tensao_nao_vai_para_a_rede_do_outro_tipo():
    rows = extract_dxf_content(_desenho(), leitura={"blocos_poste": ["pg"]})            # sem ANOTACOES/0 como MT: só há BT
    assert "_geo" not in _linha(rows, "3xA35P(7,94)-13,8kV")


@pytest.mark.parametrize("texto,nivel", [("3xA35P(7,94)-13,8kV 32 m", "MT"), ("3xA35P(7,94)", "MT"), ("3xC25-13,8kV", "MT"), ("PB35 23 m", "BT"),
                                         ("1xC16(16) 21 m", "BT"), ("3xC16(16)+16 35 m", "BT"), ("PB35 / RAMA25", "BT"), ("N3", None), ("12600D, CEJ2", None)])
def test_nivel_do_texto(texto, nivel):
    assert dg.nivel_do_texto(texto) == nivel


def test_cartao_do_cabo_usa_o_nome_do_ponto():
    src = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    assert "const numeroDoNo = no => nomeGrupoDoNo(no) || `P${linhaNo(no) + 1}`;" in src
    assert "return botao(numeroDoNo(no), () => {" in src and "return botao('P' + (linhaNo(no) + 1)" not in src
