"""TASK-175 — leitura do desenho (DXF) configurável por projeto: layers e cores de rede."""
import ezdxf
import pytest

from services import leitura_desenho as ld
from services.dxf_geometria import classificar_rede, coletar_geometria
from services.dxf_service import extract_dxf_content


def test_padrao_de_pernambuco_igual_ao_de_sempre():
    for layer, cor, esperado in [("REDE_MT", "#ff0000", "MT"), ("REDE_BT", "#00ff00", "BT"), ("REDE_BT", "#ff0000", None), ("OUTRA", "#ff0000", None),
                                 ("REDE_MT", "#000000", "MT"), ("rede_mt_x", "#ff0000", "MT"), ("REDE_MT", "", "MT"), ("REDE_MT", "#00ff00", None)]:
        assert classificar_rede(layer, cor) == classificar_rede(layer, cor, {}) == classificar_rede(layer, cor, ld.PADRAO) == esperado, (layer, cor)


def test_so_layer_com_curinga_azul_em_mt_primaria_vira_mt():
    cfg = {"layers_rede_mt": ["MT_*"], "regra": "SO_LAYER"}
    assert classificar_rede("MT_PRIMARIA", "#0000ff", cfg) == "MT"          # linha azul
    assert classificar_rede("MT_PRIMARIA", "#00ff00", cfg) == "MT"          # a cor é ignorada
    assert classificar_rede("XMT_PRIMARIA", "#0000ff", cfg) is None          # curinga vale para o nome inteiro
    assert classificar_rede("REDE_BT", "#00ff00", cfg) == "BT"               # BT segue o padrão


def test_cor_manda_confirma_e_layer_manda():
    cfg = {"layers_rede_mt": ["PRIM"], "cor_mt": 5, "regra": "COR_MANDA"}                  # ACI 5 = azul
    assert classificar_rede("PRIM", "#0000ff", cfg) == "MT" and classificar_rede("PRIM", "#ff0000", cfg) == "MT"      # vermelho já não é MT: cor neutra, vale a layer
    assert classificar_rede("PRIM", "#00ff00", {"layers_rede_mt": ["PRIM"], "regra": "COR_MANDA"}) is None            # cor BT numa layer MT descarta
    lm = {"layers_rede_mt": ["PRIM"], "regra": "LAYER_MANDA"}
    assert classificar_rede("PRIM", "#00ff00", lm) == "MT"                                   # a layer vence a cor
    assert classificar_rede("QUALQUER", "#ff0000", lm) == "MT" and classificar_rede("QUALQUER", "#000000", lm) is None   # sem layer de rede, vale a cor reconhecida


def test_formas_de_cor_e_cor_confirma_do_plugin():
    assert ld.tipo_pela_cor("#0000ff", {"cor_bt": "#0000FF"}) == "BT" and ld.tipo_pela_cor("#0000fe", {"cor_bt": "#0000FF"}) is None
    assert ld.tipo_pela_cor("#0a0aff", {"cor_mt": {"r": [0, 20], "g": [0, 20], "b": [200, 255]}}) == "MT"
    assert ld.tipo_pela_cor("#ff0000", {"cor_mt": [1, "#00ffff"]}) == "MT"
    assert ld.normalizar({"cor_confirma": False})["regra"] == "SO_LAYER" and ld.normalizar({"cor_confirma": True})["regra"] == "COR_MANDA"
    assert ld.normalizar({"regra": "layer_manda"})["regra"] == "LAYER_MANDA"
    assert ld.contem(["=0"], "0") and not ld.contem(["=0"], "01") and ld.contem(["MT"], "rede_mt")


def test_validacao_e_testar():
    assert ld.validar(None) == [] and ld.validar({"regra": "SO_LAYER", "cor_mt": 1, "cor_bt": "#00ff00"}) == []
    assert ld.validar({"regra": "X"}) and ld.validar({"cor_mt": "azul"}) and ld.validar([])
    r = ld.testar({}, "REDE_BT", "#ff0000")
    assert r["resultado"] == "DESCARTADA" and r["da_layer"] == "BT" and r["da_cor"] == "MT"
    assert ld.testar({}, "REDE_MT", 1)["resultado"] == "MT" and ld.testar({}, "OUTRA", 1)["resultado"] == "NAO_E_REDE"
    assert ld.testar({"layers_rede_mt": ["MT_*"], "regra": "SO_LAYER"}, "MT_PRIMARIA", 5)["tipo"] == "MT"


def _doc(layer):
    doc = ezdxf.new("R2010")
    doc.layers.add(layer, color=5)
    return doc


def test_dxf_usa_a_configuracao_do_projeto():
    doc = _doc("MT_PRIMARIA"); msp = doc.modelspace()
    msp.add_line((0, 0), (100, 0), dxfattribs={"layer": "MT_PRIMARIA", "color": 5})
    msp.add_text("3xA35P 40", dxfattribs={"insert": (40, 3), "height": 2.0, "layer": "0"})
    sem = coletar_geometria(doc)
    assert sem["redes"] == []                                                                  # sem a configuração: não é rede
    com = coletar_geometria(doc, {"layers_rede_mt": ["MT_*"], "regra": "SO_LAYER"})
    assert [r["tipo"] for r in com["redes"]] == ["MT"] and com["guias"] == []
    linhas, geo = extract_dxf_content(doc, com_geometria=True, leitura={"layers_rede_mt": ["MT_*"], "regra": "SO_LAYER"})
    assert linhas[0]["_geoTipo"] == "MT" and len(geo["redes"]) == 1
    assert "_geoTipo" not in extract_dxf_content(doc)[0]


def test_bloco_de_poste_e_guias_configuraveis():
    doc = ezdxf.new("R2010"); doc.layers.add("POSTES"); doc.layers.add("GUIA")
    blk = doc.blocks.new("P_X"); blk.add_circle((0, 0), 0.5)
    msp = doc.modelspace()
    msp.add_blockref("P_X", (10, 10), dxfattribs={"layer": "POSTES"})
    msp.add_line((0, 0), (5, 0), dxfattribs={"layer": "GUIA"})
    msp.add_line((0, 5), (5, 5), dxfattribs={"layer": "OUTRA"})
    assert coletar_geometria(doc)["postes"] == [] and len(coletar_geometria(doc)["guias"]) == 2
    g = coletar_geometria(doc, {"layers_poste": ["POSTES"], "layers_guia": ["GUIA"]})
    assert g["postes"] == [(10.0, 10.0)] and len(g["guias"]) == 1


def test_validar_leitura_do_plugin_aceita_os_campos_novos_e_rota_testar():
    from fastapi import HTTPException
    from routers.plugin_cad import TestarLeituraPayload, testar_leitura, validar_leitura
    ok = validar_leitura({"layers_rede_mt": ["MT_*"], "regra": "so_layer", "cor_mt": 5, "cor_bt": "#00FF00"})
    assert ok["regra"] == "SO_LAYER" and ok["cor_mt"] == 5 and ok["layers_rede_mt"] == ["MT_*"]
    with pytest.raises(HTTPException):
        validar_leitura({"regra": "XX"})
    with pytest.raises(HTTPException):
        validar_leitura({"cor_bt": "verde"})
    assert testar_leitura(TestarLeituraPayload(leitura={"layers_rede_mt": ["MT_*"], "regra": "SO_LAYER"}, layer="MT_PRIMARIA", cor="#0000ff"))["tipo"] == "MT"
    with pytest.raises(HTTPException):
        testar_leitura(TestarLeituraPayload(leitura={"regra": "XX"}, layer="a", cor="1"))


def test_ligacoes_na_tela_e_no_upload():
    import os
    import config
    ui = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    assert "abaDesenho" in ui and "/api/plugin/leitura/testar" in ui
    up = open(os.path.join(config.BASE_DIR, "routers", "upload.py"), encoding="utf-8").read()
    assert "_leitura_do_projeto" in up
    assert "/upload?projeto=" in open(os.path.join(config.STATIC_DIR, "script.js"), encoding="utf-8").read()
