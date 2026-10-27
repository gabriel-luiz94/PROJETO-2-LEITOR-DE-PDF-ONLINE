"""TASK-091 — geometria implícita do DXF: rede por cor+layer, postes por nome de bloco, guias e raio."""
import ezdxf
from services.dxf_service import extract_dxf_content
from services.dxf_geometria import classificar_rede


def _doc():
    doc = ezdxf.new("R2010")
    for nome, aci in (("REDE_MT", 1), ("REDE_BT", 3), ("GUIA", 7), ("TXT", 7)):
        doc.layers.add(nome, color=aci)
    blk = doc.blocks.new("POSTE_CONCRETO")
    blk.add_circle((0, 0), 0.5)
    return doc


def _texto(msp, txt, x, y, h=2.0):
    msp.add_text(txt, dxfattribs={"insert": (x, y), "height": h, "layer": "TXT"})


def test_cor_decide_e_layer_confirma():
    assert classificar_rede("REDE_MT", "#ff0000") == "MT"
    assert classificar_rede("REDE_BT", "#00ff00") == "BT"
    assert classificar_rede("REDE_BT", "#ff0000") is None       # cor vermelha numa layer BT: não confirma
    assert classificar_rede("OUTRA", "#ff0000") is None
    assert classificar_rede("REDE_MT", "#000000") == "MT"       # cor neutra: vale a layer


def test_cabo_ganha_a_linha_da_rede_mais_proxima_e_nao_cria_linhas():
    doc = _doc(); msp = doc.modelspace()
    msp.add_lwpolyline([(0, 0), (50, 0), (100, 10)], dxfattribs={"layer": "REDE_MT", "color": 1})
    msp.add_line((0, 30), (100, 30), dxfattribs={"layer": "REDE_BT", "color": 3})
    _texto(msp, "CAZ 9,5", 40, 3)
    _texto(msp, "CAZ 4", 40, 33)
    linhas, geo = extract_dxf_content(doc, com_geometria=True)
    assert [l["texto"] for l in linhas] == ["CAZ 4", "CAZ 9,5"]          # só textos
    mt = next(l for l in linhas if l["texto"] == "CAZ 9,5"); bt = next(l for l in linhas if l["texto"] == "CAZ 4")
    assert mt["_geoTipo"] == "MT" and mt["_geo"] == [[0.0, 0.0], [50.0, 0.0], [100.0, 10.0]]
    assert bt["_geoTipo"] == "BT" and bt["_geo"] == [[0.0, 30.0], [100.0, 30.0]]
    assert len(geo["redes"]) == 2


def test_fora_do_raio_nao_associa():
    doc = _doc(); msp = doc.modelspace()
    msp.add_line((0, 0), (100, 0), dxfattribs={"layer": "REDE_MT", "color": 1})
    _texto(msp, "LONGE", 40, 20)          # 20 > 4 x 2
    linhas = extract_dxf_content(doc)
    assert "_geo" not in linhas[0]


def test_guia_leva_o_texto_ate_a_rede():
    doc = _doc(); msp = doc.modelspace()
    msp.add_line((0, 0), (100, 0), dxfattribs={"layer": "REDE_MT", "color": 1})
    _texto(msp, "CAZ 9,5", 40, 40)                                               # longe da rede
    msp.add_line((40, 40), (40, 1), dxfattribs={"layer": "GUIA"})                # guia (fora das layers de rede)
    linhas = extract_dxf_content(doc)
    assert linhas[0]["_geoTipo"] == "MT"


def test_poste_por_nome_do_bloco_inclusive_aninhado():
    doc = _doc(); msp = doc.modelspace()
    msp.add_blockref("POSTE_CONCRETO", (10, 10))
    pai = doc.blocks.new("CONJUNTO"); pai.add_blockref("poste_x", (0, 0)); doc.blocks.new("poste_x")
    msp.add_blockref("CONJUNTO", (200, 200))
    _texto(msp, "N1", 12, 11)
    _texto(msp, "N3", 202, 201)
    linhas = {l["texto"]: l for l in extract_dxf_content(doc)}
    assert linhas["N1"]["_poste"] == [10.0, 10.0]
    assert linhas["N3"]["_poste"] == [200.0, 200.0]


def test_sem_geometria_nao_muda_o_retorno_antigo():
    doc = _doc(); _texto(doc.modelspace(), "ABC", 1, 1)
    linhas = extract_dxf_content(doc)
    assert isinstance(linhas, list) and "_geo" not in linhas[0] and "_poste" not in linhas[0]
