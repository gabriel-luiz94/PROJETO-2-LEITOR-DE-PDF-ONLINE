"""TASK-236: PDF vetorial → rede pelas linhas coloridas, postes nas pontas, escala 1:500/1:1000 pelas metragens dos cabos, tudo em metros."""
import math

import pytest

pymupdf = pytest.importorskip("pymupdf")

from services import pdf_geometria as pg

PT = 0.0254 / 72


def _pdf(denominador, metros=(38, 40, 35), com_metragem=True, paginas=1):
    """Rede MT em linha reta: vãos de `metros` (desenhados na escala), anotação do cabo acima do meio, anotação do poste ao lado da ponta."""
    doc = pymupdf.open()
    k = PT * denominador                                                      # metros por pt
    for _ in range(paginas):
        page = doc.new_page(width=1191, height=842)
        x, y = 80.0, 400.0
        page.insert_text((x - 30, y + 14), "11300D, N3", fontsize=4)
        for m in metros:
            dx = m / k
            sh = page.new_shape()
            sh.draw_line((x, y), (x + dx, y))
            sh.finish(color=(1, 0, 0), width=0.8, closePath=False)
            sh.commit()
            page.insert_text((x + dx / 2 - 20, y - 3), "3xA35P(7,94)-13,8kV" + (f" {m} m" if com_metragem else ""), fontsize=4)
            x += dx
            page.insert_text((x + 2, y + 6), "11300D, N3", fontsize=4)
        sh = page.new_shape()                                                  # rua (preta): não é rede
        sh.draw_line((50, 600), (900, 600))
        sh.finish(color=(0, 0, 0), width=0.5, closePath=False)
        sh.commit()
    return doc


@pytest.mark.parametrize("den", [500, 1000])
def test_escala_escolhida_pelas_metragens_e_geometria_em_metros(den):
    with _pdf(den) as doc:
        itens, fundo, escala = pg.extrair_pdf_com_geometria(doc)
    assert escala["denominador"] == den and escala["confirmada"] and escala["amostras"] == 3
    cabos = [i for i in itens if "kV" in i["texto"]]
    assert all("_geo" in c and c["_geoTipo"] == "MT" for c in cabos)
    comp = [math.dist(c["_geo"][0], c["_geo"][-1]) for c in cabos]
    assert [round(x) for x in comp] == [38, 40, 35]                        # comprimento real em metros
    postes = [i for i in itens if i["texto"] == "11300D, N3"]
    assert sum(1 for p in postes if "_poste" in p) >= 3                     # ponta da rede perto da anotação vira o poste
    for p in postes:
        assert "_posteAprox" in p and p["_x"] < 1191                         # _x/_y continuam em pontos da página
    assert fundo["redes"] and all(r["t"] == "MT" for r in fundo["redes"]) and len(fundo["postes"]) == 4
    # y para cima: a rede (y = 400 pt) fica acima da rua (y = 600 pt) em metros
    assert fundo["redes"][0]["p"][0][1] == round(-400 * PT * den, 2)


def test_sem_metragem_vale_1_500_nao_confirmada_e_paginas_lado_a_lado():
    with _pdf(1000, com_metragem=False, paginas=2) as doc:
        itens, fundo, escala = pg.extrair_pdf_com_geometria(doc)
    assert escala == {"denominador": 500, "metros_por_pt": PT * 500, "amostras": 0, "desvio": None, "confirmada": False}
    x1 = max(i["_posteAprox"][0] for i in itens if i["pagina"] == 1)
    x2 = min(i["_posteAprox"][0] for i in itens if i["pagina"] == 2)
    assert x2 > x1                                                            # a 2ª folha não cai em cima da 1ª


def test_escolher_escala_e_pdf_sem_desenho():
    assert pg.escolher_escala([PT * 1000 * 1.1, PT * 1000 * 0.95, PT * 1000])["denominador"] == 1000
    assert pg.escolher_escala([PT * 700])["confirmada"] is False                # uma amostra só não confirma
    doc = pymupdf.open(); doc.new_page().insert_text((50, 50), "N3", fontsize=8)
    itens, fundo, escala = pg.extrair_pdf_com_geometria(doc)
    assert fundo is None and itens[0]["texto"] == "N3" and "_geo" not in itens[0] and "_posteAprox" in itens[0]


def test_rotas_devolvem_geometria_e_escala_do_pdf():
    from fastapi.testclient import TestClient
    import app as appmod
    with _pdf(500) as doc:
        dados = doc.tobytes()
    r = TestClient(appmod.app).post("/upload", files={"file": ("rede.pdf", dados, "application/pdf")}).json()
    assert r["escala"]["denominador"] == 500 and r["geometria"]["redes"] and any("_geo" in i for i in r["data"])
