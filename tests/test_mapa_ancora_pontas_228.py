"""TASK-228 — pontas de cabo sem poste (a mais de 5 m) seguem o ponto mais perto: o cabo não fica para trás quando o ponto é movido ou rearranjado."""
import os

import config


def test_ponta_solta_ancora_no_ponto_mais_perto_e_entra_no_layout_do_rearranjar():
    ui = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    corpo = ui[ui.index("function abaMapa(corpo)"):]
    assert "const RAIO_ANCORA_VISUAL = 25;" in corpo and "function pontaAncorada(d, i)" in corpo
    assert "v.ids.map((id, i) => id ? posicoes.get(id) : v.pontas && pontaAncorada(d, i))" in corpo        # a ponta solta acompanha o ponto (antes: coordenada fixa do CAD)
    assert "const a = d.v.ids[0] || (d.ancoras && d.ancoras[0])" in corpo                                  # o rearranjo enxerga a ligação cabo↔ponto mesmo sem vínculo (senão o ponto vira «isolado» e vai para a grade)
    assert "d.ancoras[i] = melhor;" in corpo and "dmin = RAIO_ANCORA_VISUAL" in corpo
