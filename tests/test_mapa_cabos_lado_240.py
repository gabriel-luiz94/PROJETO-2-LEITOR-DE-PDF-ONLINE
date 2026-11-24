"""TASK-240: no mapa, cabos entre o MESMO par de pontos ficam lado a lado (juntos pelo par, não pela geometria de cada um), cada um do seu
lado qualquer que seja o sentido do traçado, com espaço mínimo em pixels de tela em qualquer zoom (conferido também no navegador: 9 px)."""
import os

import config

SRC = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()


def test_faixas_pelo_par_de_pontos():
    assert "function faixasDosCabos(lista)" in SRC and "faixasDosCabos(desenhos);" in SRC
    assert "const chave = par ? par.join('|') : d.v;" in SRC                     # par de pontos (sem uma das pontas: o próprio vão)
    assert "d.inverte = !!par && a !== par[0];" in SRC                             # traçado ao contrário fica do mesmo lado
    assert "deslocamento: (i - (indices.length - 1) / 2) * 10" not in SRC           # antes: só entre cabos da mesma chave de vão


def test_espaco_minimo_na_tela_em_qualquer_zoom():
    assert "const SEP_CABO_PX = 9, SEP_CABO_MIN = 10;" in SRC
    assert "SEP_CABO_PX * (vb ? vb[2] / (svg.clientWidth || W) : 1)" in SRC
    assert "Math.abs(sep - unidadesDaSeparacao) > .25 * unidadesDaSeparacao) desenharPosicoes();" in SRC   # zoom mudou: refaz o espaço
    assert "const deslocamento = d.faixa * unidadesDaSeparacao * (d.inverte ? -1 : 1);" in SRC
