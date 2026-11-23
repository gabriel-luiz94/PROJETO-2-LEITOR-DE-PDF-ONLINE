"""TASK-229 — mapa: voltar às coordenadas do CAD (descartar posições ajustadas à mão) e avisos de posição ajustada / 1:500 maior que a janela."""
import os

import config


def test_botao_posicoes_do_cad_aviso_de_ajustados_e_aviso_de_1_para_500():
    ui = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    corpo = ui[ui.index("function abaMapa(corpo)"):]
    assert "botao('Posições do CAD', () => restaurarPosicoesDoCad()" in corpo
    fn = corpo[corpo.index("function restaurarPosicoesDoCad()"):corpo.index("function pontosSemGrupo()")]
    assert "host.antes();" in fn and "delete linha[campo]" in fn and "delete c[campoC]" in fn                 # um passo de Ctrl+Z; só o campo do modo atual (Ajustar × 1:500)
    assert "ponto(s) com posição ajustada à mão" in corpo and "Em 1:500 só parte da rede cabe na janela" in corpo
