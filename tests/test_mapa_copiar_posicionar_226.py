"""TASK-226 — mapa: o ponto novo (ou copiado) pede para ser posicionado com o mouse; copiar ponto e copiar cabo."""
import os

import config

S = config.STATIC_DIR


def _ler(nome):
    return open(os.path.join(S, nome), encoding="utf-8").read()


def test_posicionamento_do_ponto_novo_acompanha_o_mouse_clique_fixa_esc_mantem():
    ui = _ler("vinculacao_ui.js")
    corpo = ui[ui.index("function abaMapa(corpo)"):]
    assert "function iniciarPosicionamento(linha)" in corpo and "function encerrarPosicionamento(confirmou, silencioso)" in corpo
    assert "if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); encerrarPosicionamento(false); }" in corpo           # Esc não fecha o modal
    assert "if (!colocando || mover) return;" in corpo and "guardarPosicao(colocando.id)" in corpo                                  # clique fixa; arrastar o fundo não conta
    assert ".vz-mapa.posicionando [data-arrastar]{pointer-events:none}" in ui                                                       # pontos/vãos não atrapalham o clique
    assert corpo.count("iniciarPosicionamento(") >= 4                                                                              # novo ponto, poste fictício e cópia
    assert corpo.index("avisos.textContent = texto;") < corpo.index("colocarNovoPonto(resultado.indice); iniciarPosicionamento(resultado.indice)")   # a instrução de posicionar não é apagada


def test_copiar_ponto_e_copiar_cabo_no_host_e_na_tela():
    res, ui = _ler("resumo.js"), _ler("vinculacao_ui.js")
    host = res[res.index("function copiarPontoNoMapa"):res.index("window.VinculacaoHost = {")]
    assert "GL.abrirPonto(nome" in host and "novosIds" in host and "_mapaPosteRef" in host                                           # grupo novo; estrutura do poste acompanha a cópia
    assert "'_mapaId'" in res[res.index("const CAMPOS_FORA_DA_COPIA"):][:700] and "'vinculoPostes'" in res[res.index("const CAMPOS_FORA_DA_COPIA"):][:700]   # a cópia nasce sem vínculos e sem posição
    assert "const indice = cabos.push(c) - 1;" in host and "copiarPontoNoMapa, copiarCaboNoMapa," in res
    assert "botao('Copiar ponto'" in ui and "botao('Copiar', () => copiarCabo(ic)" in ui
    assert ui.index("host.copiarPontoNoMapa(no.linhas)") < ui.index("iniciarPosicionamento(r.indice)")
