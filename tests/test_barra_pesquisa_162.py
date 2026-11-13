"""TASK-162 — a pesquisa da Base de Orçamento vira uma barra FIXA no topo (Ctrl+K leva o cursor para ela) e a "Consulta Rápida" deixa de existir."""
import os

import config


def ler(nome):
    return open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8").read()


def test_tela_tem_a_barra_fixa_e_nao_tem_mais_a_consulta_rapida():
    html = ler("orcamento.html")
    assert 'id="barra-pesquisa"' in html and "barra: document.getElementById('barra-pesquisa')" in html
    assert html.index('id="barra-pesquisa"') < html.index('<header class="page-header-card">')                # é a primeira coisa da tela
    for removido in ("quick-search", "qs-input", "qs-result", "quickSearch", "Consulta Rápida", "pesquisaRapida"):
        assert removido not in html, removido


def test_barra_e_o_atalho_ctrl_k():
    js = ler("base_ficha_ui.js")
    for trecho in ("function montarBarra(host)", ".bf-bar{position:sticky;top:0;", "Pesquisar ativo, apelido, código ou descrição…", "abrirPesquisa()", "barra.focar(consulta)",
                   "if (opcoes && opcoes.barra) montarBarra(opcoes.barra);"):
        assert trecho in js, trecho
    assert "pesquisaRapida" not in js                                                                       # o pop-up antigo e a pesquisa rápida saíram
    assert "janela('min(680px,94vw)'" not in js
    assert "e.key.toLowerCase() === 'k') { e.preventDefault(); abrirPesquisa(); }" in js
