"""TASK-242: no Mapa, «Editar» do cabo abre o formulário novo (o mesmo do + Cabo: F/I/M/R e comprimento separado) e «Editar» da
estrutura/linha no painel do ponto abre a linha em «Linhas do ponto» — o cartão antigo só fica para o poste fictício e quem não tem o editor
de pontos. Conferido no Chromium: o cabo virou R com 42 m e a linha do poste virou *I."""
import os

import config

SRC = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()


def test_cabo_abre_o_formulario_novo():
    assert "controleEdicao('cabo', ic)" not in SRC                                   # nem na lista nem no painel do ponto
    assert SRC.count("botaoEditarCabo(ic),") == 2                                     # lista e painel do ponto
    assert "abrirEditor('Editar cabo ' + (ic + 1), 'Salvar edição', campos" in SRC
    assert "campo('Ativo ou rótulo, sem a metragem', ativo), campo('Operação', operacao), campo('Comprimento em metros', metros)" in SRC
    assert "const operacao = seletorOperacao(opDaLinha(ref));" in SRC
    assert "ativo: mt ? rotulo + ' ' + mt + ' m' : rotulo" in SRC                      # a metragem volta para o ativo da tabela


def test_estrutura_abre_a_linha_no_editor_do_ponto():
    assert "'data-linha-ponto': item.indice" in SRC
    assert "dialogo.querySelector(`.vz-lp-painel[data-linha-ponto=\"${indice}\"]`)" in SRC
    assert "editarLinhaNoPonto(linhaDe(ie), 'estrutura ' + codigo) || controleEdicao(" in SRC
    assert "editarLinhaNoPonto(linhaNo(no), 'linha do ponto') || controleEdicao(" in SRC
