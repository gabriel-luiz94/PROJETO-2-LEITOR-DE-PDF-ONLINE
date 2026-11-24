"""TASK-249: no MAPA, as linhas do ponto (e os campos de linha/estrutura dos formulários do mapa) editam os parâmetros dos ativos com o
mesmo balão das tabelas principais — Ctrl+Enter, duplo clique, mouse parado; dentro do <dialog> modal o balão vai junto."""
import os

import config

STATIC = config.STATIC_DIR


def _ler(nome):
    with open(os.path.join(STATIC, nome), encoding="utf-8") as f:
        return f.read()


def test_balao_aceita_campos_extras_e_desenha_dentro_do_dialog():
    js = _ler("ativo_balao.js")
    assert "cfgBase.seletorExtra" in js and "cfgBase.preparar(input)" in js
    assert "input._ativoBalao ? Object.assign({}, cfgBase, input._ativoBalao) : cfgBase" in js
    assert "S.input.closest('dialog[open]')" in js
    assert "abrir(e.target, { foco: true });" in js                                    # Ctrl+Enter no campo extra
    assert "cfg.rotulo ? cfg.rotulo(S.input)" in js


def test_resumo_liga_os_campos_do_mapa_ao_balao():
    js = _ler("resumo.js")
    assert "seletorExtra: 'input[data-sugestao-origem=\"OUTROS\"]:not([data-sem-balao])'" in js
    assert "function _itensDoBalao(row, idx)" in js and "function _gravarParametroNoTexto(texto, i, nomeParam, valor)" in js
    assert "itens: input => { const idx = Number(input.dataset.parIdx); return _itensDoBalao(tableStates.outros.data[idx], idx); }" in js


def test_campos_do_mapa_dizem_a_linha_e_a_operacao():
    js = _ler("vinculacao_ui.js")
    assert js.count("ativo._operacao = () =>") >= 3
    assert "ativo.dataset.linhaOutros = String(item.indice)" in js and "ativo.dataset.linhaOutros = String(indice)" in js
    assert js.count("'data-sem-balao': ''") == 2                                           # estruturas fictícias (lista por vírgula) ficam de fora
