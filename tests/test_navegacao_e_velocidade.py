"""Enter na operação vai ao ativo, ← no começo do ativo volta à operação, linha "digitar" sempre presente na caixa de parâmetros e caches que deixam o preenchimento rápido."""
import os

import config

JS = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
DICA = open(os.path.join(config.STATIC_DIR, "ativo_dica.js"), encoding="utf-8").read()


def test_enter_na_operacao_vai_ao_ativo_da_mesma_linha():
    i = JS.index("// Enter na operação (escolhida ou digitada)")
    trecho = JS[i:i + 520]
    assert "querySelector('.inp-ativo')" in trecho and "tot-ativo-input" in trecho and "alvo.focus()" in trecho


def test_seta_esquerda_no_inicio_do_ativo_vai_a_operacao():
    i = JS.index("// ← no começo do ativo volta para a Operação")
    trecho = JS[i:i + 900]
    assert "selectionStart !== 0" in trecho and "input.inp-operacao" in trecho and "op.focus()" in trecho
    assert "#body-cabos .inp-ativo" in trecho and "#body-outros-vazio .inp-ativo-vazio" in trecho           # as duas tabelas e as linhas em branco


def test_caixa_de_parametros_sempre_tem_a_linha_de_digitar():
    assert "dp-digitar" in DICA and "✎ digitar outro valor…" in DICA and 'usar "' in DICA
    assert "valoresVisiveis(input, dicaVez, N).length + 1" in DICA                                       # as setas passam por ela
    assert "if (st.valorI === vis.length)" in DICA                                                       # aceitar nela só avança (o texto já está escrito)


def test_contextos_que_varrem_a_base_sao_guardados():
    """Antes cada tecla refazia o contexto da base INTEIRA (milhares de linhas) várias vezes: 100–350 ms por tecla. Agora só quando a base, o projeto ou a convenção mudam."""
    i = JS.index("function _ctxOutros()")
    corpo = JS[i - 400:i + 900]
    assert "_cacheCtxOutros" in corpo and "c.base === base" in corpo and "c.postes === postes" in corpo
    assert "_cacheAssin" in JS and "function aquecerCaches" in JS and "setTimeout(aquecerCaches" in JS
    assert "aquecer" in DICA
