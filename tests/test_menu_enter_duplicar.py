"""Menu do botão direito pelo teclado, "Duplicar linha", Enter que cria linha abaixo e sugestões só da base (static/resumo.js / resumo.html)."""
import os

import config

JS = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
HTML = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()


def test_menu_tem_duplicar_e_teclado():
    assert 'id="ctx-duplicar"' in HTML and "Duplicar linha" in HTML
    assert HTML.index('id="ctx-add-below"') < HTML.index('id="ctx-duplicar"') < HTML.index('id="ctx-entidade"')          # logo depois de "Inserir linha abaixo"
    for t in ("itensDoMenu", "marcarItemDoMenu", "ArrowDown", "'Escape'", ".ctx-r-item.kb"):
        assert t in JS or t in HTML, t
    corpo = JS[JS.index("function duplicarLinhas"):][:1500]
    assert "splice(depois, 0, ...copias)" in corpo and "bloqueada" in corpo and "vinculoEstruturas" in corpo      # cópia logo abaixo, sem bloqueio nem vínculo do original


def test_enter_cria_linha_abaixo_e_nao_empilha_vazias():
    i = JS.index("// Enter cria uma linha NOVA logo abaixo")
    trecho = JS[i:i + 420]
    assert "addRow(type, idx)" in trecho and "irParaCelula(type, idx + 1, 1, 1)" in trecho and "trim() === ''" in trecho


def test_sugestoes_so_da_base():
    i = JS.index("function showAutocomplete(input, type) {")
    corpo = JS[i:i + 900]
    assert "ativoSets" not in corpo                                                    # nunca as linhas já digitadas
    assert "indiceDeNomes(orcamentoBaseData" in corpo and "'CABOS'" in corpo and "type !== 'cabos'" in corpo
