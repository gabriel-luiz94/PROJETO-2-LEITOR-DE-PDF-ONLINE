"""TASK-254: Exibição → «Linhas da tabela Demais ativos»: mostrar/ocultar grupos (@nome), fictícios (F) e comentários; ligados por padrão."""
import os

import config


def _ler(nome):
    with open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8") as f:
        return f.read()


def test_menu_exibicao_tem_os_tres_botoes_ligados_por_padrao():
    html = _ler("resumo.html")
    i = html.index('id="menu-exibicao-itens"')
    menu = html[i:html.index("Ao digitar", i)]
    for k in ("grupos", "ficticios", "comentarios"):
        assert f'data-linhas-outros="{k}"' in menu
    assert menu.count('class="chip-toggle on"') >= 5 and menu.count('aria-pressed="true"') >= 5
    assert "body.oculta-grupos #table-outros tr.linha-grupo-decl" in html and "tr.linha-ficticia" in html and "tr.linha-comentario" in html


def test_js_marca_linha_ficticia_e_lembra_a_escolha():
    js = _ler("resumo.js")
    assert "tr.classList.add('linha-ficticia')" in js
    assert "localStorage.getItem(chave(k)) !== '0'" in js and "ITENS = { grupos: 'oculta-grupos', ficticios: 'oculta-ficticios', comentarios: 'oculta-comentarios' }" in js
