"""Lista de sugestões no estilo do VS Code (static/ide_lista.js): módulo carregável e ligações das telas."""
import os

import config

JS = open(os.path.join(config.STATIC_DIR, "ide_lista.js"), encoding="utf-8").read()


def test_api_da_lista():
    for t in ("criar(", "abrir", "mover", "escolherAtual", "fechar", "aberta", "role', 'listbox", "il-item.sel", "#094771"):
        assert t in JS, t


def test_ligada_no_resumo_antes_do_resumo_js():
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    assert html.index("/static/ide_lista.js") < html.index("/static/resumo.js")
    app = open(os.path.join(config.BASE_DIR, "app.py"), encoding="utf-8").read() if hasattr(config, "BASE_DIR") else open("app.py", encoding="utf-8").read()
    assert '"/static/ide_lista.js"' in app


def test_operacao_mostra_todas_as_opcoes_ao_entrar_no_campo():
    resumo = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    corpo = resumo[resumo.index("function abrirListaOperacao"):][:700]
    assert "const q = filtrar ? normalizarOperacao(inp.value) : ''" in corpo                  # sem filtrar: todas as opções (o datalist só mostrava as que casavam com o valor atual)
    assert "addEventListener('focusin'" in resumo and "e.target.select()" in resumo             # entra no campo: seleciona tudo e abre a lista completa


def test_botoes_de_marcar_linhas_nas_duas_tabelas():
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    for t in ("cabos", "outros"):
        for b in ("congelar", "comentar", "bloquear"):
            assert f'id="btn-{b}-{t}"' in html, (b, t)
    resumo = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    assert "restaurarSelecaoLinhas" in resumo and "linhasParaMarcar" in resumo
    assert 'id="cfg-avancado" open' in html and 'id="cfg-ocultar-params"' in html                 # a opção de ocultar parâmetros está no painel Configurações e visível
