"""TASK-153 — usabilidade e design: telas sem depender da internet (Tailwind e fontes locais), hierarquia dos botões
(um principal por área, vermelho só no que apaga), menus "Mais" e "Exibição" no Resumo, foco do teclado e rodapé fora do conteúdo."""
import glob
import os
import re
from html.parser import HTMLParser

from fastapi.testclient import TestClient

import config
import database

ESTATICOS = config.STATIC_DIR


def ler(nome):
    with open(os.path.join(ESTATICOS, nome), encoding="utf-8") as f:
        return f.read()


def paginas():
    return sorted(glob.glob(os.path.join(ESTATICOS, "*.html")))


def dentro(html, id_pai, id_filho):
    """O elemento `id_filho` está dentro do elemento `id_pai`?"""
    class P(HTMLParser):
        def __init__(self):
            super().__init__()
            self.pilha, self.achou = [], False

        def handle_starttag(self, tag, attrs):
            a = dict(attrs)
            if tag in ("input", "br", "img", "link", "meta"):
                if a.get("id") == id_filho and id_pai in [i for _, i in self.pilha]:
                    self.achou = True
                return
            if a.get("id") == id_filho and id_pai in [i for _, i in self.pilha]:
                self.achou = True
            self.pilha.append((tag, a.get("id")))

        def handle_endtag(self, tag):
            for k in range(len(self.pilha) - 1, -1, -1):
                if self.pilha[k][0] == tag:
                    del self.pilha[k:]
                    break
    p = P()
    p.feed(html)
    return p.achou


# ── 1. Sem internet: nada de CDN, nada de fonte de fora ──

def test_nenhuma_tela_busca_recurso_fora_do_programa():
    assert len(paginas()) >= 8
    for p in paginas():
        html = open(p, encoding="utf-8").read()
        externos = [u for u in re.findall(r'(?:src|href)="(https?://[^"]+)"', html)]
        assert externos == [], (os.path.basename(p), externos)
        assert "cdn.tailwindcss.com" not in html and "fonts.googleapis.com" not in html, p
        assert "/static/tema.css" in html, p


def test_tailwind_local_tem_todas_as_classes_do_login_e_do_admin():
    """Quem acrescentar uma classe do Tailwind nessas telas precisa rodar scripts/gerar_tailwind_local.js — senão ela não tem efeito."""
    css = ler("tailwind_local.css")
    proprias = {"glass", "btn-primary", "campo-segredo"}          # classes da própria página, não do Tailwind
    classes = set()
    for nome in ("login.html", "admin.html"):
        html = ler(nome)
        assert '/static/tailwind_local.css' in html
        for c in re.findall(r'class="([^"]+)"', html):
            classes.update(c.split())
    for linha in ler("admin.js").splitlines():          # classes postas por código: todas as strings das linhas com className
        if "className" in linha:
            for c in re.findall(r"'([^']+)'", re.sub(r"===\s*'[^']*'", "", linha)):
                classes.update(c.split())
    assert len(classes) > 40
    faltando = []
    for c in sorted(classes - proprias):
        seletor = "." + re.sub(r"([^a-zA-Z0-9_-])", r"\\\1", c)
        if seletor not in css:
            faltando.append(c)
    assert faltando == [], faltando


def test_nenhuma_tag_do_admin_tem_class_repetido():
    """`class="a" ... class="b"` na mesma tag: o navegador ignora o segundo (era o que deixava o campo Senha sem estilo)."""
    for nome in ("admin.html", "login.html"):
        for tag in re.findall(r"<[a-zA-Z][^>]*>", ler(nome)):
            assert len(re.findall(r'\sclass="', tag)) <= 1, tag


def test_fontes_locais_e_tema_servidos(tmp_path, monkeypatch):
    tema = ler("tema.css")
    assert tema.count("@font-face") == 2 and "inter-latin.woff2" in tema and "jetbrains-mono-latin.woff2" in tema
    assert ":focus-visible" in tema
    for f in ("inter-latin.woff2", "jetbrains-mono-latin.woff2"):
        with open(os.path.join(ESTATICOS, "fontes", f), "rb") as fh:
            assert fh.read(4) == b"wOF2", f
    assert "SIL Open Font License" in open(os.path.join(ESTATICOS, "fontes", "LICENCA.txt"), encoding="utf-8").read()
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    import app as appmod
    c = TestClient(appmod.app)
    for rota in ("/static/tema.css", "/static/tailwind_local.css", "/static/fontes/inter-latin.woff2", "/static/fontes/jetbrains-mono-latin.woff2"):
        assert c.get(rota).status_code == 200, rota


# ── 2. Hierarquia dos botões ──

def _regra(css, seletor):
    m = re.search(re.escape(seletor) + r"\s*\{([^}]*)\}", css)
    assert m, seletor
    return m.group(1)


def test_base_e_resultado_nao_herdam_o_degrade_do_styles_css():
    """O .btn-action do styles.css (botão Extrair) tem degradê, sombra e pulo no hover; as duas telas precisam anular isso."""
    assert "linear-gradient" in _regra(ler("styles.css"), ".btn-action")
    for nome in ("orcamento.html", "resultado_orcamento.html"):
        html = ler(nome)
        base = _regra(html, "        .btn-action")
        assert "background-image: none" in base and "box-shadow: none" in base, nome
        hover = _regra(html, "        .btn-action:hover")
        assert "transform: none" in hover and "box-shadow: none" in hover, nome
        assert "--btn-primary-text: #161b22" in html, nome


def test_base_de_orcamento_um_principal_e_exclusao_em_vermelho():
    html = ler("orcamento.html")
    i = html.index('<div class="toolbar-actions">')
    barra = html[i:html.index("</header>", i)]
    botoes = re.findall(r'<button[^>]*class="([^"]+)"[^>]*onclick="([^"]+)"', barra)
    principais = [oc for cl, oc in botoes if "btn-save" in cl.split()]
    assert principais == ["salvarAlteracoes()"]
    perigo = [oc for cl, oc in botoes if "btn-danger" in cl.split()]
    assert perigo == ["BaseGrade.excluirSelecionadas()"]
    for idv in ("btn-admin-add", "btn-admin-csv", "btn-admin-sync-all", "btn-admin-migrar"):          # ações de administrador ficam num grupo à parte
        m = re.search(r'<span class="grupo-admin"[\s\S]*?</span>\s*\n\s*\n', barra)
        assert m and f'id="{idv}"' in m.group(0), idv
    assert ".grupo-admin:not(:has(button:not(.hidden)))" in html


def test_styles_css_contraste_e_hover():
    css = ler("styles.css")
    assert "color: #161b22" in _regra(css, ".btn-primary")
    hover = _regra(css, ".btn-secondary:hover")
    assert "#ff7b72" not in hover          # passar o mouse num botão comum não fica mais vermelho
    assert "#f85149" in _regra(css, ".btn-danger, .btn-secondary.btn-danger")


def test_tela_inicial_sem_botoes_claros_e_sem_quebra_de_texto():
    html = ler("index.html")
    for cor in ("#fef08a", "#fee2e2"):
        assert cor not in html
    i = html.index('<header class="inicio-topo">')
    topo = html[i:html.index("</header>", i)]
    for idv in ("btn-admin", "btn-logout", "btn-encerrar", "select-projeto-leitor", "btn-regras-leitor", "btn-recs", "btn-cloudconvert", "btn-plugin-nanocad"):
        assert len(re.findall(r'\bid="%s"' % idv, topo)) == 1, idv
    assert re.search(r'id="btn-logout" class="[^"]*btn-danger', topo)
    assert "white-space: nowrap" in _regra(ler("styles.css"), ".btn-topo")
    sessao = topo[topo.index("inicio-sessao"):topo.index("</div>", topo.index("inicio-sessao"))]
    assert sessao.index('id="btn-logout"') < sessao.index('class="inicio-sep"') < sessao.index('id="btn-encerrar"')          # Sair e Encerrar separados


# ── 3. Resumo: menus "Mais" e "Exibição" ──

def test_resumo_menu_mais_e_menu_exibicao():
    html = ler("resumo.html")
    for idv in ("btn-save-modelo", "btn-importar-json", "btn-exportar-json", "btn-imprimir"):
        assert dentro(html, "menu-obra-itens", idv), idv
        assert dentro(html, "titlebar", idv), idv
    for idv in ("btn-ler-projeto", "btn-montar-orcamento", "btn-load-obra", "btn-save-obra"):          # o que se usa sempre continua na barra
        assert dentro(html, "titlebar", idv) and not dentro(html, "menu-obra-itens", idv), idv
    for idv in ("chip-tabela-cabos", "chip-tabela-outros", "chip-sugestoes", "seg-layout", "btn-inverter-tabelas"):
        assert dentro(html, "menu-exibicao-itens", idv), idv
        assert dentro(html, "ctlbar", idv), idv
    assert re.search(r'id="btn-menu-obra"[^>]*data-menu="menu-obra-itens"', html)
    assert re.search(r'id="btn-menu-exibicao"[^>]*data-menu="menu-exibicao-itens"', html)
    assert re.search(r'id="menu-obra-itens"[^>]*\bhidden\b', html) and re.search(r'id="menu-exibicao-itens"[^>]*\bhidden\b', html)


def test_resumo_limpar_tudo_longe_do_salvar():
    html = ler("resumo.html")
    i = html.index('id="ctlbar"')
    barra = html[i:html.index('<div class="work"', i)]
    assert barra.index('id="btn-limpar-tabelas"') < barra.rindex('class="tb-flex"') < barra.index('id="btn-salvar-alteracoes"')
    assert re.search(r'class="[^"]*btn-danger[^"]*" id="btn-limpar-tabelas"', barra)


def test_layout_js_e_css_dos_menus():
    js, css = ler("resumo_layout.js"), ler("resumo_layout.css")
    for trecho in ["fecharMenus", "[data-menu]", "data-fecha-menu", "tem-oculta", "'Escape'"]:
        assert trecho in js, trecho
    for trecho in [".menu-susp-itens[hidden] { display: none; }", ".menu-exibicao", "#btn-menu-exibicao.tem-oculta"]:
        assert trecho in css, trecho


# ── 4. Rodapé, manual, autônomo, textos ──

def test_rodape_nao_cobre_o_conteudo():
    marca = ler("marca.js")
    assert "dataset.rodape" in marca and "'esquerda'" in marca and "'fino'" in marca
    assert '<body data-rodape="fino">' in ler("resumo.html")
    assert '<body data-rodape="esquerda">' in ler("orcamento.html")


def test_manual_escuro_na_tela_e_claro_no_papel():
    html = ler("manual.html")
    assert ":root { --bg:#0d1117;" in html and "prefers-color-scheme" not in html
    assert re.search(r"@media print \{ :root \{ --bg:#fff;", html)


def test_autonomo_quadros_em_uma_linha_e_aviso_de_orientacao():
    html, js = ler("autonomo.html"), ler("autonomo.js")
    assert "minmax(130px,1fr)" in html and ".msg.atencao" in html
    assert "'atencao')" in js


def test_textos_sem_jargao_e_nomes_iguais():
    assert "Nenhum payload" not in ler("resultado_orcamento.html")
    assert "Nenhum orçamento montado ainda." in ler("resultado_orcamento.html")
    manual = open(os.path.join(config.BASE_DIR, "data", "manual_regras_e_ajustes.md"), encoding="utf-8").read()
    assert "Leitor de Projetos" not in manual and "Resumo do projeto" not in manual
    for nome, titulo in (("admin.html", "Painel Admin — OrçaCore"), ("autonomo.html", "Modo autônomo — OrçaCore"), ("manual.html", "Manual: regras e ajustes — OrçaCore")):
        assert f"<title>{titulo}</title>" in ler(nome), nome
