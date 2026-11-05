"""TASK-111 — novo layout da tela Resumo (barra de título, barra de modais, painéis laterais, linhas vazias estilo Excel): estrutura estática,
lógica pura da grade (operação herdada / colagem mista) e a rota que exporta a obra da tela como .obra.json."""
import json
import os
import re
import subprocess
from html.parser import HTMLParser

import pytest
from fastapi.testclient import TestClient

import config
import database
from middleware.auth_middleware import create_jwt_token
from services import obras_arquivo
from tests import oraculo_tela as o


def _ler(nome):
    with open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8") as f:
        return f.read()


HTML = _ler("resumo.html")


def _dentro(id_pai, id_filho):
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
            self.pilha.append((tag, a.get("id")))
            if a.get("id") == id_filho and id_pai in [i for _, i in self.pilha[:-1]]:
                self.achou = True
        def handle_endtag(self, tag):
            for k in range(len(self.pilha) - 1, -1, -1):
                if self.pilha[k][0] == tag:
                    del self.pilha[k:]
                    break
    p = P()
    p.feed(HTML)
    return p.achou


def test_ids_continuam_unicos_e_presentes():
    ids = ["input-num-obra", "btn-ler-projeto", "btn-montar-orcamento", "btn-load-obra", "btn-save-obra", "btn-save-modelo", "btn-importar-obra", "input-importar-obra",
           "btn-ajustar", "btn-validar", "btn-regras-validacao", "select-contexto", "select-projeto", "btn-tabela-orcamento", "btn-abrir-pesquisa", "btn-config-api",
           "chat-messages", "chat-input", "btn-send-chat", "voz-estado", "btn-undo", "btn-redo", "barra-modais-montaveis", "btn-gerar-totalizadora", "btn-limpar-tabelas",
           "btn-salvar-alteracoes", "section-cabos", "section-outros", "view-padrao", "view-totalizadora", "body-cabos", "body-outros", "table-cabos", "table-outros",
           "resumo-rede-postes", "resumo-rede-media", "resumo-rede-baixa", "resumo-rede-equipamentos", "modal-load-obra", "modal-save-obra", "modal-validacao"]
    for i in ids:
        assert len(re.findall(r'\bid="%s"' % re.escape(i), HTML)) == 1, i


def test_html_sem_tags_desbalanceadas():
    tags = {"div", "aside", "main", "header", "nav", "section", "table", "tbody", "thead", "tr", "td", "th", "select", "button", "label"}

    class P(HTMLParser):
        def __init__(self):
            super().__init__()
            self.pilha, self.erros = [], []
        def handle_starttag(self, tag, attrs):
            if tag in tags:
                self.pilha.append((tag, self.getpos()))
        def handle_endtag(self, tag):
            if tag not in tags:
                return
            if self.pilha and self.pilha[-1][0] == tag:
                self.pilha.pop()
            else:
                self.erros.append((tag, self.getpos()))
    p = P()
    p.feed(HTML)
    assert p.erros == [] and p.pilha == []


def test_barra_de_titulo():
    for i in ["input-num-obra", "btn-ler-projeto", "btn-montar-orcamento", "btn-load-obra", "btn-save-obra", "btn-save-modelo", "btn-importar-json", "btn-exportar-json",
              "btn-undo", "btn-redo", "btn-toggle-config", "btn-toggle-ia"]:
        assert _dentro("titlebar", i), i


def test_barra_de_modais():
    for i in ["btn-resumo-rede", "btn-ajustar", "btn-validar", "select-contexto", "btn-regras-validacao", "barra-modais-montaveis"]:
        assert _dentro("barra-modais", i), i
    for titulo in ["Gerador de Postes e Estruturas", "Gerador de Cabos", "Gerador de Conexões", "Vincular cabos a estruturas/postes"]:
        assert f'title="{titulo}"' in HTML


def test_paineis_laterais_ocultos_por_padrao_com_o_conteudo_certo():
    assert re.search(r'<aside class="painel painel-esq" id="painel-config"[^>]*\bhidden\b', HTML)
    assert re.search(r'<aside class="painel painel-dir" id="painel-ia"[^>]*\bhidden\b', HTML)
    for i in ["btn-tabela-orcamento", "select-projeto", "btn-abrir-pesquisa", "btn-config-api"]:
        assert _dentro("painel-config", i), i
    for i in ["chat-messages", "chat-input", "btn-send-chat", "voz-estado"]:
        assert _dentro("painel-ia", i), i
    assert not _dentro("titlebar", "btn-tabela-orcamento")


def test_resumo_compacto_fica_na_barra_de_controles():
    for i in ["resumo-rede-postes", "resumo-rede-media", "resumo-rede-baixa", "resumo-rede-equipamentos", "resumo-compacto"]:
        assert _dentro("ctlbar", i), i
    assert 'id="modal-resumo-rede"' in HTML and 'id="btn-resumo-rede"' in HTML


def test_css_js_do_layout_ligados_e_servidos(tmp_path, monkeypatch):
    assert '/static/resumo_layout.css' in HTML and '/static/resumo_layout.js' in HTML
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    import app as appmod
    c = TestClient(appmod.app)
    r = c.get("/static/resumo_layout.js")
    assert r.status_code == 200 and "javascript" in r.headers["content-type"] and "no-cache" in r.headers.get("cache-control", "")
    assert c.get("/static/resumo_layout.css").status_code == 200


def test_layout_js_e_css_trazem_os_comportamentos():
    js, css = _ler("resumo_layout.js"), _ler("resumo_layout.css")
    for trecho in ["resumo_layout_v1", "definirPainel", "flutuante", "empilhadas", "btn-recolher", "Escape", "modaisRecolhida"]:
        assert trecho in js, trecho
    for trecho in [".painel.flutuante", ".tables-row.empilhadas", ".barra-modais.recolhida", ".linha-vazia", ".report-section.oculta"]:
        assert trecho in css, trecho
    assert ".tables-row .table-wrap { flex: 1; min-height: 0; max-height: none;" in css          # as tabelas passam a usar a altura da tela


def test_resumo_js_tem_linhas_vazias_e_exportar():
    js = _ler("resumo.js")
    for trecho in ["ajustarLinhasVazias", "materializarLinhaVazia", "tbody-vazio", "inp-ativo-vazio", "btn-exportar-json", "/api/obras/exportar-atual", "operacaoHerdada"]:
        assert trecho in js, trecho
    assert "LINHAS_VAZIAS_MIN" not in js          # constante no mesmo escopo seria lida antes de existir na inicialização (TDZ)


# ── lógica pura da grade ──
MODULO = os.path.join(config.STATIC_DIR, "grade_tabela.js")
node = pytest.mark.skipif(not o.NODE, reason="node não instalado")


def js(expr, **entrada):
    codigo = f"const G=require({json.dumps(MODULO)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({expr}));"
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def L(op, ativo, **kw):
    return {"entidade": "0", "operacao": op, "ativo": ativo, **kw}


@node
def test_operacao_herdada():
    f = lambda d, i=None: js("G.operacaoHerdada(e.d, e.i)", d=d, i=i)
    assert f([]) == "I" and f([], 0) == "I"                                         # tabela vazia: I
    assert f([L("*R", "A")]) == "*R"                                                 # copia a de cima
    assert f([L("R", "A"), L("M", "B")], 1) == "R"                                   # inserir na posição 1: a linha preenchida acima dela
    assert f([L("R", "A"), L("M", "")]) == "R"                                       # linha em branco não conta como "preenchida"
    assert f([L("*I", ""), L("R", "  ")]) == "I"                                     # só linhas em branco acima
    assert f([L("zz", "A")]) == "I"                                                  # operação inválida na de cima: I


@node
def test_colagem_herda_operacao_de_cima_e_aceita_linhas_mistas():
    r = js("(()=>{const d=e.d;const r=G.aplicarColagem(d,e.t,{tipo:'outros',linha:1,coluna:1});return {d:d.map(x=>x.operacao+':'+x.ativo),n:r.criadas};})()",
           d=[L("R", "N1")], t="9200D\n*I\t7150FV\nCE3\n\tX")
    assert r["d"] == ["R:N1", "R:9200D", "*I:7150FV", "*I:CE3", "*I:X"] and r["n"] == 4
    vazio = js("(()=>{const d=[];G.aplicarColagem(d,e.t,{tipo:'cabos',linha:0,coluna:1});return d.map(x=>x.operacao+':'+x.ativo);})()", t="A\nB\nM\tC\nD")
    assert vazio == ["I:A", "I:B", "M:C", "M:D"]                                    # tabela vazia: I; depois cada linha herda a de cima


@node
def test_colagem_com_so_operacoes_e_ativos_de_antes_continua_igual():
    r = js("(()=>{const d=[];G.aplicarColagem(d,e.t,{tipo:'outros',linha:0,coluna:1});return d.map(x=>x.operacao+':'+x.ativo);})()", t="*R\tA 1\nR\tB 2")
    assert r == ["*R:A 1", "R:B 2"]


# ── exportar a obra da tela ──
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    import app as appmod
    return TestClient(appmod.app)


def _auth():
    return {"Authorization": f"Bearer {create_jwt_token('u1', 'a@x.com', 'admin')}"}


def test_exportar_obra_atual_gera_arquivo_importavel(client):
    estados = {"cabos": {"data": [{"entidade": "CABO", "operacao": "I", "ativo": "3XA35 30 m", "qtdAtivos": 3, "lixo": "x"}]},
               "outros": {"data": [{"entidade": "0", "operacao": "*R", "ativo": "9200D, N1"}]}, "totalizadora": {"data": []}}
    r = client.post("/api/obras/exportar-atual?projeto=081", json={"nome": "9101578729", "dados_json": json.dumps(estados)}, headers=_auth())
    assert r.status_code == 200, r.text
    env = r.json()
    assert (env["formato"], env["versao"], env["tipo"], env["nome"], env["projeto"]) == ("leitor-obra", 1, "obra", "9101578729", "081")
    assert env["dados"]["cabos"] == [{"entidade": "CABO", "operacao": "I", "ativo": "3XA35 30 m", "qtdAtivos": 3}]            # campo desconhecido não vai
    assert env["dados"]["outros"][0]["operacao"] == "*R"
    pronta = obras_arquivo.validar_importacao(env, "081")                                                             # o mesmo contrato do Importar
    assert pronta["nome"] == "9101578729" and len(pronta["dados"]["outros"]) == 1


def test_exportar_obra_atual_nome_vazio_e_dados_ruins(client):
    r = client.post("/api/obras/exportar-atual?projeto=081", json={"nome": "", "dados_json": "isto não é json"}, headers=_auth())
    assert r.status_code == 200 and r.json()["nome"] == "obra" and r.json()["dados"]["cabos"] == []


def test_exportar_obra_atual_exige_login_e_limita_tamanho(client):
    assert client.post("/api/obras/exportar-atual?projeto=081", json={"nome": "x", "dados_json": "{}"}).status_code in (401, 403)
    grande = "x" * (obras_arquivo.LIMITE_BYTES + 1)
    assert client.post("/api/obras/exportar-atual?projeto=081", json={"nome": "x", "dados_json": grande}, headers=_auth()).status_code == 413
