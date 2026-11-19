"""TASK-193 — (1) o programa desktop abre na tela inicial e a sessão lembrada vale; (2) «Abrir» não lê o arquivo na ponte: devolve um id e os bytes saem por HTTP local."""
import os
import re

import pytest
from fastapi.testclient import TestClient

import app as appmod
import config
from middleware.auth_middleware import create_jwt_token
from services import arquivos_desktop


def _ler(rel):
    with open(os.path.join(os.path.dirname(config.STATIC_DIR), rel), encoding="utf-8") as f:
        return f.read()


def cab():
    return {"Authorization": f"Bearer {create_jwt_token('u1', 'a@x.com', 'operador')}"}


@pytest.fixture(autouse=True)
def _sem_registro():
    arquivos_desktop._escolhidos.clear()
    yield
    arquivos_desktop._escolhidos.clear()


# ── abertura sem apagar a sessão ────────────────────────────────────────────────────────────────────
def test_o_programa_nao_abre_mais_no_login_que_apaga_a_sessao():
    src = _ler("app.py")
    assert "inicio_url = f\"{app_url}/\"" in src and "win.load_url(inicio_url)" in src and "webbrowser.open(inicio_url)" in src
    assert "login?new_session" not in src and "login_url" not in src                    # era o ?new_session=1 que limpava o «Lembrar de mim» a cada abertura
    assert "programa_local.url_pendente(_id_obra)" in src                         # obra do nanoCAD: abre direto na tela da obra (ela pede o login com ?next= se preciso)


def test_login_com_sessao_lembrada_vai_direto_e_new_session_continua_limpando():
    html = _ler("static/login.html")
    assert "localStorage.getItem('auth_lembrar') === 'true'" in html and "q.get('new_session') !== '1'" in html and "window.location.replace(" in html
    assert html.index("window.location.replace(") < html.index("if (params.get('new_session') === '1')")       # decide antes de limpar
    assert "localStorage.removeItem('auth_token')" in html                                                     # ?new_session=1 (link manual) segue limpando


# ── bytes por HTTP local ────────────────────────────────────────────────────────────────────────────
def test_registro_so_serve_o_que_o_usuario_escolheu(tmp_path):
    f = tmp_path / "nota.dxf"
    f.write_bytes(b"\x00\x01DXF\xff" * 1000)
    i = arquivos_desktop.registrar_escolhido(str(f))
    assert len(i) >= 32 and arquivos_desktop.caminho_escolhido(i) == str(f)
    assert arquivos_desktop.caminho_escolhido("outro-id") is None and arquivos_desktop.caminho_escolhido("") is None and arquivos_desktop.caminho_escolhido(None) is None
    f.unlink()
    assert arquivos_desktop.caminho_escolhido(i) is None                          # sumiu do disco


def test_id_vence_em_10_minutos(tmp_path, monkeypatch):
    f = tmp_path / "a.pdf"
    f.write_bytes(b"%PDF")
    t = [1000.0]
    monkeypatch.setattr(arquivos_desktop.time, "monotonic", lambda: t[0])
    i = arquivos_desktop.registrar_escolhido(str(f))
    t[0] += arquivos_desktop.VALIDADE_ESCOLHIDOS_S - 1
    assert arquivos_desktop.caminho_escolhido(i) == str(f)
    t[0] += 2
    assert arquivos_desktop.caminho_escolhido(i) is None
    arquivos_desktop.registrar_escolhido(str(f))
    assert len(arquivos_desktop._escolhidos) == 1                                 # o vencido foi varrido


def test_rota_entrega_os_bytes_exatos_e_exige_login(tmp_path):
    conteudo = bytes(range(256)) * 4096                                           # 1 MB binário
    f = tmp_path / "grande.dxf"
    f.write_bytes(conteudo)
    i = arquivos_desktop.registrar_escolhido(str(f))
    c = TestClient(appmod.app)
    r = c.get(f"/api/desktop/arquivo/{i}", headers=cab())
    assert r.status_code == 200 and r.content == conteudo and "no-store" in r.headers["cache-control"]
    assert c.get(f"/api/desktop/arquivo/{i}").status_code == 401                  # sem login não serve
    assert c.get("/api/desktop/arquivo/inventado", headers=cab()).status_code == 404
    assert c.get("/api/desktop/arquivo/..%2F..%2Fetc%2Fpasswd", headers=cab()).status_code == 404            # só id registrado, nunca caminho


def test_abrir_arquivos_nao_le_o_arquivo(tmp_path, monkeypatch):
    f = tmp_path / "x.dxf"
    f.write_bytes(b"a" * 5000)

    class Janela:
        def create_file_dialog(self, *a, **k):
            return (str(f),)
    import types
    monkeypatch.setitem(__import__("sys").modules, "webview", types.SimpleNamespace(FileDialog=types.SimpleNamespace(OPEN="open", SAVE="save")))
    api = arquivos_desktop.ApiArquivos()
    api.vincular(Janela())
    lido = []
    real = open
    monkeypatch.setattr("builtins.open", lambda p, *a, **k: (lido.append(str(p)), real(p, *a, **k))[1])
    r = api.abrir_arquivos(".dxf", False)
    assert r["ok"] and r["arquivos"][0]["tamanho"] == 5000 and r["arquivos"][0]["nome"] == "x.dxf" and "base64" not in r["arquivos"][0]
    assert str(f) not in lido                                                     # a ponte não abre o arquivo: quem lê é a rota HTTP


def test_js_mostra_abrindo_na_hora_e_busca_por_http():
    js = _ler("static/desktop_arquivos.js")
    corpo = js[js.index("async function escolher"):js.index("const cliqueOriginalInput")]
    assert corpo.index("mostrarAbrindo(r.arquivos)") < corpo.index("raiz.fetch('/api/desktop/arquivo/'")        # aviso antes de buscar
    assert "resp.blob()" in corpo and "deBase64" not in corpo and "aviso.remove()" in corpo
    assert re.search(r"finally\s*{\s*aviso\.remove\(\);", corpo)                  # o aviso sempre some, mesmo se der erro
