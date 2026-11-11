"""TASK-133 — fechar a janela pede confirmação e encerra o servidor; login aceita as senhas salvas."""
import threading
from pathlib import Path

from services import janela_desktop as jd

RAIZ = Path(__file__).resolve().parent.parent


def test_fechar_pede_confirmacao_em_portugues_e_encerra_o_servidor():
    app_py = (RAIZ / "app.py").read_text(encoding="utf-8")
    assert "confirm_close=True" in app_py
    assert "localization=janela_desktop.LOCALIZACAO" in app_py
    assert "janela_desktop.encerrar_servidor(_servidor_uvicorn, server_thread)" in app_py
    assert "encerrado" in jd.LOCALIZACAO["global.quitConfirmation"] and jd.LOCALIZACAO["global.cancel"] == "Cancelar"


def test_encerrar_para_o_uvicorn_e_sai_do_processo(monkeypatch):
    saiu = []
    monkeypatch.setattr(jd.os, "_exit", lambda c: saiu.append(c))

    class Servidor:
        should_exit = False

    s = Servidor()
    t = threading.Thread(target=lambda: None)
    t.start()
    jd.encerrar_servidor(s, t, espera=1)
    assert s.should_exit is True and saiu == [0]


def test_perfil_persistente_e_sem_webview2_nao_quebra(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    p = jd.pasta_perfil()
    assert Path(p).is_dir() and p.startswith(str(tmp_path))

    class SemNativo:
        native = None

    assert jd.ligar_senhas_salvas(SemNativo()) is False
    assert "storage_path=janela_desktop.pasta_perfil()" in (RAIZ / "app.py").read_text(encoding="utf-8")


def test_login_preparado_para_senhas_salvas():
    html = (RAIZ / "static" / "login.html").read_text(encoding="utf-8")
    assert 'name="email" autocomplete="username"' in html and 'name="password" autocomplete="current-password"' in html
    js = (RAIZ / "static" / "login.js").read_text(encoding="utf-8")
    assert "navigator.credentials.store" in js and "navigator.credentials.get" in js


def test_perfil_do_nome_anterior_e_aproveitado(tmp_path, monkeypatch):
    """TASK-135: ReadGrid → GridReader sem perder as senhas salvas da janela."""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    antiga = tmp_path / "ReadGrid" / "janela"
    antiga.mkdir(parents=True)
    (antiga / "Login Data").write_text("x")
    p = Path(jd.pasta_perfil())
    assert p == tmp_path / "GridReader" / "janela" and (p / "Login Data").read_text() == "x"
    assert not antiga.exists()
    assert "GridReader" in jd.LOCALIZACAO["global.quitConfirmation"]
