"""TASK-127 — atualização do .exe pelas Releases do GitHub: comparação de versões, leitura da Release, download com conferência de SHA-256, .bat de troca e rotas."""
import hashlib
import http.server
import os
import threading

import pytest
from fastapi.testclient import TestClient

import app as app_mod
from services import auto_updater as au

REPO = au.GITHUB_REPO
BASE = f"https://github.com/{REPO}/releases/download/v9.9.9/"


def release(tag="v9.9.9", exe=True, sha=True, **extra):
    ativos = []
    if exe:
        ativos.append({"name": au.NOME_EXE, "browser_download_url": BASE + au.NOME_EXE})
    if sha:
        ativos.append({"name": au.NOME_SHA, "browser_download_url": BASE + au.NOME_SHA})
    return {"tag_name": tag, "body": "Correções", "assets": ativos, **extra}


@pytest.fixture(autouse=True)
def limpo(monkeypatch):
    au._cache.update(t=0.0, dados=None)
    monkeypatch.setattr(au, "APP_MODE", "desktop")
    monkeypatch.delenv("ATUALIZADOR", raising=False)
    monkeypatch.setattr(au, "APP_VERSION", "2.0.0")
    yield
    au._cache.update(t=0.0, dados=None)


# ───────────────────────── versões e Release ─────────────────────────
@pytest.mark.parametrize("a,b", [("v2.1.0", (2, 1, 0)), ("2.1", (2, 1, 0)), ("V10.0.3", (10, 0, 3)), ("2", (2, 0, 0))])
def test_versao_tupla(a, b):
    assert au.versao_tupla(a) == b


@pytest.mark.parametrize("a", ["", None, "abc", "2.x", "v2.1.0-beta"])
def test_versao_invalida(a):
    assert au.versao_tupla(a) is None


def test_release_mais_nova_com_exe_e_sha_oferece():
    d = au._interpretar(release())
    assert d["has_update"] and d["latest_version"] == "9.9.9" and d["download_url"].endswith(au.NOME_EXE) and d["sha256_url"].endswith(au.NOME_SHA)


@pytest.mark.parametrize("r", [release("v2.0.0"), release("v1.9.9"), release(draft=True), release(prerelease=True), release(exe=False), release(sha=False)])
def test_release_que_nao_serve_nao_oferece(r):
    assert au._interpretar(r)["has_update"] is False


def test_url_de_fora_do_repositorio_e_recusada():
    r = release()
    r["assets"][0]["browser_download_url"] = "https://exemplo.com/Leitor_Projetos.exe"
    assert au._interpretar(r)["has_update"] is False


def test_check_usa_cache_e_nao_quebra_com_erro(monkeypatch):
    chamadas = []
    monkeypatch.setattr(au, "_get_json", lambda url: chamadas.append(url) or release())
    assert au.check_for_updates()["has_update"] and au.check_for_updates()["has_update"]
    assert len(chamadas) == 1 and REPO in chamadas[0]
    au._cache.update(t=0.0, dados=None)

    def falha(url):
        raise OSError("sem rede")
    monkeypatch.setattr(au, "_get_json", falha)
    assert au.check_for_updates() == {"has_update": False, "motivo": "erro na consulta"}


def test_servidor_e_atualizador_desligado_nao_consultam(monkeypatch):
    monkeypatch.setattr(au, "_get_json", lambda url: pytest.fail("não devia consultar"))
    monkeypatch.setattr(au, "APP_MODE", "server")
    assert au.check_for_updates() == {"has_update": False}
    monkeypatch.setattr(au, "APP_MODE", "desktop")
    monkeypatch.setenv("ATUALIZADOR", "0")
    assert au.check_for_updates() == {"has_update": False}


def test_offline_nao_consulta(monkeypatch):
    from services import rede_adaptativa
    monkeypatch.setattr(rede_adaptativa, "offline", lambda: True)
    monkeypatch.setattr(au, "_get_json", lambda url: pytest.fail("não devia consultar"))
    assert au.check_for_updates()["has_update"] is False


# ───────────────────────── download e troca ─────────────────────────
@pytest.fixture
def servidor():
    """Servidor local que serve o 'exe' e o .sha256 (o teste troca o prefixo permitido para ele)."""
    conteudo = b"MZ" + os.urandom(2048)
    sha = hashlib.sha256(conteudo).hexdigest()
    estado = {"sha": sha}

    class H(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            corpo = conteudo if self.path.endswith(".exe") else f"{estado['sha']}  Leitor_Projetos.exe\n".encode()
            self.send_response(200); self.send_header("Content-Length", str(len(corpo))); self.end_headers(); self.wfile.write(corpo)

        def log_message(self, *a):
            pass

    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{srv.server_address[1]}/", conteudo, estado
    srv.shutdown()


@pytest.fixture
def instalacao(tmp_path, monkeypatch, servidor):
    base, conteudo, estado = servidor
    exe = tmp_path / "Leitor_Projetos.exe"
    exe.write_bytes(b"MZ-versao-antiga")
    monkeypatch.setattr(au, "IS_FROZEN", True)
    monkeypatch.setattr(au.sys, "executable", str(exe))
    monkeypatch.setattr(au, "_url_permitida", lambda: base)
    monkeypatch.setattr(au, "_get_json", lambda url: {"tag_name": "v9.9.9", "assets": [
        {"name": au.NOME_EXE, "browser_download_url": base + au.NOME_EXE}, {"name": au.NOME_SHA, "browser_download_url": base + au.NOME_SHA}]})
    lancados, encerrou = [], []
    monkeypatch.setattr(au, "_iniciar_script", lambda bat: lancados.append(bat))
    monkeypatch.setattr(au, "_encerrar", lambda: encerrou.append(True))
    monkeypatch.setattr(au.threading, "Timer", lambda t, f: type("T", (), {"start": lambda self: f()})())
    return exe, conteudo, estado, lancados, encerrou


def test_aplica_baixa_confere_e_agenda_a_troca(instalacao):
    exe, conteudo, estado, lancados, encerrou = instalacao
    assert au.download_and_install_update("https://malicioso.exemplo/virus.exe") is True          # a URL do pedido é ignorada
    novo = exe.with_name(exe.name + ".novo")
    assert novo.read_bytes() == conteudo and exe.read_bytes() == b"MZ-versao-antiga"             # o .exe atual só é trocado pelo .bat, depois de fechar
    assert len(lancados) == 1 and encerrou == [True]
    bat = open(lancados[0], encoding="utf-8").read()
    assert f'set "PID={os.getpid()}"' in bat and str(novo) in bat and str(exe) + ".antigo" in bat and "PYINSTALLER_RESET_ENVIRONMENT" in bat and 'start "" "%ALVO%"' in bat


def test_sha_diferente_cancela_e_apaga_o_baixado(instalacao):
    exe, _, estado, lancados, encerrou = instalacao
    estado["sha"] = "0" * 64
    with pytest.raises(RuntimeError, match="SHA-256"):
        au.download_and_install_update()
    assert not exe.with_name(exe.name + ".novo").exists() and lancados == [] and encerrou == []


def test_nao_aplica_sem_versao_nova_nem_fora_do_exe(instalacao, monkeypatch):
    exe, *_ = instalacao
    monkeypatch.setattr(au, "APP_VERSION", "9.9.9")
    with pytest.raises(RuntimeError, match="Não há atualização"):
        au.download_and_install_update()
    monkeypatch.setattr(au, "IS_FROZEN", False)
    with pytest.raises(RuntimeError, match="só no executável"):
        au.download_and_install_update()


def test_script_de_troca_tem_a_ordem_certa():
    bat = au.gerar_script(123, r"C:\App\Leitor_Projetos.exe", r"C:\App\Leitor_Projetos.exe.novo", r"C:\App\Leitor_Projetos.exe.antigo")
    ordem = [bat.index(x) for x in (":esperar", ":trocar", 'move /y "%ALVO%" "%ANTIGO%"', 'move /y "%NOVO%" "%ALVO%"', 'start "" "%ALVO%"')]
    assert ordem == sorted(ordem) and 'del /f /q "%~f0"' in bat


def test_limpar_residuos_apaga_antigo_e_novo(tmp_path, monkeypatch):
    exe = tmp_path / "Leitor_Projetos.exe"
    for sufixo in ("", ".antigo", ".novo"):
        (tmp_path / ("Leitor_Projetos.exe" + sufixo)).write_bytes(b"x")
    (tmp_path / "atualizar_leitor.bat").write_text("x")
    monkeypatch.setattr(au, "IS_FROZEN", True)
    monkeypatch.setattr(au.sys, "executable", str(exe))
    au.limpar_residuos(atraso=0.0)
    import time
    time.sleep(0.4)
    assert exe.exists() and not (tmp_path / "Leitor_Projetos.exe.antigo").exists() and not (tmp_path / "Leitor_Projetos.exe.novo").exists() and not (tmp_path / "atualizar_leitor.bat").exists()


# ───────────────────────── rotas ─────────────────────────
def test_rota_check_desktop(monkeypatch):
    monkeypatch.setattr("routers.update.check_for_updates", lambda: {"has_update": True, "latest_version": "9.9.9"})
    c = TestClient(app_mod.app)
    assert c.get("/api/update/check").json() == {"has_update": True, "latest_version": "9.9.9"}


def test_rota_apply_ignora_a_url_do_corpo(monkeypatch):
    recebido = []
    monkeypatch.setattr("routers.update.APP_MODE", "desktop")
    monkeypatch.setattr("routers.update.download_and_install_update", lambda *a: recebido.append(a) or True)
    from middleware.auth_middleware import create_jwt_token
    c = TestClient(app_mod.app)
    h = {"Authorization": "Bearer " + create_jwt_token("u1", "a@x.com", "admin")}
    r = c.post("/api/update/apply", json={"download_url": "https://malicioso.exemplo/x.exe"}, headers=h)
    assert r.status_code == 200 and recebido == [()]
    monkeypatch.setattr("routers.update.download_and_install_update", lambda *a: (_ for _ in ()).throw(RuntimeError("só no executável")))
    r = c.post("/api/update/apply", json={}, headers=h)
    assert r.status_code == 400 and "só no executável" in r.json()["detail"]


def test_apply_sem_login_e_recusado():
    assert TestClient(app_mod.app).post("/api/update/apply", json={}).status_code in (401, 403)
