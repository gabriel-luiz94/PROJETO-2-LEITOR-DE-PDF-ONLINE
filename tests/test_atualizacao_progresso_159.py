"""TASK-159 — atualização automática com barra de progresso e botão Cancelar (o download roda em segundo plano; a tela consulta o andamento)."""
import hashlib
import http.server
import os
import threading
import time

import pytest
from fastapi.testclient import TestClient

import app as app_mod
import config
from services import auto_updater as au

BASE_FAKE = "http://127.0.0.1/"


@pytest.fixture(autouse=True)
def zera_andamento():
    au._estado.update(fase="ocioso", baixado=0, total=0, erro="")
    au._cancelar.clear()
    yield
    au._estado.update(fase="ocioso", baixado=0, total=0, erro="")
    au._cancelar.clear()


@pytest.fixture
def servidor_lento():
    """Servidor local: manda a 1ª metade do 'exe', espera o teste liberar (`segura`) e manda o resto."""
    conteudo = b"MZ" + os.urandom(600 * 1024)
    sha = hashlib.sha256(conteudo).hexdigest()
    ctrl = {"sha": sha, "segura": threading.Event(), "liberar": threading.Event()}
    ctrl["segura"].set()                                                              # por padrão não segura: manda tudo

    class H(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            if not self.path.endswith(".exe"):
                corpo = f"{ctrl['sha']}  x.exe\n".encode()
                self.send_response(200); self.send_header("Content-Length", str(len(corpo))); self.end_headers(); self.wfile.write(corpo)
                return
            self.send_response(200); self.send_header("Content-Length", str(len(conteudo))); self.end_headers()
            meio = len(conteudo) // 2
            try:
                self.wfile.write(conteudo[:meio]); self.wfile.flush()
                ctrl["liberar"].set()                                                # avisa o teste: a 1ª metade já saiu
                ctrl["segura"].wait(15)
                self.wfile.write(conteudo[meio:])
            except OSError:
                pass

        def log_message(self, *a):
            pass

    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{srv.server_address[1]}/", conteudo, ctrl
    ctrl["segura"].set()
    srv.shutdown()


@pytest.fixture
def instalacao(tmp_path, monkeypatch, servidor_lento):
    base, conteudo, ctrl = servidor_lento
    exe = tmp_path / "OrcaCore.exe"
    exe.write_bytes(b"MZ-versao-antiga")
    monkeypatch.setattr(au, "IS_FROZEN", True)
    monkeypatch.setattr(au, "APP_MODE", "desktop")
    monkeypatch.setattr(au.sys, "executable", str(exe))
    monkeypatch.setattr(au, "_url_permitida", lambda: base)
    monkeypatch.setattr(au, "_get_json", lambda url: {"tag_name": "v9.9.9", "assets": [
        {"name": au.NOME_EXE, "browser_download_url": base + au.NOME_EXE}, {"name": au.NOME_SHA, "browser_download_url": base + au.NOME_SHA}]})
    lancados, encerrou = [], []
    monkeypatch.setattr(au, "_iniciar_script", lambda bat: lancados.append(bat))
    monkeypatch.setattr(au, "_encerrar", lambda: encerrou.append(True))
    monkeypatch.setattr(au.threading, "Timer", lambda t, f: type("T", (), {"start": lambda self: f()})())
    return exe, conteudo, ctrl, lancados, encerrou


def espera(cond, segundos=10):
    fim = time.monotonic() + segundos
    while time.monotonic() < fim:
        if cond():
            return True
        time.sleep(0.02)
    return False


def test_andamento_termina_em_pronto_com_100_por_cento(instalacao):
    exe, conteudo, ctrl, lancados, encerrou = instalacao
    assert au.estado_atualizacao()["fase"] == "ocioso" and au.estado_atualizacao()["percentual"] is None
    assert au.download_and_install_update() is True
    e = au.estado_atualizacao()
    assert e["fase"] == "pronto" and e["baixado"] == e["total"] == len(conteudo) and e["percentual"] == 100.0 and e["erro"] == ""


def test_baixar_chama_o_progresso_em_ordem_crescente(instalacao, tmp_path):
    exe, conteudo, *_ = instalacao
    base = au._url_permitida()
    passos = []
    sha = au._baixar(base + "x.exe", str(tmp_path / "x.bin"), lambda b, t: passos.append((b, t)))
    assert sha == hashlib.sha256(conteudo).hexdigest() and passos[0] == (0, len(conteudo)) and passos[-1] == (len(conteudo), len(conteudo))
    assert [b for b, _ in passos] == sorted(b for b, _ in passos) and len(passos) >= 3


def test_cancelar_no_meio_do_download_apaga_o_parcial_e_nao_toca_no_exe(instalacao):
    exe, conteudo, ctrl, lancados, encerrou = instalacao
    ctrl["segura"].clear()                                                            # o servidor para na metade
    resultado = {}

    def roda():
        try:
            au.download_and_install_update()
        except Exception as e:  # noqa: BLE001
            resultado["erro"] = e
    t = threading.Thread(target=roda); t.start()
    assert ctrl["liberar"].wait(10) and espera(lambda: au.estado_atualizacao()["baixado"] > 0)
    e = au.estado_atualizacao()
    assert e["fase"] == "baixando" and 0 < e["percentual"] < 100
    assert au.cancelar_atualizacao() is True
    ctrl["segura"].set()                                                              # solta o resto: o próximo bloco já encontra o cancelamento
    t.join(15)
    assert isinstance(resultado.get("erro"), au.AtualizacaoCancelada)
    assert au.estado_atualizacao()["fase"] == "cancelado" and not exe.with_name(exe.name + ".novo").exists()
    assert exe.read_bytes() == b"MZ-versao-antiga" and lancados == [] and encerrou == []


def test_cancelar_sem_atualizacao_em_andamento_devolve_falso():
    assert au.cancelar_atualizacao() is False and not au._cancelar.is_set()


def test_sha_diferente_vira_erro_e_apaga_o_baixado(instalacao):
    exe, _, ctrl, lancados, encerrou = instalacao
    ctrl["sha"] = "0" * 64
    with pytest.raises(RuntimeError, match="SHA-256"):
        au.download_and_install_update()
    e = au.estado_atualizacao()
    assert e["fase"] == "erro" and "SHA-256" in e["erro"] and not exe.with_name(exe.name + ".novo").exists() and lancados == []


def test_tempo_maximo_total_do_download(instalacao, monkeypatch, tmp_path):
    monkeypatch.setattr(au, "MAX_DOWNLOAD_SEGUNDOS", -1)
    with pytest.raises(RuntimeError, match="tempo máximo"):
        au._baixar(au._url_permitida() + "x.exe", str(tmp_path / "x.bin"))


def test_iniciar_roda_em_segundo_plano_e_nao_deixa_iniciar_duas_vezes(instalacao):
    exe, conteudo, ctrl, lancados, encerrou = instalacao
    ctrl["segura"].clear()
    au.iniciar_atualizacao()                                                          # volta na hora, sem esperar o download
    assert au.estado_atualizacao()["fase"] in ("verificando", "baixando")
    with pytest.raises(RuntimeError, match="em andamento"):
        au.iniciar_atualizacao()
    assert ctrl["liberar"].wait(10)
    ctrl["segura"].set()
    assert espera(lambda: au.estado_atualizacao()["fase"] == "pronto")
    assert len(lancados) == 1 and encerrou == [True]
    with pytest.raises(RuntimeError, match="em andamento"):                          # depois de pronto o programa vai fechar: não reinicia
        au.iniciar_atualizacao()


def test_iniciar_fora_do_exe_e_recusado(monkeypatch):
    monkeypatch.setattr(au, "APP_MODE", "desktop")
    monkeypatch.setattr(au, "IS_FROZEN", False)
    with pytest.raises(RuntimeError, match="só no executável"):
        au.iniciar_atualizacao()
    assert au.estado_atualizacao()["fase"] == "ocioso"


# ───────── rotas ─────────
def cabecalho():
    from middleware.auth_middleware import create_jwt_token
    return {"Authorization": "Bearer " + create_jwt_token("u1", "a@x.com", "admin")}


def test_rotas_status_e_cancelar(monkeypatch):
    c = TestClient(app_mod.app)
    assert c.get("/api/update/status").status_code in (401, 403) and c.post("/api/update/cancel").status_code in (401, 403)
    s = c.get("/api/update/status", headers=cabecalho())
    assert s.status_code == 200 and s.json() == {"fase": "ocioso", "baixado": 0, "total": 0, "erro": "", "percentual": None}
    assert c.post("/api/update/cancel", headers=cabecalho()).json() == {"cancelado": False}
    au._estado.update(fase="baixando", baixado=50, total=200)
    assert c.get("/api/update/status", headers=cabecalho()).json()["percentual"] == 25.0
    assert c.post("/api/update/cancel", headers=cabecalho()).json() == {"cancelado": True} and au._cancelar.is_set()


def test_rota_apply_inicia_em_segundo_plano(monkeypatch):
    chamadas = []
    monkeypatch.setattr("routers.update.APP_MODE", "desktop")
    monkeypatch.setattr("routers.update.iniciar_atualizacao", lambda: chamadas.append(1))
    r = TestClient(app_mod.app).post("/api/update/apply", json={}, headers=cabecalho())
    assert r.status_code == 200 and r.json()["status"] == "iniciada" and chamadas == [1]


# ───────── tela ─────────
def ler(*partes):
    return open(os.path.join(*partes), encoding="utf-8").read()


def test_tela_tem_barra_de_progresso_e_botao_cancelar():
    html, js = ler(config.STATIC_DIR, "index.html"), ler(config.STATIC_DIR, "script.js")
    for trecho in ('id="update-progress"', 'id="update-bar"', 'id="update-progress-text"', 'id="btn-cancel-update"'):
        assert trecho in html, trecho
    for trecho in ("/api/update/status", "/api/update/cancel", "Baixando… ", "Cancelando…", "Conferindo a integridade", "Atualização cancelada", "setInterval"):
        assert trecho in js, trecho
