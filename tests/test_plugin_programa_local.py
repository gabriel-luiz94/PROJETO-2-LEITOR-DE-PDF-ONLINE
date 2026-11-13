"""TASK-143 — o plugin do nanoCAD entrega a obra ao GridReader instalado (aberto) ou o abre já com ela (fechado)."""
import json
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import app as appmod
import config
import database
from middleware.auth_middleware import create_jwt_token
from services import janela_desktop, programa_local

RAIZ = Path(__file__).resolve().parent.parent
ENV = json.loads((RAIZ / "plugin_nanocad" / "exemplos" / "Projeto_Rua_das_Flores.obra.json").read_text(encoding="utf-8"))


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('PERNAMBUCO', '081')")
    conn.commit()
    conn.close()
    monkeypatch.setattr(config, "APP_MODE", "desktop")
    return TestClient(appmod.app)


def cab():
    return {"Authorization": f"Bearer {create_jwt_token('u1', 'u1@x.com', 'operador')}"}


def test_programa_se_apresenta_sem_login(client, monkeypatch):
    j = client.get("/api/plugin/desktop").json()
    assert j["app"] == "GridReader" and j["desktop"] is True
    monkeypatch.setattr(config, "APP_MODE", "server")
    assert client.get("/api/plugin/desktop").json()["desktop"] is False


def test_plugin_antigo_so_com_a_obra_ainda_e_recebido_sem_gravar_no_banco(client, monkeypatch):
    urls = []
    monkeypatch.setattr(janela_desktop, "mostrar", lambda url, porta=None: urls.append(url) or True)
    r = client.post("/api/plugin/receber", json={"chave": "C:/obras/rua.dwg", "envelope": ENV})       # sem token: é o plugin
    assert r.status_code == 200 and r.json()["janela"] is True
    assert urls == [r.json()["url"]] and r.json()["url"].startswith("/resumo?plugin_pendente=")
    i = r.json()["id"]
    assert client.get(f"/api/plugin/pendente/{i}").status_code == 401                                 # a janela busca com login
    d = client.get(f"/api/plugin/pendente/{i}", headers=cab()).json()
    assert d["chave"] == "C:/obras/rua.dwg" and d["envelope"]["tipo"] == "obra"
    assert client.get("/api/obras?projeto=081", headers=cab()).json() == []                           # TASK-162: nada no banco


def test_so_no_programa_instalado_e_so_do_proprio_computador(client, monkeypatch):
    monkeypatch.setattr(janela_desktop, "mostrar", lambda url, porta=None: True)
    monkeypatch.setattr(config, "APP_MODE", "server")
    assert client.post("/api/plugin/receber", json={"chave": "a", "envelope": ENV}).status_code == 403
    monkeypatch.setattr(config, "APP_MODE", "desktop")
    remoto = TestClient(appmod.app, client=("192.168.0.9", 50000))
    assert remoto.post("/api/plugin/receber", json={"chave": "a", "envelope": ENV}).status_code == 403
    assert client.post("/api/plugin/receber", json={"chave": "a", "envelope": {"tipo": "modelo"}}).status_code == 400
    assert client.get("/api/plugin/pendente/naoexiste", headers=cab()).status_code == 404


def test_programa_fechado_abre_com_o_arquivo_do_plugin(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setenv("GRIDREADER_REGISTRAR", "1")
    programa_local.registrar(8123, "2.2.0")
    reg = json.loads((tmp_path / "OrcamentoNanoCAD" / "programa.json").read_text(encoding="utf-8"))
    assert reg["porta"] == 8123 and reg["versao"] == "2.2.0" and reg["exe"]
    arq = tmp_path / "obra.json"
    arq.write_text(json.dumps({"chave": "rua.dwg", "envelope": ENV}), encoding="utf-8")
    assert programa_local.ler_argumento(["--abrir-obra", str(arq)]) == str(arq)
    assert programa_local.ler_argumento([f"--abrir-obra={arq}"]) == str(arq)
    i = programa_local.carregar_arquivo(str(arq))
    assert i and programa_local.pegar(i)["chave"] == "rua.dwg" and not arq.exists()      # o arquivo é apagado
    ruim = tmp_path / "ruim.json"
    ruim.write_text("{}", encoding="utf-8")
    assert programa_local.carregar_arquivo(str(ruim)) == ""


def test_app_e_resumo_ligados():
    app_py = (RAIZ / "app.py").read_text(encoding="utf-8")
    assert "programa_local.registrar(port, APP_VERSION)" in app_py and "janela_desktop.vincular(window)" in app_py
    assert "programa_local.url_pendente(_id_obra)" in app_py
    js = (RAIZ / "static" / "resumo.js").read_text(encoding="utf-8")
    assert "plugin_pendente" in js and "/api/plugin/pendente/" in js
