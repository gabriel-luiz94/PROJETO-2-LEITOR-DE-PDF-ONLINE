"""TASK-131 — plugin do nanoCAD: lista de projetos da aplicação e configuração de LEITURA do desenho guardada por projeto."""
import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import services.sync_service as sync
import services.supabase_client as sb
from middleware.auth_middleware import create_jwt_token


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    monkeypatch.setattr(sb, "get_supabase", lambda: None)
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('PERNAMBUCO', '081')")
    conn.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('BAHIA', '071')")
    conn.commit()
    conn.close()
    monkeypatch.setattr(sync, "get_merged_orcamento", lambda user_id=None: [])
    return TestClient(appmod.app)


def cab():
    return {"Authorization": f"Bearer {create_jwt_token('u1', 'u1@x.com', 'operador')}"}


LEITURA = {"layers_rede_mt": ["REDE_MT", "ANOTACOES", "=0"], "layers_rede_bt": ["REDE_BT"], "blocos_poste": ["pg", "KMLP"],
           "cor_confirma": True, "raio_alturas": 4, "placa_regex": r"^[A-Z]{1,2}\d{5,7}$"}


def test_lista_todos_os_projetos_com_login(client):
    assert client.get("/api/plugin/projetos").status_code == 401
    r = client.get("/api/plugin/projetos", headers=cab()).json()
    cods = {p["codigo"]: p for p in r["projetos"]}
    assert {"081", "071"} <= set(cods)
    assert cods["071"]["tem_leitura"] is False


def test_leitura_guardada_vai_nas_regras_e_muda_a_versao(client):
    antes = client.get("/api/plugin/versao?projeto=081", headers=cab()).json()["versao"]["regras"]
    assert client.get("/api/plugin/regras?projeto=081", headers=cab()).json()["leitura"] is None
    r = client.put("/api/plugin/leitura?projeto=081", json={"leitura": dict(LEITURA, campo_estranho=1)}, headers=cab())
    assert r.status_code == 200 and r.json()["aviso"] is None
    assert "campo_estranho" not in r.json()["leitura"]
    regras = client.get("/api/plugin/regras?projeto=081", headers=cab()).json()
    assert regras["leitura"]["blocos_poste"] == ["pg", "KMLP"]
    assert client.get("/api/plugin/versao?projeto=081", headers=cab()).json()["versao"]["regras"] != antes   # o plugin percebe e baixa de novo
    assert client.get("/api/plugin/leitura?projeto=081", headers=cab()).json()["leitura"]["layers_rede_mt"] == ["REDE_MT", "ANOTACOES", "=0"]
    proj = {p["codigo"]: p for p in client.get("/api/plugin/projetos", headers=cab()).json()["projetos"]}
    assert proj["081"]["tem_leitura"] is True and proj["071"]["tem_leitura"] is False


@pytest.mark.parametrize("ruim", [{"layers_rede_mt": "REDE_MT"}, {"raio_alturas": -1}, {"raio_alturas": True}, {"placa_regex": "(["},
                                  {"cor_confirma": "sim"}, {"blocos_poste": [1, 2]}])
def test_configuracao_invalida_e_recusada(client, ruim):
    r = client.put("/api/plugin/leitura?projeto=081", json={"leitura": ruim}, headers=cab())
    assert r.status_code == 400 and r.json()["detail"]["erros"]


def test_projeto_inexistente_e_sem_login(client):
    assert client.put("/api/plugin/leitura?projeto=999", json={"leitura": LEITURA}, headers=cab()).status_code == 404
    assert client.put("/api/plugin/leitura?projeto=081", json={"leitura": LEITURA}).status_code == 401
