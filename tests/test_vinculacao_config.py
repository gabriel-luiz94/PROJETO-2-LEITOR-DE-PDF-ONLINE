"""TASK-084 — API da configuração de vinculação: salva > Pernambuco > migração do legado, validação e arquivos servidos."""
import json
import os
import sqlite3

import pytest
from fastapi.testclient import TestClient

import config
import database
from middleware.auth_middleware import create_jwt_token


@pytest.fixture
def cli(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    import app as app_mod
    cli_ = TestClient(app_mod.app)
    cli_.headers.update({"Authorization": f"Bearer {create_jwt_token('u1', 'a@x.com', 'admin')}"})
    return cli_, tmp_path / "t.db"


def test_sem_config_migra_o_legado_e_pernambuco_nasce_com_ponto_de_partida(cli):
    c, _ = cli
    r = c.get("/api/vinculacao-config", params={"projeto_codigo": "229"}).json()
    assert r["origem"] == "migrada" and len(r["config"]["estruturas"]) >= 3 and r["config"]["cabos"] == []
    pe = c.get("/api/vinculacao-config", params={"projeto_codigo": "081"}).json()
    assert pe["origem"] == "pernambuco" and len(pe["config"]["cabos"]) >= 5 and any(e["padrao"] == "^N1$" for e in pe["config"]["estruturas"])
    assert c.get("/api/vinculacao-config", params={"projeto_codigo": "999"}).json()["origem"] == "vazia"


def test_salvar_e_ler_de_volta_sem_efeito_colateral_na_leitura(cli):
    c, db = cli
    c.get("/api/vinculacao-config", params={"projeto_codigo": "081"})
    assert sqlite3.connect(db).execute("SELECT COUNT(*) FROM vinculacao_config").fetchone()[0] == 0      # ler não grava
    cfg = c.get("/api/vinculacao-config", params={"projeto_codigo": "081"}).json()["config"]
    cfg["estruturas"][0]["qtd_cabos"] = 3
    assert c.post("/api/vinculacao-config", json={"projeto_codigo": "081", "config": cfg}).status_code == 200
    r = c.get("/api/vinculacao-config", params={"projeto_codigo": "081"}).json()
    assert r["origem"] == "salva" and r["config"]["estruturas"][0]["qtd_cabos"] == 3


def test_config_invalida_e_recusada(cli):
    c, _ = cli
    r = c.post("/api/vinculacao-config", json={"projeto_codigo": "081", "config": {"cabos": [{"padrao": "^("}], "estruturas": [{"padrao": "N1", "qtd_cabos": 0}]}})
    assert r.status_code == 400 and len(r.json()["detail"]["erros"]) == 2


def test_arquivos_de_vinculacao_servidos_sem_cache(cli):
    c, _ = cli
    for f in ("vinculacao_regras.js",):
        r = c.get(f"/static/{f}")
        assert r.status_code == 200 and "no-cache" in r.headers.get("cache-control", "").lower()
