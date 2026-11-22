"""Login "Lembrar de mim": marcado, a sessão dura 30 dias e o programa abre direto; desmarcado, vale só enquanto o programa está aberto (a tela de login volta ao reabrir)."""
import os
import subprocess
import shutil

import jwt
import pytest
from fastapi.testclient import TestClient

import config
import database
from middleware import auth_middleware as am

NODE = shutil.which("node")


@pytest.fixture
def cliente(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    monkeypatch.setattr(config, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO usuarios_locais (email, senha_hash, is_admin, role) VALUES (?, ?, 0, 'operador')", ("lembrar@x.com", database.hash_password("senha-123")))
    conn.commit()
    conn.close()
    import app as appmod
    return TestClient(appmod.app)


def _duracao(token):
    p = jwt.decode(token, am.JWT_SECRET, algorithms=[am.JWT_ALGORITHM])
    return (p["exp"] - p["iat"]) / 3600


def test_lembrar_marcado_sessao_de_30_dias_e_desmarcado_24_horas(cliente):
    marcado = cliente.post("/api/auth/login", json={"email": "lembrar@x.com", "password": "senha-123", "lembrar": True})
    assert marcado.status_code == 200 and abs(_duracao(marcado.json()["access_token"]) - 24 * 30) < 0.01
    desmarcado = cliente.post("/api/auth/login", json={"email": "lembrar@x.com", "password": "senha-123", "lembrar": False})
    assert desmarcado.status_code == 200 and abs(_duracao(desmarcado.json()["access_token"]) - 24) < 0.01
    sem_campo = cliente.post("/api/auth/login", json={"email": "lembrar@x.com", "password": "senha-123"})
    assert abs(_duracao(sem_campo.json()["access_token"]) - 24) < 0.01                                  # clientes antigos: como sempre
    # o token longo vale de verdade nas rotas protegidas
    assert cliente.get("/api/auth/me", headers={"Authorization": "Bearer " + marcado.json()["access_token"]}).status_code == 200
    assert cliente.post("/api/auth/login", json={"email": "lembrar@x.com", "password": "errada", "lembrar": True}).status_code == 401


def test_token_padrao_continua_24_horas():
    t = am.create_jwt_token("1", "a@b.c", "operador")
    assert abs(_duracao(t) - 24) < 0.01 and abs(_duracao(am.create_jwt_token("1", "a@b.c", "operador", 5)) - 5) < 0.01


def _static(nome):
    return open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8").read()


def test_ligacoes_do_navegador():
    login, fetch_js, script = _static("login.js"), _static("auth_fetch.js"), _static("script.js")
    assert "localStorage.setItem('auth_lembrar', lembrar ? 'true' : 'false')" in login and "sessionStorage.setItem('sessao_ativa', '1')" in login
    assert "auth_lembrar') === 'false' && !sessionStorage.getItem('sessao_ativa')" in fetch_js and "'auth_lembrar'" in fetch_js
    assert "localStorage.removeItem('auth_lembrar')" in script


def test_segredo_jwt_do_desktop_e_persistente_e_isolado_por_perfil(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "APP_MODE", "desktop")
    monkeypatch.delenv("HEADLESS", raising=False)
    monkeypatch.setenv("JWT_SECRET", "")
    monkeypatch.setenv("APPDATA", str(tmp_path / "perfil1"))
    segredo = config._segredo_jwt()
    assert len(segredo) == 64 and config._segredo_jwt() == segredo
    monkeypatch.setenv("APPDATA", str(tmp_path / "perfil2"))
    assert config._segredo_jwt() != segredo


@pytest.mark.parametrize("modo,headless", [("server", "0"), ("desktop", "1")])
@pytest.mark.parametrize("segredo", ["", "dev-secret-change-in-production", "gere-uma-chave-secreta-forte-aqui"])
def test_jwt_servidor_recusa_segredo_padrao(modo, headless, segredo, monkeypatch):
    monkeypatch.setattr(config, "APP_MODE", modo)
    monkeypatch.setenv("HEADLESS", headless)
    monkeypatch.setenv("JWT_SECRET", segredo)
    with pytest.raises(RuntimeError, match="Configure JWT_SECRET"):
        config._segredo_jwt()
