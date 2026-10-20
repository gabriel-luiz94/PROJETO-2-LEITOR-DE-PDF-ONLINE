"""TASK-074/075 — admin derivado do role na nuvem; leitura de regras do leitor ignora linha vazia da nuvem e a repovoa."""
import json
import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.auth as auth
import routers.regras_leitor as rl


class _Q:
    def __init__(self, nuvem, tabela):
        self.nuvem, self.tabela, self.op = nuvem, tabela, None
    def select(self, *_): self.op = "sel"; return self
    def eq(self, _c, v): self.v = v; return self
    def upsert(self, d): self.nuvem.setdefault(self.tabela, {})[d["projeto_codigo"]] = d["regras_json"]; self.op = "up"; return self
    def execute(self):
        class R: pass
        r = R(); r.data = []
        if self.op == "sel" and self.v in self.nuvem.get(self.tabela, {}):
            r.data = [{"regras_json": self.nuvem[self.tabela][self.v]}]
        return r


class _Sb:
    def __init__(self, nuvem): self.nuvem = nuvem
    def table(self, t): return _Q(self.nuvem, t)


@pytest.fixture
def banco(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"), raising=False)
    import config
    monkeypatch.setattr(config, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    return tmp_path


def _regra(n):
    return [{"ordem": n, "fase": 1, "modo": "DEFINIR", "texto_regex": "x"}]


def _local(proj, regras, tabela="processamento"):
    conn = database.get_connection()
    conn.execute(f"INSERT OR REPLACE INTO regras_leitor_{tabela} (projeto_codigo, regras_json) VALUES (?, ?)", (proj, json.dumps(regras)))
    conn.commit(); conn.close()


def test_nuvem_vazia_cai_no_sqlite_e_sobe_para_nuvem(banco, monkeypatch):
    nuvem = {"regras_leitor_processamento": {"P1": "[]"}}
    monkeypatch.setattr(rl, "get_supabase", lambda: _Sb(nuvem))
    _local("P1", _regra(1))
    assert rl._get_regras("processamento", "P1") == _regra(1)
    assert json.loads(nuvem["regras_leitor_processamento"]["P1"]) == _regra(1)


def test_nuvem_com_regras_tem_prioridade(banco, monkeypatch):
    nuvem = {"regras_leitor_processamento": {"P1": json.dumps(_regra(9))}}
    monkeypatch.setattr(rl, "get_supabase", lambda: _Sb(nuvem))
    _local("P1", _regra(1))
    assert rl._get_regras("processamento", "P1") == _regra(9)


def test_falha_na_leitura_da_nuvem_nao_envia_nada(banco, monkeypatch):
    class Quebrado:
        def table(self, _): raise RuntimeError("fora")
    monkeypatch.setattr(rl, "get_supabase", lambda: Quebrado())
    _local("P1", _regra(1))
    assert rl._get_regras("processamento", "P1") == _regra(1)


def test_sem_supabase_usa_sqlite(banco, monkeypatch):
    monkeypatch.setattr(rl, "get_supabase", lambda: None)
    _local("P1", _regra(1))
    assert rl._get_regras("processamento", "P1") == _regra(1)
    assert rl._get_regras("processamento", "NAO") == []


def _login_nuvem(monkeypatch, linha):
    class Q:
        def select(self, *_): return self
        def eq(self, *_): return self
        def execute(self):
            class R: data = [linha]
            return R()
    class Sb:
        def table(self, _): return Q()
    monkeypatch.setattr(auth, "get_supabase", lambda: Sb())
    monkeypatch.setattr(auth, "verify_password", lambda *_: True)
    c = TestClient(appmod.app)
    return c.post("/api/auth/login", json={"email": "a@b.c", "password": "x"})


def test_login_nuvem_role_admin_sem_coluna_is_admin(banco, monkeypatch):
    r = _login_nuvem(monkeypatch, {"id": 1, "email": "a@b.c", "senha_hash": "h", "role": "admin"})
    assert r.status_code == 200 and r.json()["is_admin"] is True and r.json()["role"] == "admin"


def test_login_nuvem_operador_nao_e_admin(banco, monkeypatch):
    r = _login_nuvem(monkeypatch, {"id": 2, "email": "a@b.c", "senha_hash": "h", "role": "operador"})
    assert r.status_code == 200 and r.json()["is_admin"] is False
