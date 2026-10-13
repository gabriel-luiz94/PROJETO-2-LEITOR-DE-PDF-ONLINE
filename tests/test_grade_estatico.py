"""TASK-063/064/065/067 — estrutura estática do Resumo: sem coluna Entidade, módulo da grade servido, CSS do seletor."""
import os
import re

import pytest
from fastapi.testclient import TestClient

import config
import database


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    import app as appmod
    return TestClient(appmod.app)


def _ler(nome):
    with open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8") as f:
        return f.read()


def test_tabelas_cabos_e_outros_sem_coluna_ou_filtro_de_entidade():
    h = _ler("resumo.html")
    for t in ("cabos", "outros"):
        tabela = re.search(r'<table class="resumo-table" id="table-%s">.*?</thead>' % t, h, re.S).group(0)
        assert "Entidade" not in tabela and "entidade" not in tabela
        assert 'class="col-num"' in tabela
    assert "rfc-cabos-entidade" not in h and "rfc-outros-entidade" not in h


def test_modulo_da_grade_e_carregado_antes_do_resumo_e_tem_rota_sem_cache(client):
    h = _ler("resumo.html")
    assert h.index("/static/grade_tabela.js") < h.index("/static/resumo.js")
    r = client.get("/static/grade_tabela.js")
    assert r.status_code == 200 and "no-cache" in r.headers.get("cache-control", "").lower()


def test_lista_dos_selects_tem_fundo_cinza():
    css = _ler("styles.css")
    assert re.search(r"select option[^{]*\{[^}]*background-color:\s*#1c1f26", css)


def test_dado_de_cabos_continua_com_entidade_cabo_no_codigo():
    js = _ler("grade_tabela.js")
    assert "'CABO'" in js
