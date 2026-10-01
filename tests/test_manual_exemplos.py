"""TASK-030 — o manual existe, exige login e TODOS os seus exemplos JSON são válidos (nunca desatualiza)."""
import json
import re

import pytest
from fastapi.testclient import TestClient

import config
from services.ajustes_planilhas import ajustar, validar_acoes
from services.regras_dominio import _validar_cond, avaliar, validar_regras

with open(config.MANUAL_PATH, encoding="utf-8") as f:
    TEXTO = f.read()
with open(config.REGRAS_DOMINIO_SEED_PATH, encoding="utf-8") as f:
    GRUPOS = json.load(f)["grupos"]


def _blocos(tipo):
    return [json.loads(b) for b in re.findall(r"```json " + tipo + r"\n(.*?)```", TEXTO, re.S)]


def test_manual_tem_exemplos_de_cada_tipo():
    assert len(_blocos("regra")) >= 4 and len(_blocos("acoes")) >= 8 and len(_blocos("condicao")) >= 2


@pytest.mark.parametrize("regra", _blocos("regra"), ids=lambda r: r["id"])
def test_exemplo_de_regra_valido(regra):
    assert validar_regras([regra], GRUPOS) == []


@pytest.mark.parametrize("cond", _blocos("condicao"))
def test_exemplo_de_condicao_valido(cond):
    erros = []
    _validar_cond(cond, "c", erros, GRUPOS, [0])
    assert erros == []


@pytest.mark.parametrize("acoes", _blocos("acoes"))
def test_exemplo_de_acoes_valido_e_executa(acoes):
    assert validar_acoes(acoes, GRUPOS) == []
    r = ajustar(acoes, [{"operacao": "I", "ativo": "DT11/300 1-AA", "entidade": "0"}],
                [{"operacao": "I", "ativo": "1-CFU 1-U4", "entidade": "0"}], GRUPOS)
    assert "operacoes" in r


def test_exemplo_e_cfu_u4_avisa_so_quando_tem_os_dois():
    regra = next(r for r in _blocos("regra") if r["id"] == "EX-CFU-U4-SUPL")
    def alertas(texto):
        return avaliar([regra], [], [{"operacao": "I", "ativo": texto, "entidade": "0"}], GRUPOS)
    assert len(alertas("DT11/300 1-CFU 1-U4")) == 1
    assert alertas("DT11/300 1-CFU") == []          # só CFU: não avisa
    assert alertas("DT11/300 1-U4") == []           # só U4: não avisa
    assert alertas("DT11/300 1-CFU 1-U4 1-SUPL") == []


@pytest.fixture
def cliente(tmp_path, monkeypatch):
    import app as app_mod
    import database
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    return TestClient(app_mod.app)


def test_manual_exige_login(cliente):
    assert cliente.get("/api/manual").status_code == 401


@pytest.mark.parametrize("papel", ["operador", "admin"])
def test_manual_disponivel_para_qualquer_perfil(cliente, papel):
    from middleware.auth_middleware import create_jwt_token
    cab = {"Authorization": f"Bearer {create_jwt_token('u1', f'{papel}@x.com', papel)}"}
    r = cliente.get("/api/manual", headers=cab)
    assert r.status_code == 200 and "Como usar o \"E\"" in r.text
    assert "markdown" in r.headers["content-type"]


def test_pagina_do_manual_abre(cliente):
    r = cliente.get("/manual")
    assert r.status_code == 200 and "Manual" in r.text
