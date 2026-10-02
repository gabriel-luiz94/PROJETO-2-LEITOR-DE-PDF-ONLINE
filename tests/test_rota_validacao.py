"""tests/test_rota_validacao.py — POST /api/validacao/planilhas."""
import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from middleware.auth_middleware import create_jwt_token


@pytest.fixture
def client():
    database.init_db()          # o banco de cada teste é temporário (tests/conftest.py): cria as tabelas aqui
    return TestClient(appmod.app)


def test_exige_autenticacao(client):
    assert client.post("/api/validacao/planilhas", json={}).status_code == 401


def test_devolve_achados_e_resumo(client):
    token = create_jwt_token("u1", "a@b.com", "operador")
    resp = client.post(
        "/api/validacao/planilhas",
        headers={"Authorization": f"Bearer {token}"},
        json={"cabos": [{"ativo": "3-CFU", "operacao": "I"}], "outros": [{"ativo": "3-IP RECAL", "operacao": "I"}]},
    )
    assert resp.status_code == 200
    corpo = resp.json()
    assert sorted(a["regra_id"] for a in corpo["achados"]) == ["C1-CABO-PARTE", "C1-OUT-ORFAO"]
    assert corpo["resumo"] == {"erro": 2, "aviso": 0, "info": 0}
