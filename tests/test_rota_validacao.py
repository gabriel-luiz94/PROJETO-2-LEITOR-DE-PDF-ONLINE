"""tests/test_rota_validacao.py — POST /api/validacao/planilhas."""
from fastapi.testclient import TestClient

import app as appmod
from middleware.auth_middleware import create_jwt_token

client = TestClient(appmod.app)


def test_exige_autenticacao():
    assert client.post("/api/validacao/planilhas", json={}).status_code == 401


def test_devolve_achados_e_resumo():
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
