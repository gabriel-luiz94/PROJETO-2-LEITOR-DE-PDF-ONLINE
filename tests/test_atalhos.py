"""TASK-068 — API de atalhos/favoritos dos modais geradores (escopo Projeto)."""
import pytest
from fastapi.testclient import TestClient

import database
from middleware.auth_middleware import create_jwt_token


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    import app as appmod
    return TestClient(appmod.app)


def hdr(uid, role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@empresa.com', role)}"}


def item(id_, valor="CAA 2", grupo="ativos", **kw):
    return {"id": id_, "grupo": grupo, "valor": valor, **kw}


def fav(id_, **linha):
    return {"id": id_, "linha": {"op": "I", "ativo": "CAA 2", "fase": "ABC", "comp": 80, **linha}}


def put(client, h, dados, modal="cabos", projeto="229"):
    return client.put(f"/api/atalhos/{modal}", json={"projeto_codigo": projeto, "dados": dados}, headers=h)


def test_vazio_por_padrao(client):
    r = client.get("/api/atalhos/cabos?projeto_codigo=229", headers=hdr("ana"))
    assert r.status_code == 200 and r.json()["dados"] == {"itens": [], "favoritos": []}


def test_grava_le_e_isola_por_projeto_e_modal(client):
    r = put(client, hdr("ana"), {"itens": [item("a1", comp=40)], "favoritos": [fav("f1")]})
    assert r.status_code == 200
    d = client.get("/api/atalhos/cabos?projeto_codigo=229", headers=hdr("bia")).json()["dados"]
    assert d["itens"][0]["valor"] == "CAA 2" and d["itens"][0]["autor"] == "ana@empresa.com" and d["favoritos"][0]["id"] == "f1"
    assert client.get("/api/atalhos/cabos?projeto_codigo=027", headers=hdr("bia")).json()["dados"]["itens"] == []
    assert client.get("/api/atalhos/postes?projeto_codigo=229", headers=hdr("bia")).json()["dados"]["itens"] == []


def test_autor_vem_do_servidor_nao_do_cliente(client):
    put(client, hdr("ana"), {"itens": [item("a1", autor="outro@x.com")]})
    assert client.get("/api/atalhos/cabos?projeto_codigo=229", headers=hdr("ana")).json()["dados"]["itens"][0]["autor"] == "ana@empresa.com"


@pytest.mark.parametrize("dados", [
    {"itens": [item("a1", valor="<b>x</b>")]},
    {"itens": [item("a1", grupo="mt")]},                       # grupo de outro modal
    {"itens": [item("a 1")]},                                  # id inválido
    {"itens": [item("a1"), item("a1")]},                       # id repetido
    {"itens": [item("a1", comp=-3)]},
    {"itens": [item("x%d" % i) for i in range(201)]},
    {"favoritos": [fav("f1", op="Z")]},
    {"favoritos": [fav("f1", fase="XYZ")]},
    {"itens": "não é lista"},
])
def test_validacao_recusa(client, dados):
    assert put(client, hdr("ana"), dados).status_code == 400


def test_modal_desconhecido(client):
    assert client.get("/api/atalhos/xyz", headers=hdr("ana")).status_code == 404
    assert put(client, hdr("ana"), {}, modal="xyz").status_code == 400


def test_so_admin_exclui_e_so_o_dono_altera(client):
    put(client, hdr("ana"), {"itens": [item("a1")]})
    # bia acrescenta o dela mantendo o da ana: ok
    assert put(client, hdr("bia"), {"itens": [item("a1"), item("b1", valor="P 50")]}).status_code == 200
    # bia não exclui o da ana
    r = put(client, hdr("bia"), {"itens": [item("b1", valor="P 50")]})
    assert r.status_code == 403 and "administrador" in r.json()["detail"]
    # bia não altera o da ana
    r = put(client, hdr("bia"), {"itens": [item("a1", valor="MUDOU"), item("b1", valor="P 50")]})
    assert r.status_code == 403
    # bia altera o dela
    assert put(client, hdr("bia"), {"itens": [item("a1"), item("b1", valor="P 70")]}).status_code == 200
    # admin exclui e altera tudo
    assert put(client, hdr("adm", "admin"), {"itens": [item("b1", valor="P 70")]}).status_code == 200
    assert [x["id"] for x in client.get("/api/atalhos/cabos?projeto_codigo=229", headers=hdr("ana")).json()["dados"]["itens"]] == ["b1"]


def test_postes_usa_outros_grupos(client):
    d = {"itens": [item("p1", "DT11/300", "postes"), item("m1", "N1", "mt"), item("b1", "SI1", "bt")], "favoritos": [{"id": "f1", "linha": {"op": "I", "ativo": "DT11/300 1-N1"}}]}
    assert put(client, hdr("ana"), d, modal="postes").status_code == 200


def test_falha_da_nuvem_salva_local_e_avisa(client, monkeypatch):
    class Ruim:
        def table(self, *_a):
            raise RuntimeError("nuvem fora")
    from routers import atalhos
    monkeypatch.setattr(atalhos, "get_supabase", lambda: Ruim())
    r = put(client, hdr("ana"), {"itens": [item("a1")]})
    assert r.status_code == 500 and "localmente" in r.json()["detail"]
    monkeypatch.setattr(atalhos, "get_supabase", lambda: None)
    assert client.get("/api/atalhos/cabos?projeto_codigo=229", headers=hdr("ana")).json()["dados"]["itens"][0]["id"] == "a1"
