"""TASK-144 — catálogo da Base de Orçamento para o autocompletar do plugin do nanoCAD (mesma lógica da dica de ativos da aplicação)."""
import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import services.sync_service as sync
from middleware.auth_middleware import create_jwt_token
from routers.plugin_cad import montar_catalogo


def L(ativo, desc="", projeto="PERNAMBUCO", origem="OUTROS", cod="1"):
    return {"ativo": ativo, "desc_ativo": desc, "componente": "", "projeto": projeto, "mdo": "MATERIAL", "codigo": cod, "desc_codigo": "",
            "fator_i": 1.0, "fator_r": 1.0, "filtro": "", "origem": origem, "lv": "A"}


BASE = [L("N1(tipocabo=P|M, poste)", "ESTRUTURA N1"), L("", "ESTRUTURA N1", cod="2"), L("DT11/300, 11300D", "POSTE DT 11/300"),
        L("A35P", "CABO A35P", origem="CABO"), L("U3", "ESTRAI U3", projeto="BAHIA")]


def test_catalogo_traz_nomes_descricoes_apelidos_e_parametros():
    c = montar_catalogo(BASE, "PERNAMBUCO")
    nomes = {a["nome"]: a for a in c["ativos"]}
    assert {"N1", "DT11/300", "A35P"} <= set(nomes) and "U3" not in nomes          # só o que vale no projeto
    assert nomes["N1"]["desc"] == "ESTRUTURA N1" and nomes["A35P"]["origem"] == "CABO"
    assert c["apelidos"].get("11300D") == "DT11/300"
    ps = c["parametros"]["N1"]
    assert [p["nome"].upper() for p in ps][0] == "TIPOCABO"
    assert set(ps[0]["valores"]) >= {"P", "M"}
    assert c["versao"] == montar_catalogo(BASE, "PERNAMBUCO")["versao"]


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('PERNAMBUCO', '081')")
    conn.commit()
    conn.close()
    monkeypatch.setattr(sync, "get_merged_orcamento", lambda user_id=None: BASE)
    return TestClient(appmod.app)


def test_rota_com_login(client):
    assert client.get("/api/plugin/catalogo?projeto=081").status_code == 401
    h = {"Authorization": f"Bearer {create_jwt_token('u1', 'u1@x.com', 'operador')}"}
    j = client.get("/api/plugin/catalogo?projeto=081", headers=h).json()
    assert j["projeto"]["codigo"] == "081" and any(a["nome"] == "N1" for a in j["ativos"])
