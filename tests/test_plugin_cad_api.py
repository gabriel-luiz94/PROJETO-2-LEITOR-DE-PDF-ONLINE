"""TASK-104 — API do plugin nanoCAD (routers/plugin_cad.py): versão (para o plugin perceber mudanças), regras do projeto numa chamada e
conferência da obra contra a Base de Orçamento (mesmo caminho do modo autônomo: Totalizadora → cálculo)."""
import json
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import services.sync_service as sync
from middleware.auth_middleware import create_jwt_token

RAIZ = Path(__file__).resolve().parent.parent
EXEMPLO = json.loads((RAIZ / "plugin_nanocad" / "exemplos" / "Projeto_Rua_das_Flores.obra.json").read_text(encoding="utf-8"))


def _linha(ativo, cod, proj="PERNAMBUCO", comp=""):
    return {"ativo": ativo, "desc_ativo": f"desc {ativo}", "componente": comp, "projeto": proj, "mdo": "MATERIAL", "codigo": cod,
            "desc_codigo": f"d {cod}", "fator_i": 1.0, "fator_r": 1.0, "filtro": "", "lv": "A"}


# base do teste: tudo do exemplo cadastrado para PERNAMBUCO, MENOS a estrutura N4 e o cabo S04
BASE = [_linha(a, str(100 + i)) for i, a in enumerate(
    ["N1", "N3", "B2", "DT9/200", "DT9/400", "DT11/300", "RC10/300", "CV9/300", "FV7/150", "A35P", "CAZ7,94", "C16", "RAMA25"])]


@pytest.fixture
def base():
    return list(BASE)


@pytest.fixture
def client(tmp_path, monkeypatch, base):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('PERNAMBUCO', '081')")
    conn.commit()
    conn.close()
    monkeypatch.setattr(sync, "get_merged_orcamento", lambda user_id=None: base)
    return TestClient(appmod.app)


def cab(uid="u1", role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@x.com', role)}"}


def test_exige_login_e_projeto_cadastrado(client):
    assert client.get("/api/plugin/versao?projeto=081").status_code == 401
    assert client.get("/api/plugin/versao?projeto=999", headers=cab()).status_code == 404


def test_regras_do_projeto_numa_chamada(client):
    r = client.get("/api/plugin/regras?projeto=081", headers=cab()).json()
    seed = json.loads((RAIZ / "data" / "regras_leitor_processamento_pernambuco_seed.json").read_text(encoding="utf-8"))
    assert r["projeto"] == {"codigo": "081", "nome": "PERNAMBUCO"}
    assert r["processamento"] == seed and len(r["classificacao"]) == 2
    assert {"cabos", "estruturas"} <= set(r["vinculacao"]) and r["vinculacao_origem"] == "pernambuco"
    assert set(r["versao"]) == {"regras", "conversao", "base", "geral"}


def test_versao_muda_quando_regras_ou_base_mudam(client, base):
    v0 = client.get("/api/plugin/versao?projeto=081", headers=cab()).json()["versao"]
    assert client.get("/api/plugin/versao?projeto=081", headers=cab()).json()["versao"] == v0         # estável
    cfg = client.get("/api/plugin/regras?projeto=081", headers=cab()).json()["vinculacao"]
    cfg["estruturas"][0]["qtd_cabos"] = 3
    assert client.post("/api/vinculacao-config", json={"projeto_codigo": "081", "config": cfg}, headers=cab(role="admin")).status_code == 200
    v1 = client.get("/api/plugin/versao?projeto=081", headers=cab()).json()["versao"]
    assert v1["regras"] != v0["regras"] and v1["base"] == v0["base"] and v1["geral"] != v0["geral"]
    base.append(_linha("N4", "999"))                                                                    # cadastrou N4 na base
    v2 = client.get("/api/plugin/versao?projeto=081", headers=cab()).json()["versao"]
    assert v2["base"] != v1["base"] and v2["regras"] == v1["regras"]


def _itens(resp, tabela, indice):
    return {x["ativo"]: x for l in resp["linhas"] if l["tabela"] == tabela and l["indice"] == indice for x in l["itens"]}


def test_conferir_aponta_o_que_nao_esta_na_base(client, base):
    r = client.post("/api/plugin/conferir", json={"projeto": "081", "dados": EXEMPLO["dados"]}, headers=cab())
    assert r.status_code == 200, r.text
    r = r.json()
    p2 = _itens(r, "outros", 1)                                  # P2: 9400D, N1, N4
    assert p2["DT9/400"]["naoEncontrado"] is False and p2["N1"]["naoEncontrado"] is False
    assert p2["N4"]["naoEncontrado"] is True                     # N4 não cadastrada para PERNAMBUCO
    v3 = _itens(r, "cabos", 2)                                   # V3: 3xS04-7,97kV 38 m
    assert v3["S04"]["naoEncontrado"] is True and v3["S04"]["foraDoCalculo"] is True
    assert _itens(r, "cabos", 0)["A35P"]["naoEncontrado"] is False
    assert "S04" in r["nao_encontrados"]
    # cadastrou N4 e S04 → nada pendente
    base += [_linha("N4", "998"), _linha("S04", "997")]
    r2 = client.post("/api/plugin/conferir", json={"projeto": "081", "dados": EXEMPLO["dados"]}, headers=cab()).json()
    assert r2["resumo"]["com_pendencia"] == 0 and r2["nao_encontrados"] == []


def test_conferir_ignora_campos_estranhos_e_limita_tamanho(client):
    dados = {"cabos": [{"entidade": "CABO", "operacao": "I", "ativo": "RAMA25 20 m", "lixo": {"x": 1}}], "outros": []}
    r = client.post("/api/plugin/conferir", json={"projeto": "081", "dados": dados}, headers=cab()).json()
    assert _itens(r, "cabos", 0)["RAMA25"]["naoEncontrado"] is False
    grande = {"cabos": [{"entidade": "CABO", "operacao": "I", "ativo": "RAMA25 1 m"}] * 5001, "outros": []}
    assert client.post("/api/plugin/conferir", json={"projeto": "081", "dados": grande}, headers=cab()).status_code == 413
