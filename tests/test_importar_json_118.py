"""TASK-118: Importar JSON direto na tela (Carregar / Adicionar / Subtrair) — rota /api/obras/ler-arquivo e estrutura da tela."""
import json
import os

import pytest
from fastapi.testclient import TestClient

import config
import database
from middleware.auth_middleware import create_jwt_token

STATIC = config.STATIC_DIR


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    import app as appmod
    return TestClient(appmod.app)


def hdr(role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token('u1', 'u1@x.com', role)}"}


def _arquivo(client, projeto="027"):
    tabelas = {"cabos": {"bodyId": "body-cabos", "data": [{"entidade": "CABO", "operacao": "I", "ativo": "P 50"}, {"entidade": "CABO", "operacao": "I", "ativo": "CAZ9,5 abc 50 m"}]},
               "outros": {"bodyId": "body-outros", "data": [{"entidade": "0", "operacao": "I", "ativo": "1-N1"}]}}
    r = client.post(f"/api/obras/exportar-atual?projeto={projeto}", json={"nome": "Obra X", "dados_json": json.dumps(tabelas)}, headers=hdr())
    assert r.status_code == 200, r.text
    return r.json()


def _ler(client, corpo, projeto="027"):
    return client.post(f"/api/obras/ler-arquivo?projeto={projeto}", content=corpo if isinstance(corpo, (bytes, str)) else json.dumps(corpo),
                       headers={**hdr(), "Content-Type": "application/json"})


def test_ler_arquivo_devolve_as_linhas_sem_gravar(client):
    antes = client.get("/api/obras?projeto=027", headers=hdr()).json()
    r = _ler(client, _arquivo(client))
    assert r.status_code == 200, r.text
    d = r.json()
    assert (d["nome"], d["projeto"], d["tipo"], d["cabos"], d["outros"]) == ("Obra X", "027", "obra", 2, 1)
    assert [x["ativo"] for x in d["dados"]["cabos"]] == ["P 50", "CAZ9,5 abc 50 m"] and d["dados"]["outros"][0]["ativo"] == "1-N1"
    assert "totalizadora" in d["dados"]
    assert client.get("/api/obras?projeto=027", headers=hdr()).json() == antes                 # nada foi salvo


def test_ler_arquivo_recusa_outro_projeto_json_invalido_e_formato_errado(client):
    arq = _arquivo(client)
    r = _ler(client, arq, projeto="999")
    assert r.status_code == 400 and "projeto" in " ".join(r.json()["detail"]["erros"]).lower()
    assert _ler(client, "isto nao e json").status_code == 400
    assert _ler(client, {"formato": "outro"}).status_code == 400
    assert client.post("/api/obras/ler-arquivo?projeto=027", content=b"{}").status_code in (401, 403)          # exige login


def test_tela_importar_json_abre_o_dialogo_direto():
    html = open(os.path.join(STATIC, "resumo.html"), encoding="utf-8").read()
    assert 'id="input-importar-json"' in html
    js = open(os.path.join(STATIC, "resumo.js"), encoding="utf-8").read()
    i = js.index("TASK-118: \"Importar JSON\"")
    corpo = js[i:js.index("Payload de cálculo do orçamento")]
    assert "/api/obras/ler-arquivo" in corpo and "restoreObraSnapshot(d.dados)" in corpo
    assert "adicionarObraAoProjeto(d.dados)" in corpo and "subtrairObraDoProjeto(d.dados)" in corpo
    lay = open(os.path.join(STATIC, "resumo_layout.js"), encoding="utf-8").read()
    assert "$('input-importar-json')" in lay and "input-importar-obra" not in lay
