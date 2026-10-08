"""TASK-057 — obras públicas/particulares: matriz de acesso (privacidade), filtros, permissões, IA, autônomo e nuvem (cliente falso)."""
import json
import types

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.ai_chat as chat
import routers.obras as rota
from middleware.auth_middleware import create_jwt_token
from services import obras_contexto as oc
from services.autonomo import pipeline

SNAP = {"cabos": {"bodyId": "body-cabos", "data": [{"entidade": "CABO", "operacao": "I", "ativo": "CAA 2 ABC 35 m"}]},
        "outros": {"bodyId": "body-outros", "data": [{"entidade": "ESTRUTURA", "operacao": "I", "ativo": "1-U4"}]}}


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(uid, role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@empresa.com', role)}"}


def salvar(client, uid, obra_id, nome, projeto="229", publica=None, role="operador"):
    corpo = {"id": obra_id, "nome": nome, "data": "01/10/2026, 10:00:00", "dados_json": json.dumps(SNAP), "projeto": projeto}
    if publica is not None:
        corpo["publica"] = publica
    r = client.post("/api/obras", json=corpo, headers=cab(uid, role))
    assert r.status_code == 200, r.text
    return r


def ids(client, uid, projeto="229", extra="", role="operador"):
    r = client.get(f"/api/obras?projeto={projeto}{extra}", headers=cab(uid, role))
    assert r.status_code == 200
    return sorted(o["id"] for o in r.json())


@pytest.fixture
def cenario(client):
    """A: particular PA e pública UA (229), pública de outro projeto UA27 (027). B: particular PB (229)."""
    salvar(client, "a", "PA", "Privada de A")
    salvar(client, "a", "UA", "Pública de A", publica=True)
    salvar(client, "a", "UA27", "Pública de A no 027", projeto="027", publica=True)
    salvar(client, "b", "PB", "Privada de B")
    return client


def test_padrao_e_particular(client):
    assert salvar(client, "a", "x", "X").json()["publica"] is False
    assert client.get("/api/obras/x", headers=cab("a")).json()["publica"] is False


def test_matriz_de_privacidade_em_todas_as_rotas(cenario):
    c = cenario
    # B (projeto 229): vê a sua e a pública de A; nunca a particular de A nem a pública de outro projeto
    assert ids(c, "b") == ["PB", "UA"]
    assert ids(c, "b", "027") == ["UA27"]
    assert ids(c, "b", extra="&origem=outros") == ["UA"]
    for rota_ in ("/api/obras/PA?projeto=229", "/api/obras/PA", "/api/obras/PA/exportar?projeto=229", "/api/obras/UA27?projeto=229"):
        assert c.get(rota_, headers=cab("b")).status_code == 404, rota_
    assert [o["id"] for o in c.get("/api/obras/indice?projeto=229", headers=cab("b")).json()] == ["PB", "UA"]
    # pública de outro SÓ com o projeto informado (sem ele, é como se não existisse)
    assert c.get("/api/obras/UA", headers=cab("b")).status_code == 404
    assert c.get("/api/obras/UA?projeto=229", headers=cab("b")).status_code == 200
    assert c.get("/api/obras/UA/exportar?projeto=229", headers=cab("b")).status_code == 200
    # A continua vendo as suas (inclusive de outro projeto na consulta própria)
    assert ids(cenario, "a") == ["PA", "UA"] and ids(cenario, "a", "027") == ["UA27"]
    # sem projeto só as próprias
    r = c.get("/api/obras", headers=cab("b")).json()
    assert [o["id"] for o in r] == ["PB"]


def test_resposta_nao_vaza_user_id_nem_email_so_nome_curto(cenario):
    o = cenario.get("/api/obras/UA?projeto=229", headers=cab("b")).json()
    assert o["dono"] == "a" and o["minha"] is False and o["publica"] is True
    texto = json.dumps(cenario.get("/api/obras?projeto=229", headers=cab("b")).json())
    assert "@" not in texto and "user_id" not in texto


def test_filtros_de_visibilidade_e_origem(cenario):
    c = cenario
    assert ids(c, "a", extra="&visibilidade=publicas") == ["UA"]
    assert ids(c, "a", extra="&visibilidade=particulares") == ["PA"]
    assert ids(c, "b", extra="&origem=minhas") == ["PB"]
    assert ids(c, "b", extra="&origem=outros&visibilidade=particulares") == []


def test_so_o_dono_altera_visibilidade_e_apaga(cenario):
    c = cenario
    put = lambda uid, oid, v: c.put(f"/api/obras/{oid}/visibilidade", json={"publica": v}, headers=cab(uid))
    assert put("b", "UA", False).status_code == 403            # enxerga (pública) mas não é dono
    assert put("b", "PA", True).status_code == 404             # não enxerga: não revela que existe
    assert put("b", "inexistente", True).status_code == 404
    assert ids(c, "b") == ["PB", "UA"]
    c.delete("/api/obras/UA", headers=cab("b"))                # tentativa de apagar a pública de A
    assert c.get("/api/obras/UA?projeto=229", headers=cab("a")).status_code == 200
    assert put("a", "PA", True).json()["publica"] is True and ids(c, "b") == ["PA", "PB", "UA"]
    assert put("a", "PA", False).status_code == 200 and ids(c, "b") == ["PB", "UA"]       # voltou a particular: some para B


def test_copia_de_b_continua_dele_quando_a_volta_a_particular(cenario):
    c = cenario
    sn = c.get("/api/obras/UA?projeto=229", headers=cab("b")).json()
    r = c.post("/api/obras", json={"id": "copiaB", "nome": sn["nome"], "data": "x", "dados_json": sn["dados_json"], "projeto": "229"}, headers=cab("b"))
    assert r.json()["publica"] is False
    c.put("/api/obras/UA/visibilidade", json={"publica": False}, headers=cab("a"))
    assert ids(c, "b") == ["PB", "copiaB"] and ids(c, "a", extra="&origem=minhas") == ["PA", "UA"]
    assert c.get("/api/obras/copiaB?projeto=229", headers=cab("a")).status_code == 404


def test_nao_sobrescreve_obra_de_outro_usuario_com_mesmo_id(cenario):
    c = cenario
    r = c.post("/api/obras", json={"id": "UA", "nome": "Sequestrada", "data": "x", "dados_json": "{}", "projeto": "229"}, headers=cab("b"))
    assert r.status_code == 403
    r = c.post("/api/obras", json={"id": "PA", "nome": "Sequestrada", "data": "x", "dados_json": "{}", "projeto": "229", "publica": True}, headers=cab("b"))
    assert r.status_code == 403
    assert c.get("/api/obras/UA?projeto=229", headers=cab("a")).json()["nome"] == "Pública de A"


def test_dono_sobrescreve_e_mantem_a_visibilidade_quando_nao_informada(cenario):
    c = cenario
    c.post("/api/obras", json={"id": "UA", "nome": "Nova versão", "data": "x", "dados_json": json.dumps(SNAP), "projeto": "229"}, headers=cab("a"))
    o = c.get("/api/obras/UA?projeto=229", headers=cab("b")).json()
    assert o["nome"] == "Nova versão" and o["publica"] is True


def test_obras_sem_dono_so_para_administradores_que_assumem(client):
    conn = database.get_connection()
    conn.execute("INSERT INTO obras (id, nome, data, dados_json, user_id, projeto) VALUES ('velha', 'Antiga', 'd', ?, NULL, '229')", (json.dumps(SNAP),))
    conn.commit(); conn.close()
    assert ids(client, "op") == [] and client.get("/api/obras/velha?projeto=229", headers=cab("op")).status_code == 404
    assert client.post("/api/obras/velha/assumir", headers=cab("op")).status_code == 403
    assert client.post("/api/obras", json={"id": "velha", "nome": "x", "data": "x", "dados_json": "{}", "projeto": "229"}, headers=cab("op")).status_code == 403
    assert ids(client, "adm", role="admin") == ["velha"]
    item = client.get("/api/obras?projeto=229", headers=cab("adm", "admin")).json()[0]
    assert item["sem_dono"] is True and item["minha"] is False
    assert client.post("/api/obras/velha/assumir", headers=cab("adm", "admin")).status_code == 200
    assert ids(client, "adm", role="admin", extra="&visibilidade=particulares&origem=minhas") == ["velha"]
    assert ids(client, "op") == [] and ids(client, "adm2", role="admin") == []          # agora é particular do adm
    assert client.post("/api/obras/velha/assumir", headers=cab("adm2", "admin")).status_code == 404


def test_importacao_continua_particular(cenario):
    c = cenario
    arq = c.get("/api/obras/UA/exportar?projeto=229", headers=cab("b")).json()
    r = c.post("/api/obras/importar?projeto=229", content=json.dumps(arq), headers={**cab("b"), "Content-Type": "application/json"})
    assert r.status_code == 200
    assert c.get(f"/api/obras/{r.json()['id']}?projeto=229", headers=cab("b")).json()["publica"] is False
    assert c.get(f"/api/obras/{r.json()['id']}?projeto=229", headers=cab("a")).status_code == 404


def test_autonomo_grava_sempre_particular(client):
    pipeline._salvar_obra("auto1", "Auto", "229", "a", SNAP)
    assert client.get("/api/obras/auto1", headers=cab("a")).json()["publica"] is False
    assert ids(client, "b") == []


# ── IA ──────────────────────────────────────────────────────────────────────
def req_falso(prompt, projeto="229", uid="b"):
    return types.SimpleNamespace(prompt=prompt, projeto_codigo=projeto), types.SimpleNamespace(state=types.SimpleNamespace(user={"user_id": uid, "email": f"{uid}@empresa.com", "role": "operador"}))


def test_ia_ve_proprias_e_publicas_de_outros_nunca_particulares(cenario):
    extra, instr = chat._contexto_obras(*req_falso("quais obras eu tenho?"))
    assert 'nome="Privada de B"' in extra and 'nome="Pública de A" | data=01/10/2026, 10:00:00 | pública de a' in extra
    assert "Privada de A" not in extra and "027" not in extra and "UA27" not in extra
    assert extra.index("Privada de B") < extra.index("Pública de A")                       # as próprias primeiro
    assert "pública de <nome>" in instr
    extra, _ = chat._contexto_obras(*req_falso("o que tem na obra Pública de A?"))
    assert 'CONTEÚDO DA OBRA "Pública de A"' in extra and "CAA 2 ABC 35 m" in extra
    extra, _ = chat._contexto_obras(*req_falso("o que tem na obra Privada de A?"))
    assert "CAA 2 ABC 35 m" not in extra and "CONTEÚDO" not in extra
    extra, _ = chat._contexto_obras(*req_falso("quais obras eu tenho?", projeto="027"))
    assert "Pública de A no 027" in extra and 'nome="Pública de A"' not in extra


def test_indice_marca_publicas_proprias_sem_mudar_as_demais():
    linhas = oc.montar_indice([{"id": "1", "nome": "N", "data": "d", "minha": True, "publica": False},
                               {"id": "2", "nome": "M", "data": "d", "minha": True, "publica": True},
                               {"id": "3", "nome": "O", "data": "d", "minha": False, "publica": True, "dono": "ana"}])
    assert '- id=1 | nome="N" | data=d\n' in linhas + "\n" and "| pública\n" in linhas and "pública de ana" in linhas
    assert oc.montar_indice([{"id": "1", "nome": "N", "data": "d"}]).endswith('data=d')       # formato antigo intacto


# ── nuvem (cliente FALSO) ─────────────────────────────────────────────────────────────────────────
class NuvemFalsa:
    """Mínimo do cliente Supabase usado pelas rotas. `colunas_novas=False` simula a tabela antes do SQL da TASK-057."""
    def __init__(self, colunas_novas=True):
        self.colunas_novas, self.linhas, self._f, self._op, self._pl, self._cols = colunas_novas, [], [], None, None, None

    def table(self, _):
        self._f, self._op, self._pl, self._cols = [], None, None, None
        return self

    def _checa(self, nomes):
        if not self.colunas_novas and {"publica", "dono_nome"} & set(nomes):
            raise RuntimeError("column does not exist")

    def upsert(self, p):
        self._checa(p)
        self.linhas = [x for x in self.linhas if x["id"] != p["id"]] + [dict(p)]
        return self

    def update(self, p):
        self._checa(p); self._op, self._pl = "update", p
        return self

    def select(self, cols="*"):
        if cols != "*":
            self._checa([c.strip() for c in cols.split(",")])
        self._op = "select"
        return self

    def eq(self, k, v):
        self._checa([k]); self._f.append((k, v)); return self

    def is_(self, k, v):
        self._f.append((k, None)); return self

    def order(self, *_, **__):
        return self

    def execute(self):
        alvo = [x for x in self.linhas if all((x.get(k) or None) == v if v is None else x.get(k) == v for k, v in self._f)]
        if self._op == "update":
            for x in alvo:
                x.update(self._pl)
        return types.SimpleNamespace(data=[dict(x) for x in alvo])


def test_nuvem_com_colunas_novas_aplica_a_mesma_regra(client, monkeypatch):
    nuvem = NuvemFalsa()
    monkeypatch.setattr(rota, "get_supabase", lambda: nuvem)
    salvar(client, "a", "PA", "Privada de A"); salvar(client, "a", "UA", "Pública de A", publica=True); salvar(client, "b", "PB", "Privada de B")
    assert {x["id"]: x["publica"] for x in nuvem.linhas} == {"PA": False, "UA": True, "PB": False}
    assert nuvem.linhas[0]["dono_nome"] == "a"
    assert ids(client, "b") == ["PB", "UA"]
    assert client.get("/api/obras/PA?projeto=229", headers=cab("b")).status_code == 404
    assert client.put("/api/obras/PA/visibilidade", json={"publica": True}, headers=cab("a")).status_code == 200
    assert next(x for x in nuvem.linhas if x["id"] == "PA")["publica"] is True and ids(client, "b") == ["PA", "PB", "UA"]
    nuvem.linhas.append({"id": "velha", "nome": "V", "data": "d", "dados_json": "{}", "user_id": None, "projeto": "229"})
    assert ids(client, "b") == ["PA", "PB", "UA"] and ids(client, "adm", role="admin") == ["PA", "UA", "velha"]
    assert client.post("/api/obras/velha/assumir", headers=cab("adm", "admin")).status_code in (200, 404)   # 404: só existe na nuvem (local não tem)


def test_nuvem_sem_as_colunas_novas_continua_funcionando_como_antes(client, monkeypatch):
    nuvem = NuvemFalsa(colunas_novas=False)
    monkeypatch.setattr(rota, "get_supabase", lambda: nuvem)
    salvar(client, "a", "PA", "Privada de A"); salvar(client, "b", "PB", "Privada de B")
    assert [x["id"] for x in nuvem.linhas] == ["PA", "PB"] and all("publica" not in x for x in nuvem.linhas)     # gravou sem as colunas novas
    assert ids(client, "a") == ["PA"] and ids(client, "b") == ["PB"]
    assert client.get("/api/obras/PA?projeto=229", headers=cab("b")).status_code == 404
    r = client.put("/api/obras/PA/visibilidade", json={"publica": True}, headers=cab("a"))
    assert r.status_code == 200 and "SQL" in r.json()["aviso"]


def test_nuvem_fora_do_ar_usa_o_local_com_a_mesma_regra(cenario, monkeypatch):
    class Quebrada:
        def table(self, *_): raise RuntimeError("fora do ar")
    monkeypatch.setattr(rota, "get_supabase", lambda: Quebrada())
    assert ids(cenario, "b") == ["PB", "UA"]
    assert cenario.get("/api/obras/PA?projeto=229", headers=cab("b")).status_code == 404


def test_regra_de_acesso_unitaria():
    pode = rota._pode_ver
    a, b, adm = {"user_id": "a"}, {"user_id": "b"}, {"user_id": "adm", "role": "admin"}
    assert pode({"user_id": "a", "projeto": "229"}, a, None)                                          # dono, sem projeto
    assert not pode({"user_id": "a", "projeto": "229", "publica": 0}, b, "229")                       # particular de outro
    assert pode({"user_id": "a", "projeto": "229", "publica": 1}, b, "229")                           # pública, mesmo projeto
    assert not pode({"user_id": "a", "projeto": "229", "publica": 1}, b, "027")                       # pública, outro projeto
    assert not pode({"user_id": "a", "projeto": "229", "publica": 1}, b, None)                        # pública, projeto não informado
    assert not pode({"user_id": None, "projeto": "229"}, b, "229") and pode({"user_id": None, "projeto": "229"}, adm, "229")
    assert not pode({"user_id": "", "projeto": "229", "publica": 1}, b, "229")                        # sem dono nunca é pública
    assert not pode({"user_id": None}, {"user_id": None}, None) and not pode({"user_id": "a"}, None, "229")   # anônimo
