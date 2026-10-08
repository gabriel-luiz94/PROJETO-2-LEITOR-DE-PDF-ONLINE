"""tests/test_orcamento_projeto.py — TASK-052: base de orçamento escopada por projeto.

Cobre: migração de linhas compartilhadas (projeto "A/B") em linhas exclusivas; filtro por
categoria (projeto exato, genéricas, não reconhecido) em GET /api/orcamento/dados; DELETE escopado
por projeto em upload-master/sync-master-all/salvar (nunca atinge outro projeto).
"""
import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from middleware.auth_middleware import create_jwt_token
from services.sync_service import (
    migrar_projetos_compartilhados, normalizar_projeto, validar_linhas_do_projeto,
    filtrar_linhas_por_categoria,
)


# ── funções puras ────────────────────────────────────────────────────────────

def test_normalizar_projeto():
    assert normalizar_projeto("  rondonia ") == "RONDONIA"
    assert normalizar_projeto(None) == ""
    assert normalizar_projeto("") == ""


def test_validar_linhas_do_projeto_preenche_vazio_e_rejeita_divergente():
    dados = [{"ativo": "A", "codigo": "1", "projeto": ""}, {"ativo": "B", "codigo": "2", "projeto": "paraiba"}]
    resolvidas = validar_linhas_do_projeto(dados, "PARAIBA")
    assert all(r["projeto"] == "PARAIBA" for r in resolvidas)

    with pytest.raises(ValueError):
        validar_linhas_do_projeto([{"ativo": "C", "codigo": "3", "projeto": "RONDONIA"}], "PARAIBA")


def test_filtrar_linhas_por_categoria():
    linhas = [
        {"ativo": "A", "projeto": ""},
        {"ativo": "B", "projeto": "RONDONIA"},
        {"ativo": "C", "projeto": "PARAIBA"},
        {"ativo": "D", "projeto": "XPTO"},
    ]
    assert [r["ativo"] for r in filtrar_linhas_por_categoria(linhas, projeto="RONDONIA")] == ["B"]
    assert [r["ativo"] for r in filtrar_linhas_por_categoria(linhas, genericas=True)] == ["A"]
    assert [r["ativo"] for r in filtrar_linhas_por_categoria(
        linhas, nao_reconhecido=True, projetos_validos=["RONDONIA", "PARAIBA"])] == ["D"]
    assert filtrar_linhas_por_categoria(linhas) == linhas  # sem filtro: devolve tudo


# ── migração (banco local) ───────────────────────────────────────────────────

@pytest.fixture
def banco(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    conn = database.get_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM tabela_orcamento_master")  # init_db não popula a master
    cur.execute('''INSERT INTO tabela_orcamento_master
        (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        ("DT15/3000", "POSTE X", "PST1", "PARAIBA/PARAIBANOVO", "MATERIAL", "647859", "POSTE X", 1.0, 1.0, "1.POSTE", ""))
    cur.execute('''INSERT INTO tabela_orcamento_master
        (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        ("TR315", "TRAFO", "ROTRF", "RONDONIA", "MATERIAL", "90050", "TRAFO", 1.0, 1.0, "", ""))
    cur.execute('''INSERT INTO tabela_orcamento_master
        (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        ("", "ALCA GENERICA", "C1", "", "MATERIAL", "90253", "ALCA", 1.0, 1.0, "91.ALCA", ""))
    conn.commit()
    conn.close()
    return tmp_path


def _linhas_master():
    conn = database.get_row_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM tabela_orcamento_master")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def test_migracao_divide_linha_compartilhada_sem_tocar_as_outras(banco):
    antes = _linhas_master()
    assert len(antes) == 3

    resultado = migrar_projetos_compartilhados()
    assert resultado["linhas_divididas_local"] == 1
    assert resultado["linhas_criadas_local"] == 1

    depois = _linhas_master()
    assert len(depois) == 4
    projetos = sorted(normalizar_projeto(r["projeto"]) for r in depois)
    assert projetos == ["", "PARAIBA", "PARAIBANOVO", "RONDONIA"]

    por_codigo = {r["codigo"]: r for r in depois}
    paraiba = next(r for r in depois if r["codigo"] == "647859" and r["projeto"] == "PARAIBA")
    paraibanovo = next(r for r in depois if r["codigo"] == "647859" and r["projeto"] == "PARAIBANOVO")
    for campo in ("ativo", "desc_ativo", "componente", "mdo", "desc_codigo", "fator_i", "fator_r", "filtro"):
        assert paraiba[campo] == paraibanovo[campo]
    # RONDONIA e a genérica não foram tocadas (mesmo id, mesmo projeto)
    rondonia = por_codigo["90050"]
    assert normalizar_projeto(rondonia["projeto"]) == "RONDONIA"
    generica = por_codigo["90253"]
    assert normalizar_projeto(generica["projeto"]) == ""


def test_migracao_e_idempotente(banco):
    migrar_projetos_compartilhados()
    depois_da_primeira = len(_linhas_master())
    resultado2 = migrar_projetos_compartilhados()
    assert resultado2["linhas_divididas_local"] == 0
    assert resultado2["linhas_criadas_local"] == 0
    assert len(_linhas_master()) == depois_da_primeira


def test_migracao_sem_linhas_compartilhadas_nao_faz_nada(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste2.db"))
    database.init_db()
    resultado = migrar_projetos_compartilhados()
    assert resultado == {"linhas_divididas_local": 0, "linhas_criadas_local": 0,
                          "linhas_divididas_nuvem": 0, "linhas_criadas_nuvem": 0}


# ── rotas ────────────────────────────────────────────────────────────────────

@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    conn = database.get_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM tabela_orcamento_master")
    cur.execute('''INSERT INTO tabela_orcamento_master
        (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        ("TR315", "TRAFO", "ROTRF", "RONDONIA", "MATERIAL", "90050", "TRAFO", 1.0, 1.0, "", ""))
    cur.execute('''INSERT INTO tabela_orcamento_master
        (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        ("DT15", "POSTE", "PST1", "PARAIBA", "MATERIAL", "647859", "POSTE", 1.0, 1.0, "", ""))
    cur.execute('''INSERT INTO tabela_orcamento_master
        (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        ("", "ALCA", "C1", "", "MATERIAL", "90253", "ALCA", 1.0, 1.0, "", ""))
    cur.execute('''INSERT INTO tabela_orcamento_master
        (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        ("XYZ", "ESTRANHO", "C2", "PROJETO_FANTASMA", "MATERIAL", "99999", "X", 1.0, 1.0, "", ""))
    cur.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES ('RONDONIA', '229')")
    cur.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES ('PARAIBA', '027')")
    conn.commit()
    conn.close()
    return TestClient(appmod.app)


def cab(role):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


def test_get_dados_sem_filtro_devolve_tudo(client):
    dados = client.get("/api/orcamento/dados").json()["dados"]
    assert len(dados) == 4


def test_get_dados_filtra_por_projeto_exato(client):
    dados = client.get("/api/orcamento/dados?projeto=RONDONIA").json()["dados"]
    assert len(dados) == 1 and dados[0]["codigo"] == "90050"

    dados_paraiba = client.get("/api/orcamento/dados?projeto=paraiba").json()["dados"]  # case-insensitive
    assert len(dados_paraiba) == 1 and dados_paraiba[0]["codigo"] == "647859"


def test_get_dados_genericas(client):
    dados = client.get("/api/orcamento/dados?genericas=true").json()["dados"]
    assert len(dados) == 1 and dados[0]["codigo"] == "90253"


def test_get_dados_nao_reconhecido(client):
    dados = client.get("/api/orcamento/dados?nao_reconhecido=true").json()["dados"]
    assert len(dados) == 1 and dados[0]["codigo"] == "99999"


def test_upload_master_escopado_nao_atinge_outro_projeto(client):
    csv_paraiba = (
        "ATIVO;DESC ATIVO;COMPONENTE;PROJETO;MDO;CODIGO;DESC CODIGO;FATOR I;FATOR R;FILTRO;ORIGEM\n"
        "DT99;POSTE NOVO;PST9;PARAIBA;MATERIAL;11111;POSTE NOVO;1;1;;\n"
    )
    r = client.post(
        "/api/admin/upload-master",
        files={"file": ("base.csv", csv_paraiba.encode("utf-8-sig"), "text/csv")},
        data={"projeto": "PARAIBA"},
        headers=cab("admin"),
    )
    assert r.status_code == 200, r.text

    dados = client.get("/api/orcamento/dados").json()["dados"]
    por_codigo = {d["codigo"]: d for d in dados}
    assert "11111" in por_codigo and por_codigo["11111"]["projeto"] == "PARAIBA"
    assert "647859" not in por_codigo  # a linha antiga de PARAIBA foi substituída
    assert "90050" in por_codigo       # RONDONIA intacta
    assert "90253" in por_codigo       # genérica intacta
    assert "99999" in por_codigo       # não reconhecida intacta


def test_upload_master_escopado_rejeita_linha_de_outro_projeto(client):
    csv_misto = (
        "ATIVO;DESC ATIVO;COMPONENTE;PROJETO;MDO;CODIGO;DESC CODIGO;FATOR I;FATOR R;FILTRO;ORIGEM\n"
        "DT99;POSTE NOVO;PST9;RONDONIA;MATERIAL;11111;POSTE NOVO;1;1;;\n"
    )
    r = client.post(
        "/api/admin/upload-master",
        files={"file": ("base.csv", csv_misto.encode("utf-8-sig"), "text/csv")},
        data={"projeto": "PARAIBA"},
        headers=cab("admin"),
    )
    assert r.status_code == 400
    # nada foi gravado: RONDONIA original continua intacta
    dados = client.get("/api/orcamento/dados?projeto=RONDONIA").json()["dados"]
    assert len(dados) == 1 and dados[0]["codigo"] == "90050"


def test_sync_master_all_escopado_nao_atinge_outro_projeto(client):
    r = client.post(
        "/api/admin/sync-master-all",
        json={"dados": [{"ativo": "DT77", "codigo": "22222", "projeto": "PARAIBA", "fator_i": 1, "fator_r": 1}],
              "projeto": "PARAIBA"},
        headers=cab("admin"),
    )
    assert r.status_code == 200, r.text
    dados = client.get("/api/orcamento/dados").json()["dados"]
    por_codigo = {d["codigo"]: d for d in dados}
    assert "22222" in por_codigo and "647859" not in por_codigo
    assert "90050" in por_codigo and "90253" in por_codigo and "99999" in por_codigo


def test_sync_master_all_sem_projeto_continua_substituindo_tudo(client):
    r = client.post(
        "/api/admin/sync-master-all",
        json={"dados": [{"ativo": "SO", "codigo": "1", "projeto": "RONDONIA", "fator_i": 1, "fator_r": 1}]},
        headers=cab("admin"),
    )
    assert r.status_code == 200, r.text
    dados = client.get("/api/orcamento/dados").json()["dados"]
    assert len(dados) == 1 and dados[0]["codigo"] == "1"  # tudo substituído, modo "todos os projetos"


def test_migrar_endpoint_requer_admin(client):
    assert client.post("/api/admin/migrar-projetos-compartilhados", headers=cab("operador")).status_code == 403


def test_migrar_endpoint_funciona(client):
    conn = database.get_connection()
    cur = conn.cursor()
    cur.execute('''INSERT INTO tabela_orcamento_master
        (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        ("SHARED", "X", "C3", "RONDONIA/PARAIBA", "MATERIAL", "55555", "X", 1.0, 1.0, "", ""))
    conn.commit()
    conn.close()

    r = client.post("/api/admin/migrar-projetos-compartilhados", headers=cab("admin"))
    assert r.status_code == 200
    assert r.json()["linhas_divididas_local"] == 1
    assert r.json()["linhas_criadas_local"] == 1

    dados = client.get("/api/orcamento/dados").json()["dados"]
    projetos_55555 = sorted(d["projeto"] for d in dados if d["codigo"] == "55555")
    assert projetos_55555 == ["PARAIBA", "RONDONIA"]


# ── Supabase (simulado) ──────────────────────────────────────────────────────
class FakeTabelaMaster:
    """Simula `tabela_orcamento_master` no Supabase: linhas em memória, `id` autoincrementado."""

    def __init__(self, linhas, proximo_id):
        self._linhas = linhas
        self._proximo_id = proximo_id
        self._filtros_eq = {}
        self._filtros_like = {}
        self._modo = None  # "select" | "insert" | "update" | "delete"
        self._payload = None

    def select(self, *_a, **_k):
        self._modo = "select"
        return self

    def insert(self, payload):
        self._modo = "insert"
        self._payload = payload
        return self

    def update(self, payload):
        self._modo = "update"
        self._payload = payload
        return self

    def delete(self):
        self._modo = "delete"
        return self

    def eq(self, campo, valor):
        self._filtros_eq[campo] = valor
        return self

    def neq(self, campo, valor):
        self._filtros_eq[(campo, "neq")] = valor
        return self

    def like(self, campo, padrao):
        self._filtros_like[campo] = padrao
        return self

    def ilike(self, campo, valor):
        self._filtros_like[campo] = ("ilike", valor)
        return self

    def _bate_filtros(self, linha):
        for campo, valor in self._filtros_eq.items():
            if isinstance(campo, tuple):
                continue
            if linha.get(campo) != valor:
                return False
        for campo, padrao in self._filtros_like.items():
            atual = (linha.get(campo) or "")
            if isinstance(padrao, tuple) and padrao[0] == "ilike":
                if atual.strip().upper() != padrao[1].strip().upper():
                    return False
            elif "%" in padrao:
                meio = padrao.strip("%")
                if meio not in atual:
                    return False
        return True

    def execute(self):
        class R:
            data = []
        r = R()
        if self._modo == "select":
            r.data = [l for l in self._linhas if self._bate_filtros(l)]
        elif self._modo == "insert":
            payloads = self._payload if isinstance(self._payload, list) else [self._payload]
            for p in payloads:
                novo = {**p, "id": self._proximo_id}
                self._proximo_id += 1
                self._linhas.append(novo)
        elif self._modo == "update":
            for l in self._linhas:
                if self._bate_filtros(l):
                    l.update(self._payload)
        elif self._modo == "delete":
            if (("id", "neq") in self._filtros_eq):
                self._linhas.clear()
            else:
                restantes = [l for l in self._linhas if not self._bate_filtros(l)]
                self._linhas[:] = restantes
        return r


class FakeSupabaseMaster:
    def __init__(self, linhas_iniciais):
        self._linhas = [dict(l) for l in linhas_iniciais]
        self._proximo_id = (max((l["id"] for l in self._linhas), default=0) + 1)

    def table(self, nome):
        assert nome == "tabela_orcamento_master"
        return FakeTabelaMaster(self._linhas, self._proximo_id)

    @property
    def linhas(self):
        return self._linhas


def test_migracao_tambem_divide_no_supabase_simulado(monkeypatch):
    database.init_db()  # cria as tabelas no banco temporário desta sessão de teste (sem linhas na master local)
    import services.sync_service as ss
    fake = FakeSupabaseMaster([
        {"id": 1, "ativo": "DT15", "desc_ativo": "P", "componente": "C", "projeto": "PARAIBA/PARAIBANOVO",
         "mdo": "MATERIAL", "codigo": "647859", "desc_codigo": "P", "fator_i": 1.0, "fator_r": 1.0,
         "filtro": "", "origem": ""},
        {"id": 2, "ativo": "TR", "desc_ativo": "T", "componente": "C2", "projeto": "RONDONIA",
         "mdo": "MATERIAL", "codigo": "90050", "desc_codigo": "T", "fator_i": 1.0, "fator_r": 1.0,
         "filtro": "", "origem": ""},
    ])
    monkeypatch.setattr(ss, "get_supabase", lambda: fake)

    resultado = ss.migrar_projetos_compartilhados()
    assert resultado["linhas_divididas_nuvem"] == 1
    assert resultado["linhas_criadas_nuvem"] == 1

    projetos = sorted(normalizar_projeto(l["projeto"]) for l in fake.linhas if l["codigo"] == "647859")
    assert projetos == ["PARAIBA", "PARAIBANOVO"]
    rondonia = next(l for l in fake.linhas if l["codigo"] == "90050")
    assert normalizar_projeto(rondonia["projeto"]) == "RONDONIA"  # intacta


def test_upload_master_escopado_no_supabase_simulado_nao_apaga_outro_projeto(client, monkeypatch):
    import routers.admin as admin_router
    fake = FakeSupabaseMaster([
        {"id": 1, "ativo": "TR315", "desc_ativo": "T", "componente": "C", "projeto": "RONDONIA",
         "mdo": "MATERIAL", "codigo": "90050", "desc_codigo": "T", "fator_i": 1.0, "fator_r": 1.0,
         "filtro": "", "origem": ""},
        {"id": 2, "ativo": "DT15", "desc_ativo": "P", "componente": "C2", "projeto": "PARAIBA",
         "mdo": "MATERIAL", "codigo": "647859", "desc_codigo": "P", "fator_i": 1.0, "fator_r": 1.0,
         "filtro": "", "origem": ""},
    ])
    monkeypatch.setattr(admin_router, "get_supabase", lambda: fake)

    csv_paraiba = (
        "ATIVO;DESC ATIVO;COMPONENTE;PROJETO;MDO;CODIGO;DESC CODIGO;FATOR I;FATOR R;FILTRO;ORIGEM\n"
        "DT99;POSTE NOVO;PST9;PARAIBA;MATERIAL;11111;POSTE NOVO;1;1;;\n"
    )
    r = client.post(
        "/api/admin/upload-master",
        files={"file": ("base.csv", csv_paraiba.encode("utf-8-sig"), "text/csv")},
        data={"projeto": "PARAIBA"},
        headers=cab("admin"),
    )
    assert r.status_code == 200, r.text

    codigos = {l["codigo"] for l in fake.linhas}
    assert "90050" in codigos    # RONDONIA intacta na nuvem
    assert "647859" not in codigos  # a linha antiga de PARAIBA foi substituída
    assert "11111" in codigos    # a nova linha de PARAIBA entrou
