"""TASK-078 — coluna LV (LV / LM / A) da base de orçamento: linha viva (operação com `*`) x linha morta."""
import random
import sqlite3

import pytest
from fastapi.testclient import TestClient

import database
from middleware.auth_middleware import create_jwt_token
from services import orcamento_calc as calc
from services.orcamento_lv import lv_tolerante, normalizar_lv, operacao_viva, passa_lv
from tests import referencia_orcamento_calc_antes_lv as antes


# ── funções puras ──
def test_normalizar_lv():
    assert [normalizar_lv(x) for x in ["lv", " Lm ", "a", "", None, "  "]] == ["LV", "LM", "A", "A", "A", "A"]
    for ruim in ["XX", "viva", "L", "AMBOS", "1"]:
        with pytest.raises(ValueError):
            normalizar_lv(ruim)
    assert lv_tolerante("XX") == "A" and lv_tolerante("lv") == "LV"          # o cálculo nunca falha por dado ruim


def test_operacao_viva_e_matriz_passa_lv():
    assert [operacao_viva(o) for o in ["*I", "*R", "*M", " *I", "I", "R", "M", "", None]] == [True, True, True, True, False, False, False, False, False]
    for op in ["I", "R", "M"]:
        assert passa_lv("A", op) and passa_lv("A", "*" + op) and passa_lv(None, op) and passa_lv("", "*" + op)
        assert passa_lv("LM", op) and not passa_lv("LM", "*" + op)
        assert not passa_lv("LV", op) and passa_lv("LV", "*" + op)


# ── cálculo ──
def linha(ativo, componente, codigo, lv=None, mdo="MAO-DE-OBRA", desc_ativo="", fi=1.0, fr=1.0, projeto="", origem=""):
    r = {"ativo": ativo, "desc_ativo": desc_ativo or ativo, "componente": componente, "projeto": projeto, "mdo": mdo, "codigo": codigo,
         "desc_codigo": f"desc {codigo}", "fator_i": fi, "fator_r": fr, "filtro": "", "origem": origem}
    if lv is not None:
        r["lv"] = lv
    return r


def outros(texto, op):
    return {"ativo": texto, "operacao": op, "entidade": "0"}


def codigos(res):
    return sorted((r["codigo"], r["operacao"]) for r in res["resultado"])


# a linha do ativo (POSTE) vale para as duas operações; os demais códigos do componente variam
BASE = [linha("POSTE", "PST", "c_ambas", "A"), linha("", "PST", "c_viva", "LV"), linha("", "PST", "c_morta", "LM"), linha("", "PST", "c_sem")]


def test_expansao_do_componente_respeita_lv():
    morta = calc.processar_calculo([], [outros("2-POSTE", "I")], "", BASE)
    assert codigos(morta) == [("c_ambas", "I"), ("c_morta", "I"), ("c_sem", "I")]
    viva = calc.processar_calculo([], [outros("2-POSTE", "*I")], "", BASE)
    assert codigos(viva) == [("c_ambas", "I"), ("c_sem", "I"), ("c_viva", "I")]
    assert {r["codigo"]: r["total"] for r in viva["resultado"]}["c_viva"] == 2.0
    # remoção em linha viva usa os mesmos códigos (e a soma "R")
    rv = calc.processar_calculo([], [outros("1-POSTE", "*R")], "", BASE)
    assert codigos(rv) == [("c_ambas", "R"), ("c_sem", "R"), ("c_viva", "R")]


def test_ativo_so_com_o_outro_tipo_vai_para_nao_encontrados_com_dica():
    so_morta = [linha("CAA2", "CC", "k1", "LM")]
    r = calc.processar_calculo([], [outros("1-CAA2", "*I")], "", so_morta)
    assert r["resultado"] == [] and r["nao_encontrados"] == ["CAA2 (só existe para linha morta)"]
    assert calc.processar_calculo([], [outros("1-CAA2", "I")], "", so_morta)["nao_encontrados"] == []
    so_viva = [linha("CAA2", "CC", "k1", "LV")]
    r = calc.processar_calculo([], [outros("1-CAA2", "I")], "", so_viva)
    assert r["nao_encontrados"] == ["CAA2 (só existe para linha viva)"]
    assert calc.processar_calculo([], [outros("1-INEXISTENTE", "I")], "", so_viva)["nao_encontrados"] == ["INEXISTENTE"]       # sem dica


@pytest.mark.parametrize("passo, base, busca", [
    ("ativo", [linha("X1", "C1", "k1", "LV")], "X1"),
    ("componente", [linha("OUTRO", "X2", "k1", "LV")], "X2"),
    ("desc exata", [linha("OUTRO", "C3", "k1", "LV", desc_ativo="DESCRICAO EXATA")], "DESCRICAO EXATA"),
    ("desc parcial", [linha("OUTRO", "C4", "k1", "LV", desc_ativo="POSTE DUPLO T 9X150")], "DUPLO T"),
])
def test_cada_passo_da_cascata_filtra_lv(passo, base, busca):
    viva = calc.processar_calculo([], [outros(f"1-{busca}", "*I")], "", base)
    assert [r["codigo"] for r in viva["resultado"]] == ["k1"], passo
    morta = calc.processar_calculo([], [outros(f"1-{busca}", "I")], "", base)
    assert morta["resultado"] == [] and len(morta["nao_encontrados"]) == 1, passo


def test_codigo_direto_filtra_lv():
    base = [linha("", "", "9001", "LV")]
    assert [r["codigo"] for r in calc.processar_calculo([], [outros("3-9001", "*I")], "", base)["resultado"]] == ["9001"]
    r = calc.processar_calculo([], [outros("3-9001", "I")], "", base)
    assert r["resultado"] == [] and r["nao_encontrados"] == ["9001 (só existe para linha viva)"]


def test_cabos_tambem_respeitam_lv():
    base = [linha("CAA2", "CABO", "ka", "A"), linha("", "CABO", "kv", "LV"), linha("", "CABO", "km", "LM")]
    cab = lambda op: {"ativo": "CAA 2 ABC 80 m", "operacao": op, "entidade": "CABO"}
    assert codigos(calc.processar_calculo([cab("*I")], [], "", base)) == [("ka", "I"), ("kv", "I")]
    assert codigos(calc.processar_calculo([cab("I")], [], "", base)) == [("ka", "I"), ("km", "I")]


def test_regressao_base_toda_A_igual_ao_calculo_de_antes():
    """Base sem `lv`, com lv vazio ou com 'A' dá EXATAMENTE o resultado do cálculo anterior à TASK-078 (ops com e sem `*`)."""
    rnd = random.Random(78)
    ativos = [("POSTE", "PST"), ("CAA2", "CC"), ("DT9", "PST"), ("TR", "TRF")]
    base = []
    for i, (a, c) in enumerate(ativos):
        base.append(linha(a, c, f"cod{i}a", mdo="MATERIAL", fi=1.5, fr=0.5, projeto="RONDONIA" if i % 2 else ""))
        base.append(linha("", c, f"cod{i}b", mdo="MAO-DE-OBRA", fi=2.0, fr=1.0))
    base.append(linha("", "", "9001"))
    for _ in range(40):
        itens = [outros(f"{rnd.randint(1, 5)}-{rnd.choice(['POSTE', 'CAA2', 'DT9', 'TR', '9001', 'NADA'])}", rnd.choice(["I", "*I", "R", "*R", "M", "*M"]))
                 for _ in range(rnd.randint(1, 4))]
        esperado = antes.processar_calculo([], itens, rnd.choice(["", "RONDONIA"]), base)
        for variante in (base, [dict(r, lv="") for r in base], [dict(r, lv="A") for r in base], [dict(r, lv=None) for r in base]):
            proj = "RONDONIA"
            assert calc.processar_calculo([], itens, proj, variante) == antes.processar_calculo([], itens, proj, base)
        assert esperado is not None


# ── banco: migração e importações ──
@pytest.fixture
def banco(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    return tmp_path / "t.db"


def test_migracao_banco_antigo_sem_coluna_vira_A_e_e_idempotente(banco):
    con = sqlite3.connect(banco)
    for t in ("tabela_orcamento", "tabela_orcamento_master"):
        con.execute(f"CREATE TABLE {t} (id INTEGER PRIMARY KEY AUTOINCREMENT, ativo TEXT, desc_ativo TEXT, componente TEXT, projeto TEXT, mdo TEXT, codigo TEXT, desc_codigo TEXT, fator_i REAL, fator_r REAL, filtro TEXT, origem TEXT)")
        con.execute(f"INSERT INTO {t} (ativo, codigo) VALUES ('A1', 'c1'), ('A2', 'c2')")
    con.commit(); con.close()
    database.init_db()
    database.init_db()      # de novo: não muda nada
    con = sqlite3.connect(banco)
    for t in ("tabela_orcamento_master", "tabela_orcamento"):
        assert [r[0] for r in con.execute(f"SELECT lv FROM {t} WHERE ativo IN ('A1','A2')")] == ["A", "A"], t
    con.execute("UPDATE tabela_orcamento_master SET lv = '' WHERE ativo = 'A1'"); con.execute("UPDATE tabela_orcamento_master SET lv = NULL WHERE ativo = 'A2'"); con.commit(); con.close()
    database.init_db()
    con = sqlite3.connect(banco)
    assert [r[0] for r in con.execute("SELECT lv FROM tabela_orcamento_master WHERE ativo IN ('A1','A2')")] == ["A", "A"]


def test_seed_da_base_traz_lv_A_explicito(banco):
    database.init_db()
    con = sqlite3.connect(banco)
    total, nao_a = con.execute("SELECT COUNT(*), SUM(CASE WHEN lv <> 'A' THEN 1 ELSE 0 END) FROM tabela_orcamento").fetchone()
    assert total > 1000 and nao_a == 0
    with open("data/tabela_seed.csv", encoding="utf-8-sig") as f:
        assert f.readline().strip().endswith(";LV")


@pytest.fixture
def client(banco, monkeypatch):
    database.init_db()
    import routers.admin as adm, routers.orcamento as orc
    for m in (adm, orc):
        monkeypatch.setattr(m, "get_supabase", lambda: None, raising=False)
    import app as appmod
    c = TestClient(appmod.app)
    con = database.get_connection(); con.execute("DELETE FROM tabela_orcamento_master"); con.commit(); con.close()
    return c


def adm(): return {"Authorization": f"Bearer {create_jwt_token('u1', 'a@x.com', 'admin')}"}


CAB = "ATIVO;DESC ATIVO;COMPONENTE;PROJETO;MDO;CODIGO;DESC CODIGO;FATOR I;FATOR R;FILTRO"


def lvs(tabela="tabela_orcamento_master"):
    con = database.get_connection()
    try:
        return sorted(con.execute(f"SELECT codigo, lv FROM {tabela}").fetchall())
    finally:
        con.close()


def test_upload_master_com_e_sem_coluna_lv(client):
    csv_lv = CAB + ";LV\nA1;d;C;RONDONIA;MATERIAL;k1;d;1;1;;lv\nA2;d;C;RONDONIA;MATERIAL;k2;d;1;1;;Lm\nA3;d;C;RONDONIA;MATERIAL;k3;d;1;1;;\nA4;d;C;RONDONIA;MATERIAL;k4;d;1;1;;A\n"
    r = client.post("/api/admin/upload-master", files={"file": ("b.csv", csv_lv.encode("utf-8-sig"), "text/csv")}, headers=adm())
    assert r.status_code == 200, r.text
    assert [tuple(x) for x in lvs()] == [("k1", "LV"), ("k2", "LM"), ("k3", "A"), ("k4", "A")]
    antigo = CAB + "\nB1;d;C;RONDONIA;MATERIAL;z1;d;1;1;\n"                 # arquivo antigo, sem a coluna
    assert client.post("/api/admin/upload-master", files={"file": ("b.csv", antigo.encode("utf-8-sig"), "text/csv")}, headers=adm()).status_code == 200
    assert [tuple(x) for x in lvs()] == [("z1", "A")]


def test_upload_master_com_lv_invalido_recusa_e_nao_altera_nada(client):
    ok = CAB + ";LV\nA1;d;C;RONDONIA;MATERIAL;k1;d;1;1;;LV\n"
    client.post("/api/admin/upload-master", files={"file": ("b.csv", ok.encode("utf-8-sig"), "text/csv")}, headers=adm())
    ruim = CAB + ";LV\nA9;d;C;RONDONIA;MATERIAL;k9;d;1;1;;A\nA8;d;C;RONDONIA;MATERIAL;k8;d;1;1;;XX\n"
    r = client.post("/api/admin/upload-master", files={"file": ("b.csv", ruim.encode("utf-8-sig"), "text/csv")}, headers=adm())
    assert r.status_code == 400 and "Linha 3" in r.json()["detail"] and "XX" in r.json()["detail"]
    assert [tuple(x) for x in lvs()] == [("k1", "LV")]


def test_upload_local_lv_e_recusa_sem_apagar(client):
    ok = CAB + ";LV\nA1;d;C;P;MATERIAL;k1;d;1;1;;LM\n"
    assert client.post("/api/orcamento/upload", files={"file": ("b.csv", ok.encode("utf-8-sig"), "text/csv")}, headers=adm()).status_code == 200
    assert [tuple(x) for x in lvs("tabela_orcamento")] == [("k1", "LM")]
    ruim = CAB + ";LV\nA2;d;C;P;MATERIAL;k2;d;1;1;;ZZ\n"
    r = client.post("/api/orcamento/upload", files={"file": ("b.csv", ruim.encode("utf-8-sig"), "text/csv")}, headers=adm())
    assert r.status_code == 400 and "Nada foi alterado" in r.json()["error"]
    assert [tuple(x) for x in lvs("tabela_orcamento")] == [("k1", "LM")]


def test_add_row_salvar_e_sync_all(client):
    r = client.post("/api/admin/add-row", json={"ativo": "A1", "codigo": "k1"}, headers=adm())
    assert r.status_code == 200 and lvs() == [("k1", "A")]                                  # padrão A
    assert client.post("/api/admin/add-row", json={"ativo": "A2", "codigo": "k2", "lv": "lv"}, headers=adm()).status_code == 200
    assert client.post("/api/admin/add-row", json={"ativo": "A3", "codigo": "k3", "lv": "nope"}, headers=adm()).status_code == 400
    assert [tuple(x) for x in lvs()] == [("k1", "A"), ("k2", "LV")]
    dados = [{"ativo": "S1", "codigo": "s1", "lv": "LM"}, {"ativo": "S2", "codigo": "s2"}]
    assert client.post("/api/admin/sync-master-all", json={"dados": dados}, headers=adm()).status_code == 200
    assert [tuple(x) for x in lvs()] == [("s1", "LM"), ("s2", "A")]
    r = client.post("/api/admin/sync-master-all", json={"dados": [{"ativo": "Q", "codigo": "q", "lv": "??"}]}, headers=adm())
    assert r.status_code == 400 and [tuple(x) for x in lvs()] == [("s1", "LM"), ("s2", "A")]       # nada mudou
    assert client.post("/api/orcamento/salvar", json={"dados": [{"ativo": "U1", "codigo": "u1", "lv": "LV"}, {"ativo": "U2", "codigo": "u2", "lv": ""}]}, headers=adm()).status_code == 200
    assert [tuple(x) for x in lvs("tabela_orcamento") if x[0] in ("u1", "u2")] == [("u1", "LV"), ("u2", "A")]
    assert client.post("/api/orcamento/salvar", json={"dados": [{"ativo": "U3", "codigo": "u3", "lv": "X"}]}, headers=adm()).status_code == 400
    assert "u3" not in [x[0] for x in lvs("tabela_orcamento")]


def test_get_dados_devolve_lv(client):
    client.post("/api/admin/add-row", json={"ativo": "A2", "codigo": "k2", "lv": "LM"}, headers=adm())
    dados = client.get("/api/orcamento/dados", headers=adm()).json()["dados"]
    assert [d["lv"] for d in dados if d["codigo"] == "k2"] == ["LM"]


def test_sync_da_nuvem_sem_coluna_lv_vira_A(client, monkeypatch):
    import services.sync_service as ss

    class Tabela:
        def __init__(self, linhas): self.linhas = linhas
        def select(self, *_): return self
        def range(self, *_): return self
        def execute(self):
            class R: pass
            r = R(); r.data = self.linhas; return r

    class Sb:
        def table(self, _): return Tabela([{"ativo": "N1", "codigo": "n1"}, {"ativo": "N2", "codigo": "n2", "lv": "LV"}, {"ativo": "N3", "codigo": "n3", "lv": "xx"}])
    monkeypatch.setattr(ss, "get_supabase", lambda: Sb())
    monkeypatch.setattr(ss, "process_queue", lambda: None, raising=False)
    ss.sync_tabela_master()
    assert [tuple(x) for x in lvs()] == [("n1", "A"), ("n2", "LV"), ("n3", "A")]


def test_tela_da_base_tem_a_coluna_lv():
    html = open("static/orcamento.html", encoding="utf-8").read()
    assert 'data-col="11"' in html and "lvHTML(" in html and "'FILTRO', 'LV'" in html
    assert html.count("inputs.length < 12") >= 3 and "lv: inputs[11].value" in html
