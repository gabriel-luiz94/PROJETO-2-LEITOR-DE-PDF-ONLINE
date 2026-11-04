"""TASK-107 — coluna ARRED (NORMAL / INTEIRO / CIMA / BAIXO) da base de orçamento: arredondamento da quantidade de cada uso da linha."""
import sqlite3

import pytest
from fastapi.testclient import TestClient

import database
from middleware.auth_middleware import create_jwt_token
from services import orcamento_calc as calc
from services.orcamento_arred import aplicar_arred, arred_tolerante, normalizar_arred


def test_normalizar_arred():
    assert [normalizar_arred(x) for x in ["cima", " Baixo ", "inteiro", "normal", "", None, "  ", "para cima", "PARA BAIXO"]] == \
        ["CIMA", "BAIXO", "INTEIRO", "NORMAL", "NORMAL", "NORMAL", "NORMAL", "CIMA", "BAIXO"]
    for ruim in ["XX", "ceil", "2", "ARREDONDA"]:
        with pytest.raises(ValueError):
            normalizar_arred(ruim)
    assert arred_tolerante("XX") == "NORMAL" and arred_tolerante("cima") == "CIMA"       # o cálculo nunca falha por dado ruim


def test_aplicar_arred():
    assert [aplicar_arred(3.75, m) for m in ("NORMAL", "INTEIRO", "CIMA", "BAIXO")] == [3.75, 4.0, 4.0, 3.0]
    assert [aplicar_arred(3.5, "INTEIRO"), aplicar_arred(2.5, "INTEIRO"), aplicar_arred(3.49, "INTEIRO")] == [4.0, 3.0, 3.0]      # meio para cima
    assert aplicar_arred(3.0, "CIMA") == 3.0 and aplicar_arred(3.0, "BAIXO") == 3.0
    assert aplicar_arred(30 * 0.1, "CIMA") == 3.0 and aplicar_arred(29.999999999, "CIMA") == 30.0
    assert aplicar_arred(None or 0.0, "CIMA") == 0.0 and aplicar_arred(-3.2, "CIMA") == -3.0 and aplicar_arred(-3.2, "BAIXO") == -4.0
    assert aplicar_arred(3.75, "xx") == 3.75 and aplicar_arred(3.75, None) == 3.75


def L(ativo, cod, fi=1.0, arred=None, **k):
    r = {"ativo": ativo, "desc_ativo": "d", "componente": f"C{cod}", "projeto": "", "mdo": "MATERIAL", "codigo": cod, "desc_codigo": f"d {cod}",
         "fator_i": fi, "fator_r": fi, "filtro": "", "origem": "OUTROS", "lv": "A"}
    if arred is not None:
        r["arred"] = arred
    r.update(k)
    return r


def total(base, linhas, op="I"):
    r = calc.processar_calculo([], [{"ativo": t, "operacao": op, "entidade": "0"} for t in linhas], "", base)
    return {x["codigo"]: x["total"] for x in r["resultado"]}


def test_espacador_um_a_cada_8_metros_por_uso():
    base = [L("ESP", "3426163", fi=0.125, arred="CIMA")]
    assert total(base, ["30-ESP"]) == {"3426163": 4.0}
    assert total(base, ["64-ESP"]) == {"3426163": 8.0}
    assert total(base, ["5-ESP"]) == {"3426163": 1.0}
    assert total(base, ["30-ESP", "30-ESP", "5-ESP"]) == {"3426163": 9.0}          # cada uso arredonda sozinho (4 + 4 + 1)
    assert total([L("ESP", "3426163", fi=0.125)], ["30-ESP", "30-ESP", "6-ESP"]) == {"3426163": 8.25}        # NORMAL: soma o exato (3,75+3,75+0,75), 2 casas


def test_modos_inteiro_e_baixo_e_remocao_usa_fator_r():
    assert total([L("X", "k", fi=0.125, arred="BAIXO")], ["30-X"]) == {"k": 3.0}
    assert total([L("X", "k", fi=0.125, arred="INTEIRO")], ["30-X"]) == {"k": 4.0}
    assert total([L("X", "k", fi=0.125, arred="INTEIRO")], ["28-X"]) == {"k": 4.0}      # 3,5 → 4
    r = calc.processar_calculo([], [{"ativo": "30-X", "operacao": "R", "entidade": "0"}], "", [L("X", "k", fi=0.125, arred="CIMA", fator_r=0.25)])
    assert [(x["operacao"], x["total"]) for x in r["resultado"]] == [("R", 8.0)]          # 30 * 0,25 = 7,5 → 8 (arredonda a remoção também)


def test_sem_coluna_ou_normal_nao_muda_nada():
    sem = [L("X", "k", fi=1.5), L("Y", "j", fi=0.333)]
    normal = [L("X", "k", fi=1.5, arred="NORMAL"), L("Y", "j", fi=0.333, arred="")]
    ruim = [L("X", "k", fi=1.5, arred="LIXO"), L("Y", "j", fi=0.333, arred=None)]
    for linhas in (["3-X 7-Y", "2-X"], ["1-Y"]):
        assert total(sem, linhas) == total(normal, linhas) == total(ruim, linhas)
    assert total(sem, ["3-X 7-Y"]) == {"k": 4.5, "j": 2.33}


def test_arred_na_expansao_do_componente_e_codigo_direto_neutro():
    base = [L("X", "k1", fi=0.125, arred="CIMA", componente="CX"), L("", "k2", fi=0.5, arred="CIMA", componente="CX")]
    assert total(base, ["30-X"]) == {"k1": 4.0, "k2": 15.0}                               # 30 * 0,5 = 15 (já inteiro); cada linha com o seu arredondamento
    assert total([L("", "9001", fi=0.125, arred="CIMA", componente="")], ["2.5-9001"]) == {"9001": 2.5}      # código direto: fatores e arred neutros (como antes)


# ── banco: migração e importações ──
@pytest.fixture
def banco(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    return tmp_path / "t.db"


def test_migracao_banco_antigo_vira_normal_e_e_idempotente(banco):
    con = sqlite3.connect(banco)
    for t in ("tabela_orcamento", "tabela_orcamento_master"):
        con.execute(f"CREATE TABLE {t} (id INTEGER PRIMARY KEY AUTOINCREMENT, ativo TEXT, desc_ativo TEXT, componente TEXT, projeto TEXT, mdo TEXT, codigo TEXT, desc_codigo TEXT, fator_i REAL, fator_r REAL, filtro TEXT, origem TEXT)")
        con.execute(f"INSERT INTO {t} (ativo, codigo) VALUES ('A1', 'c1'), ('A2', 'c2')")
    con.commit(); con.close()
    database.init_db()
    database.init_db()
    con = sqlite3.connect(banco)
    for t in ("tabela_orcamento_master", "tabela_orcamento"):
        assert [r[0] for r in con.execute(f"SELECT arred FROM {t} WHERE ativo IN ('A1','A2')")] == ["NORMAL", "NORMAL"], t
    con.execute("UPDATE tabela_orcamento_master SET arred = 'CIMA' WHERE ativo = 'A1'"); con.execute("UPDATE tabela_orcamento_master SET arred = NULL WHERE ativo = 'A2'"); con.commit(); con.close()
    database.init_db()
    con = sqlite3.connect(banco)
    assert [r[0] for r in con.execute("SELECT arred FROM tabela_orcamento_master WHERE ativo IN ('A1','A2') ORDER BY ativo")] == ["CIMA", "NORMAL"]      # o que já é CIMA não volta para NORMAL


def test_seed_da_base_entra_como_normal(banco):
    database.init_db()
    con = sqlite3.connect(banco)
    total_, nao_normal = con.execute("SELECT COUNT(*), SUM(CASE WHEN arred <> 'NORMAL' THEN 1 ELSE 0 END) FROM tabela_orcamento").fetchone()
    assert total_ > 1000 and nao_normal == 0


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


def arreds(tabela="tabela_orcamento_master"):
    con = database.get_connection()
    try:
        return sorted(tuple(x) for x in con.execute(f"SELECT codigo, arred FROM {tabela}").fetchall())
    finally:
        con.close()


def test_upload_master_com_e_sem_coluna_arred(client):
    csv_ = CAB + ";ARRED\nA1;d;C;RONDONIA;MATERIAL;k1;d;1;1;;cima\nA2;d;C;RONDONIA;MATERIAL;k2;d;1;1;;Baixo\nA3;d;C;RONDONIA;MATERIAL;k3;d;1;1;;\nA4;d;C;RONDONIA;MATERIAL;k4;d;1;1;;inteiro\nA5;d;C;RONDONIA;MATERIAL;k5;d;1;1;;PARA CIMA\n"
    r = client.post("/api/admin/upload-master", files={"file": ("b.csv", csv_.encode("utf-8-sig"), "text/csv")}, headers=adm())
    assert r.status_code == 200, r.text
    assert arreds() == [("k1", "CIMA"), ("k2", "BAIXO"), ("k3", "NORMAL"), ("k4", "INTEIRO"), ("k5", "CIMA")]
    antigo = CAB + "\nB1;d;C;RONDONIA;MATERIAL;z1;d;1;1;\n"                              # arquivo antigo, sem a coluna
    assert client.post("/api/admin/upload-master", files={"file": ("b.csv", antigo.encode("utf-8-sig"), "text/csv")}, headers=adm()).status_code == 200
    assert arreds() == [("z1", "NORMAL")]


def test_upload_master_com_arred_invalido_recusa_e_nao_altera_nada(client):
    ok = CAB + ";ARRED\nA1;d;C;RONDONIA;MATERIAL;k1;d;1;1;;CIMA\n"
    client.post("/api/admin/upload-master", files={"file": ("b.csv", ok.encode("utf-8-sig"), "text/csv")}, headers=adm())
    ruim = CAB + ";ARRED\nA9;d;C;RONDONIA;MATERIAL;k9;d;1;1;;BAIXO\nA8;d;C;RONDONIA;MATERIAL;k8;d;1;1;;XX\n"
    r = client.post("/api/admin/upload-master", files={"file": ("b.csv", ruim.encode("utf-8-sig"), "text/csv")}, headers=adm())
    assert r.status_code == 400 and "Linha 3" in r.json()["detail"] and "XX" in r.json()["detail"]
    assert arreds() == [("k1", "CIMA")]


def test_upload_local_arred_e_recusa_sem_apagar(client):
    ok = CAB + ";ARRED\nA1;d;C;P;MATERIAL;k1;d;1;1;;BAIXO\n"
    assert client.post("/api/orcamento/upload", files={"file": ("b.csv", ok.encode("utf-8-sig"), "text/csv")}, headers=adm()).status_code == 200
    assert arreds("tabela_orcamento") == [("k1", "BAIXO")]
    ruim = CAB + ";ARRED\nA2;d;C;P;MATERIAL;k2;d;1;1;;ZZ\n"
    r = client.post("/api/orcamento/upload", files={"file": ("b.csv", ruim.encode("utf-8-sig"), "text/csv")}, headers=adm())
    assert r.status_code == 400 and "Nada foi alterado" in r.json()["error"]
    assert arreds("tabela_orcamento") == [("k1", "BAIXO")]


def test_add_row_e_sync_all(client):
    assert client.post("/api/admin/add-row", json={"ativo": "A1", "codigo": "k1"}, headers=adm()).status_code == 200
    assert client.post("/api/admin/add-row", json={"ativo": "A2", "codigo": "k2", "arred": "cima"}, headers=adm()).status_code == 200
    assert client.post("/api/admin/add-row", json={"ativo": "A3", "codigo": "k3", "arred": "nope"}, headers=adm()).status_code == 400
    assert arreds() == [("k1", "NORMAL"), ("k2", "CIMA")]
