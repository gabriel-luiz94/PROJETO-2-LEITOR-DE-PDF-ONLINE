"""TASK-109 — FATOR I/R da base em função de parâmetro (fórmula "=NUM(fios)/8")."""
import sqlite3

import pytest
from fastapi.testclient import TestClient

import database
from middleware.auth_middleware import create_jwt_token
from services import ativo_parametros as ap
from services import orcamento_calc as calc
from services.fator_formula import FormulaErro, avaliar, compilar, identificadores, normalizar_fator


# ── a linguagem ──
@pytest.mark.parametrize("formula, params, esperado", [
    ("=NUM(fios)", {"FIOS": "8F"}, 8.0),
    ("=NUM(fios)/8", {"FIOS": "12F"}, 1.5),
    ("NUM(fios)*2", {"fios": "10F"}, 20.0),                                  # sem o "=" e nome em minúsculas
    ("=1,5*2", {}, 3.0),
    ("=1.5*2", {}, 3.0),
    ("=-NUM(f)+10", {"F": "4"}, 6.0),
    ("=2*(3+4)", {}, 14.0),
    ("=fios*2", {"FIOS": "8"}, 16.0),                                        # texto numérico puro entra em conta
    ('=SE(tipo="A"; 2; 1)', {"TIPO": "a"}, 2.0),
    ('=SE(tipo="A"; 2; 1)', {"TIPO": "B"}, 1.0),
    ("=SE(NUM(fios)>=10; 2; 1)", {"FIOS": "12F"}, 2.0),
    ("=SE(E(NUM(a)>1; NUM(a)<5); 1; 0)", {"A": "3X"}, 1.0),
    ("=SE(OU(tipo=\"A\"; tipo=\"B\"); 7; 8)", {"TIPO": "B"}, 7.0),
    ("=MAX(NUM(cabo); 35) * 2", {"CABO": "3XA16"}, 70.0),
    ("=MIN(3; 2; 5)", {}, 2.0),
    ("=ABS(-4)", {}, 4.0),
    ("=ARRED(NUM(f)/8; 0)", {"F": "12F"}, 2.0),
    ("=CIMA(NUM(f)/8)", {"F": "9F"}, 2.0),
    ("=BAIXO(NUM(f)/8)", {"F": "15F"}, 1.0),
    ("=NUM(cabo#2)", {"CABO": "A1", "CABO#2": "B9"}, 9.0),
    ("=NUM(cabo)", {"CABO#1": "A7"}, 7.0),
    ("=NUM(\"S336,4\")", {}, 336.4),                                           # vírgula decimal entre dígitos
])
def test_avaliar(formula, params, esperado):
    assert avaliar(formula, params) == esperado


@pytest.mark.parametrize("formula, trecho", [
    ("=fios*2", "use NUM(fios)"),                                            # "8F" não é número
    ("=x", "não informado"),
    ("=1/0", "divisão por zero"),
    ("=NUM(f)", "não tem número"),
    ("=1+", "incompleta"),
    ("=foo(1)", "função desconhecida"),
    ("=SE(1)", "3 argumentos"),
    ("=2 $ 3", "inesperado"),
    ("=", "vazia"),
    ("=(1+2", "incompleta"),
    ("=1 2", "sobrou"),
])
def test_erros(formula, trecho):
    with pytest.raises(FormulaErro) as e:
        avaliar(formula, {"F": "abc", "FIOS": "8F"})
    assert trecho in str(e.value)


def test_nada_de_eval():
    for ruim in ["=__import__('os')", "=open('x')", "=1; 2", "=a.b.c"]:
        with pytest.raises(FormulaErro):
            avaliar(ruim, {})


def test_identificadores_e_normalizar_fator():
    assert identificadores("=NUM(fios)/8+cabo#2*SE(tipo=\"A\";1;2)") == {"FIOS", "CABO", "TIPO"}
    assert normalizar_fator("1,5") == (1.5, "") and normalizar_fator(2) == (2.0, "") and normalizar_fator(None) == (0.0, "") and normalizar_fator("abc") == (0.0, "")
    assert normalizar_fator(" =NUM(a) / 8 ") == (0.0, "=NUM(a) / 8")
    with pytest.raises(ValueError):
        normalizar_fator("=1+")
    compilar("=1+1")


# ── cálculo ──
def L(ativo, comp, cod, fi=1.0, **k):
    r = {"ativo": ativo, "desc_ativo": "d", "componente": comp, "projeto": "", "mdo": "MATERIAL", "codigo": cod, "desc_codigo": f"d {cod}",
         "fator_i": fi, "fator_r": fi, "filtro": "", "origem": "", "lv": "A"}
    r.update(k)
    return r


CERCA = [L("CERCA=CERCA(fios)", "", ""), L("CERCA", "CCOM", "poste"),
         L("CERCA(fios=%F)", "CTODOS", "arame", 0.0, formula_i="=NUM(fios)", formula_r="=NUM(fios)"),
         L("CERCA(fios=%F)", "CTODOS", "grampo", 1.0)]


def calcular(base, texto, op="I"):
    r = calc.processar_calculo([], [{"ativo": texto, "operacao": op, "entidade": "0"}], "", base)
    return {x["codigo"]: x["total"] for x in r["resultado"]}, r["nao_encontrados"], r.get("avisos")


def test_fator_pela_variante_do_pedido():
    assert calcular(CERCA, "10-CERCA(8F)")[:2] == ({"arame": 80.0, "grampo": 10.0, "poste": 10.0}, [])
    assert calcular(CERCA, "10-CERCA(12F)")[0]["arame"] == 120.0
    assert calcular(CERCA, "10-CERCA(fios=14F)")[0]["arame"] == 140.0


def test_pedido_sem_o_parametro_so_calcula_o_que_nao_depende_dele():
    tot, nao, avisos = calcular(CERCA, "10-CERCA")
    assert tot == {"poste": 10.0} and nao == [] and any("faltou informar" in a for a in avisos)


def test_padrao_da_declaracao_entra_na_formula():
    base = [L("CERCA=CERCA(fios=8F)", "", "")] + CERCA[1:]
    assert calcular(base, "10-CERCA")[0]["arame"] == 80.0


def test_remocao_usa_a_formula_do_fator_r():
    base = [L("CERCA=CERCA(fios)", "", ""), L("CERCA(fios=%F)", "C", "arame", 0.0, formula_i="=NUM(fios)", formula_r="=NUM(fios)/2")]
    r = calc.processar_calculo([], [{"ativo": "10-CERCA(8F)", "operacao": "R", "entidade": "0"}], "", base)
    assert [(x["operacao"], x["total"]) for x in r["resultado"]] == [("R", 40.0)]


def test_arred_vale_sobre_o_resultado_da_formula():
    base = [L("CERCA=CERCA(fios)", "", ""), L("CERCA(fios=%F)", "C", "arame", 0.0, formula_i="=NUM(fios)/8", arred="CIMA")]
    assert calcular(base, "1-CERCA(10F)")[0] == {"arame": 2.0}               # 10/8 = 1,25 → 2
    assert calcular(base, "1-CERCA(16F)")[0] == {"arame": 2.0}


def test_formula_com_erro_avisa_e_a_linha_nao_soma():
    base = [L("CERCA=CERCA(fios)", "", ""), L("CERCA(fios=%F)", "C", "arame", 0.0, formula_i="=fios*2"), L("CERCA(fios=%F)", "C", "grampo", 1.0)]
    tot, nao, avisos = calcular(base, "10-CERCA(8F)")
    assert tot == {"grampo": 10.0} and any("use NUM(fios)" in a and "arame" in a for a in avisos)


def test_sem_formula_nada_muda():
    base = [L("X", "C", "k", 1.5), L("Y", "D", "j", 0.333)]
    assert calcular(base, "3-X 7-Y")[0] == {"k": 4.5, "j": 2.33}
    assert calcular([dict(r, formula_i="", formula_r=None) for r in base], "3-X 7-Y")[0] == {"k": 4.5, "j": 2.33}


# ── validação da base ──
def test_validar_base_formula():
    ok = ap.validar_base([dict(r) for r in CERCA])
    assert ok["erros"] == []
    sintaxe = ap.validar_base([L("CERCA=CERCA(fios)", "", ""), L("CERCA(fios=8F)", "C", "k", 0.0, formula_i="=NUM(fios")])
    assert any("FATOR I" in e and "incompleta" in e for e in sintaxe["erros"])
    outro = ap.validar_base([L("CERCA=CERCA(fios)", "", ""), L("CERCA(fios=8F)", "C", "k", 0.0, formula_r="=NUM(poste)")])
    assert outro["erros"] == [] and any("FATOR R" in e and "'poste'" in e and "fios" in e and "não declara" in e for e in outro["avisos"])      # parâmetro não declarado: só aviso
    sem_param = ap.validar_base([L("X", "C", "k", 0.0, formula_i="=NUM(fios)")])
    assert sem_param["erros"] == [] and any("'fios'" in e and "nenhum" in e for e in sem_param["avisos"])
    do_grupo_generico = ap.validar_base([L("CERCA=CERCA(fios)", "", ""), L("CERCA(fios=8F)", "C", "k", 0.0, formula_i="=NUM(fios)", proj="PE")])
    assert do_grupo_generico["erros"] == [] and do_grupo_generico["avisos"] == []              # declaração genérica vale para a linha do projeto
    constante = ap.validar_base([L("X", "C", "k", 0.0, formula_i="=2*3")])
    assert constante["erros"] == []                                          # fórmula sem parâmetro vale em ativo simples


# ── banco e importações ──
@pytest.fixture
def banco(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    return tmp_path / "t.db"


def test_migracao_cria_as_colunas_e_e_idempotente(banco):
    con = sqlite3.connect(banco)
    for t in ("tabela_orcamento", "tabela_orcamento_master"):
        con.execute(f"CREATE TABLE {t} (id INTEGER PRIMARY KEY AUTOINCREMENT, ativo TEXT, desc_ativo TEXT, componente TEXT, projeto TEXT, mdo TEXT, codigo TEXT, desc_codigo TEXT, fator_i REAL, fator_r REAL, filtro TEXT, origem TEXT)")
        con.execute(f"INSERT INTO {t} (ativo, codigo, fator_i) VALUES ('A1', 'c1', 2.5)")
    con.commit(); con.close()
    database.init_db(); database.init_db()
    con = sqlite3.connect(banco)
    for t in ("tabela_orcamento", "tabela_orcamento_master"):
        assert con.execute(f"SELECT fator_i, formula_i, formula_r FROM {t} WHERE ativo = 'A1'").fetchone() == (2.5, "", "")


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


def fatores(tabela="tabela_orcamento_master"):
    con = database.get_connection()
    try:
        return sorted(tuple(x) for x in con.execute(f"SELECT codigo, fator_i, formula_i, fator_r, formula_r FROM {tabela}").fetchall())
    finally:
        con.close()


def csv_(*linhas):
    return (CAB + "\n" + "\n".join(linhas) + "\n").encode("utf-8-sig")


def test_upload_master_aceita_numero_e_formula(client):
    dados = csv_('CERCA=CERCA(fios);d;;P;MATERIAL;;d;;;', 'CERCA(fios=%F);d;C1;P;MATERIAL;arame;d;"=NUM(fios)";"=NUM(fios)/2";',
                 'CERCA(fios=%F);d;C1;P;MATERIAL;grampo;d;1,5;2;')
    assert client.post("/api/admin/upload-master", files={"file": ("b.csv", dados, "text/csv")}, headers=adm()).status_code == 200
    f = {r[0]: r for r in fatores() if r[0]}
    assert f["arame"] == ("arame", 0.0, "=NUM(fios)", 0.0, "=NUM(fios)/2") and f["grampo"] == ("grampo", 1.5, "", 2.0, "")


@pytest.mark.parametrize("fator, trecho", [("=NUM(fios", "FATOR I"), ("=1+", "incompleta")])
def test_upload_recusa_formula_ruim_sem_alterar_nada(client, fator, trecho):
    ok = csv_("X;d;C;P;MATERIAL;k1;d;1;1;")
    client.post("/api/admin/upload-master", files={"file": ("b.csv", ok, "text/csv")}, headers=adm())
    ruim = csv_('CERCA=CERCA(fios);d;;P;MATERIAL;;d;;;', f'CERCA(fios=8F);d;C;P;MATERIAL;k2;d;"{fator}";1;')
    r = client.post("/api/admin/upload-master", files={"file": ("b.csv", ruim, "text/csv")}, headers=adm())
    assert r.status_code == 400 and trecho in r.json()["detail"], r.text
    assert [x[0] for x in fatores()] == ["k1"]


def test_parametro_so_na_formula_salva_com_aviso(client):
    dados = csv_('CERCA=CERCA(fios);d;;P;MATERIAL;;d;;;', 'CERCA(fios=8F);d;C1;P;MATERIAL;arame;d;"=NUM(fios)*NUM(poste)";1;')
    r = client.post("/api/admin/upload-master", files={"file": ("b.csv", dados, "text/csv")}, headers=adm())
    assert r.status_code == 200, r.text
    assert any("'poste'" in a and "não declara" in a for a in r.json().get("avisos", []))
    assert [x for x in fatores() if x[0] == "arame"] == [("arame", 0.0, "=NUM(fios)*NUM(poste)", 1.0, "")]
    r = client.post("/api/orcamento/upload", files={"file": ("b.csv", dados, "text/csv")}, headers=adm())
    assert r.status_code == 200 and any("'poste'" in a for a in r.json().get("avisos", []))
    linhas = [{"ativo": "CERCA=CERCA(fios)", "projeto": "P"}, {"ativo": "CERCA(fios=8F)", "componente": "C", "projeto": "P", "codigo": "arame", "fator_i": "=NUM(poste)", "fator_r": "1"}]
    assert client.post("/api/orcamento/salvar", json={"dados": linhas, "projeto": "P"}, headers=adm()).status_code == 200
    assert client.post("/api/admin/sync-master-all", json={"dados": linhas, "projeto": "P"}, headers=adm()).status_code == 200


def test_formula_que_usa_parametro_nao_informado_avisa_no_calculo():
    base = [L("CERCA=CERCA(fios)", "", ""), L("CERCA(fios=%F)", "C", "arame", 0.0, formula_i="=NUM(fios)*NUM(poste)"), L("CERCA(fios=%F)", "C", "grampo", 1.0)]
    tot, nao, avisos = calcular(base, "10-CERCA(8F)")
    assert tot == {"grampo": 10.0} and any("poste" in a and "não informado" in a for a in avisos)


def test_upload_local_e_salvar_da_tela(client):
    dados = csv_('CERCA=CERCA(fios);d;;P;MATERIAL;;d;;;', 'CERCA(fios=%F);d;C1;P;MATERIAL;arame;d;=NUM(fios);1;')
    assert client.post("/api/orcamento/upload", files={"file": ("b.csv", dados, "text/csv")}, headers=adm()).status_code == 200
    assert [x for x in fatores("tabela_orcamento") if x[0] == "arame"] == [("arame", 0.0, "=NUM(fios)", 1.0, "")]
    linhas = [{"ativo": "CERCA=CERCA(fios)", "projeto": "P", "fator_i": 0, "fator_r": 0},
              {"ativo": "CERCA(fios=%F)", "componente": "C1", "projeto": "P", "codigo": "arame", "fator_i": "=NUM(fios)*2", "fator_r": "1,5"},
              {"ativo": "X", "componente": "C2", "projeto": "P", "codigo": "k", "fator_i": "2,5", "fator_r": ""}]
    r = client.post("/api/orcamento/salvar", json={"dados": linhas, "projeto": "P"}, headers=adm())
    assert r.status_code == 200, r.text
    salvos = [x for x in fatores("tabela_orcamento") if x[0] in ("arame", "k")]
    assert ("arame", 0.0, "=NUM(fios)*2", 1.5, "") in salvos and ("k", 2.5, "", 0.0, "") in salvos
    ruim = [{"ativo": "X", "componente": "C", "projeto": "P", "codigo": "k", "fator_i": "=1+"}]
    r = client.post("/api/orcamento/salvar", json={"dados": ruim, "projeto": "P"}, headers=adm())
    assert r.status_code == 400 and "Linha 1" in r.json()["detail"]
    dados_api = client.get("/api/orcamento/dados", headers=adm()).json()["dados"]
    assert dados_api and all("formula_i" in d and "formula_r" in d for d in dados_api)        # a tela recebe as colunas


def test_add_row_e_sync_all(client):
    assert client.post("/api/admin/add-row", json={"ativo": "CERCA=CERCA(fios)", "projeto": "P"}, headers=adm()).status_code == 200
    r = client.post("/api/admin/add-row", json={"ativo": "CERCA(fios=8F)", "componente": "C", "projeto": "P", "codigo": "arame", "fator_i": "=NUM(fios)", "fator_r": 2}, headers=adm())
    assert r.status_code == 200, r.text
    assert client.post("/api/admin/add-row", json={"ativo": "CERCA(fios=9F)", "componente": "C", "projeto": "P", "codigo": "k9", "fator_i": "=NUM(x)"}, headers=adm()).status_code == 200      # parâmetro não declarado: salva (só aviso)
    assert client.post("/api/admin/add-row", json={"ativo": "CERCA(fios=7F)", "componente": "C", "projeto": "P", "codigo": "k7", "fator_i": "=NUM("}, headers=adm()).status_code == 400      # sintaxe inválida: recusa
    assert [x for x in fatores() if x[0] == "arame"] == [("arame", 0.0, "=NUM(fios)", 2.0, "")]
    linhas = [{"ativo": "CERCA=CERCA(fios)", "projeto": "P"}, {"ativo": "CERCA(fios=8F)", "componente": "C", "projeto": "P", "codigo": "arame", "fator_i": "=NUM(fios)/4", "fator_r": "1"}]
    assert client.post("/api/admin/sync-master-all", json={"dados": linhas, "projeto": "P"}, headers=adm()).status_code == 200
    assert [x for x in fatores() if x[0] == "arame"] == [("arame", 0.0, "=NUM(fios)/4", 1.0, "")]


def test_tela_da_base_aceita_formula():
    html = open("static/orcamento.html", encoding="utf-8").read()
    assert "fatorTela(" in html and 'type="number" step="0.01"' not in html and "fator_i: inputs[8].value.trim() || 0" in html
