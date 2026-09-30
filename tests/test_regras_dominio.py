"""tests/test_regras_dominio.py — camada 2: as 11 regras da semente v2 (TASK-013/016), schema e rotas.

Cada teste cita a decisão do usuário. A equivalência com o motor v1 está em tests/test_regras_v2.py."""
import copy
import json

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from config import REGRAS_DOMINIO_SEED_PATH
from middleware.auth_middleware import create_jwt_token
from services.regras_dominio import avaliar, validar_conjunto

with open(REGRAS_DOMINIO_SEED_PATH, encoding="utf-8") as f:
    SEED = json.load(f)
SEMENTE = SEED["regras"]
GRUPOS = SEED["grupos"]


def ativas(*ids):
    """Semente com só as regras pedidas ligadas (a semente vem toda desligada)."""
    regras = copy.deepcopy(SEMENTE)
    for r in regras:
        r["ativa"] = r["id"] in ids
    return regras


def linha(ativo, op="I"):
    return {"ativo": ativo, "operacao": op}


def ids(achados):
    return sorted(a["regra_id"] for a in achados)


def rodar(regra_id, outros, cabos=()):
    return avaliar(ativas(regra_id), list(cabos), [linha(*o) if isinstance(o, tuple) else linha(o) for o in outros], GRUPOS)


# ── semente ──────────────────────────────────────────────────────────────────
def test_semente_e_valida_e_vem_toda_desligada():
    assert validar_conjunto(SEMENTE, GRUPOS) == []
    assert len(SEMENTE) == 11 and not any(r["ativa"] for r in SEMENTE)


def test_regras_inativas_nao_geram_achado():
    assert avaliar(SEMENTE, [], [linha("DT11/300 1-CFU")], GRUPOS) == []


# ── CFU exige SUPL; CFU e CFUR exigem EF ─────────────────────────────────────
def test_cfu_supl_basta_presenca_na_linha_em_qualquer_posicao():
    assert rodar("C2-CFU-SUPL", ["DT11/300 1-CFU"]) != []
    assert rodar("C2-CFU-SUPL", ["DT11/300 1-CFU 1-SUPL"]) == []
    assert rodar("C2-CFU-SUPL", ["DT11/300 1-SUPL 1-CFU"]) == []
    assert rodar("C2-CFU-SUPL", ["DT11/300 1-CFUR"]) == []  # SUPL vale só para CFU


def test_chave_exige_elo_fusivel_cfu_e_cfur_mas_nao_cfa():
    assert ids(rodar("C2-CFU-EF", ["DT11/300 1-CFU"])) == ["C2-CFU-EF"]
    assert ids(rodar("C2-CFU-EF", ["DT11/300 1-CFUR"])) == ["C2-CFU-EF"]
    assert rodar("C2-CFU-EF", ["DT11/300 1-CFA"]) == []
    assert rodar("C2-CFU-EF", ["DT11/300 1-CFU 1-EF3H"]) == []


# ── trafo: PR15 e PR220 (TR1xx mono, TR2xx bi como mono, TR3xx tri) ──────────
def test_trafo_exige_pr15_exceto_com_rpr():
    assert ids(rodar("C2-TR-PR15", ["DT11/300 1-TR125"])) == ["C2-TR-PR15"]
    assert rodar("C2-TR-PR15", ["DT11/300 1-TR125 1-PR15"]) == []
    assert rodar("C2-TR-PR15", ["DT11/300 1-TR125 1-RPR"]) == []


@pytest.mark.parametrize("trafo", ["TR125", "TR215"])  # monofásico e bifásico usam a mesma regra
def test_pr220_minimo_2_para_mono_e_bi(trafo):
    assert rodar("C2-TR-PR220-MONO", [f"DT11/300 1-{trafo} 1-PR220"]) != []
    assert rodar("C2-TR-PR220-MONO", [f"DT11/300 1-{trafo} 2-PR220"]) == []


def test_pr220_minimo_3_para_trifasico():
    assert rodar("C2-TR-PR220-TRI", ["DT11/300 1-TR330 2-PR220"]) != []
    assert rodar("C2-TR-PR220-TRI", ["DT11/300 1-TR330 3-PR220"]) == []
    assert rodar("C2-TR-PR220-TRI", ["DT11/300 1-TR125 1-PR220"]) == []  # mono não cai na regra tri


@pytest.mark.parametrize("descidas,pr220,ok", [
    ("1-SI3", 2, True),                # uma descida: mínimo 2
    ("1-SI3 1-SI2", 2, True),          # outra estrutura de BT: continua uma descida
    ("1-SI4", 2, False),               # duas descidas: dobra para 4
    ("1-SI4", 4, True),
    ("2-SI3", 2, False),               # 2-SI3 = duas descidas
    ("3-SI3", 3, False),               # decisão: SI3 com qtd >= 2 também conta (3 SI3 -> duas descidas)
    ("3-SI3", 4, True),
])
def test_duas_descidas_dobram_o_pr220(descidas, pr220, ok):
    achados = rodar("C2-TR-PR220-MONO", [f"DT11/300 1-TR125 {pr220}-PR220 {descidas}"])
    assert (achados == []) is ok


def test_pr220_trifasico_tambem_dobra_com_duas_descidas():
    assert rodar("C2-TR-PR220-TRI", ["DT11/300 1-TR330 3-PR220 1-SI4"]) != []
    assert rodar("C2-TR-PR220-TRI", ["DT11/300 1-TR330 6-PR220 1-SI4"]) == []


# ── P50 pelo total da planilha, comprimento bruto, só I/*I ───────────────────
def test_p50_total_da_planilha_soma_trafos_e_pr15():
    outros = ["DT11/300 1-TR125 1-PR15", "DT11/300 1-TR330 1-PR15"]  # 2 + 6 + 2 = 10 m
    assert ids(rodar("C2-P50", outros, [linha("P50 ABC 8 m")])) == ["C2-P50"]
    assert rodar("C2-P50", outros, [linha("P50 ABC 10 m")]) == []


def test_p50_usa_comprimento_bruto_sem_multiplicar_por_fases():
    # "P50 ABC 2 m" tem 3 fases, mas conta 2 m (decisão do usuário), então não cobre um trafo trifásico (6 m)
    assert rodar("C2-P50", ["DT11/300 1-TR330"], [linha("P50 ABC 2 m")]) != []
    assert rodar("C2-P50", ["DT11/300 1-TR330"], [linha("P50 ABC 6 m")]) == []


def test_p50_soma_varias_linhas_de_cabo_e_ignora_outros_cabos():
    cabos = [linha("P50 A 3 m"), linha("P50 B 3 m"), linha("CAA 2 ABC 100 m")]
    assert rodar("C2-P50", ["DT11/300 1-TR330"], cabos) == []


def test_p50_sem_trafo_nem_pr15_nao_exige_nada():
    assert rodar("C2-P50", ["DT11/300 1-CFU"], []) == []


def test_p50_achado_e_unico_e_geral():
    achados = rodar("C2-P50", ["DT11/300 1-TR125", "DT11/300 1-TR125"], [])
    assert len(achados) == 1 and achados[0]["linha_id"] == "GERAL"


def test_regras_so_valem_para_instalacao_por_padrao():
    assert rodar("C2-CFU-SUPL", [("DT11/300 1-CFU", "R")]) == []
    assert rodar("C2-CFU-SUPL", [("DT11/300 1-CFU", "*I")]) != []
    assert rodar("C2-P50", [("DT11/300 1-TR330", "R")], []) == []


def test_regra_sem_campo_operacoes_assume_so_instalacao():
    regras = ativas("C2-CFU-SUPL")
    del regras[0]["operacoes"]
    assert avaliar(regras, [], [linha("DT11/300 1-CFU", "R"), linha("DT11/300 1-CFU", "M")]) == []
    assert avaliar(regras, [], [linha("DT11/300 1-CFU", "I")]) != []


def test_operacoes_da_regra_sao_editaveis():
    regras = ativas("C2-CFU-SUPL")
    regras[0]["operacoes"] = ["I", "*I", "R"]
    assert avaliar(regras, [], [linha("DT11/300 1-CFU", "R")], GRUPOS) != []


# ── poste de 10 m em MT ──────────────────────────────────────────────────────
@pytest.mark.parametrize("linha_", [
    "DT10/300 1-U1", "DT10/300 1-N4", "DT10/300 1-T2", "DT10/300 1-TE", "DT10/300 1-R2",
    "DT10/300 1-U1C", "DT10/300 1-N1IV",           # variantes
    "DT10/300 1-CFU", "DT10/300 1-TR125",          # chave/trafo de MT
    "CV10/300 1-U1", "1-4DT10/300 1-U1",  # poste com prefixo numérico, forma lida pelo cálculo
])
def test_poste_10m_em_mt_e_erro(linha_):
    achados = rodar("C2-POSTE10-MT", [linha_])
    assert ids(achados) == ["C2-POSTE10-MT"] and achados[0]["severidade"] == "erro"


@pytest.mark.parametrize("linha_", [
    "DT10/300 1-SI3",        # só baixa tensão
    "DT11/300 1-U1",         # 11 m pode
    "DT10/300",              # sem nada de MT
])
def test_poste_10m_sem_mt_ou_poste_maior_nao_gera_alerta(linha_):
    assert rodar("C2-POSTE10-MT", [linha_]) == []


# ── estrutura isolada ────────────────────────────────────────────────────────
@pytest.mark.parametrize("isolada", ["U3", "N3", "R3", "T3", "U3C", "N3IV"])
def test_estrutura_isolada_sozinha_gera_aviso(isolada):
    assert ids(rodar("C2-ESTR-ISOL", [f"DT11/300 1-{isolada}"])) == ["C2-ESTR-ISOL"]


@pytest.mark.parametrize("linha_", [
    "DT11/300 1-U3 1-U1",     # outra estrutura MT
    "DT11/300 1-U3 1-N3",     # decisão: mesmo outra isolada satisfaz
    "DT11/300 1-U3 1-TR125",  # trafo
    "DT11/300 2-U3",          # duas unidades
])
def test_estrutura_isolada_acompanhada_esta_ok(linha_):
    assert rodar("C2-ESTR-ISOL", [linha_]) == []


def test_estrutura_de_baixa_tensao_nao_acompanha():
    assert rodar("C2-ESTR-ISOL", ["DT11/300 1-U3 1-SI3"]) != []


# ── formato do poste ─────────────────────────────────────────────────────────
@pytest.mark.parametrize("linha_", ["1-DT11/300 1-CFU", "POSTE11 1-CFU", "DT11 1-CFU"])
def test_poste_fora_do_padrao_e_apenas_info(linha_):
    achados = rodar("C2-POSTE-FMT", [linha_])
    assert ids(achados) == ["C2-POSTE-FMT"] and achados[0]["severidade"] == "info"


def test_poste_no_padrao_nao_gera_info():
    assert rodar("C2-POSTE-FMT", ["DT11/300 1-CFU", "CV11/600"]) == []


# ── trafo com elo fusível só se houver chave (CFU ou CFUR) ────────────────────
def test_trafo_com_elo_sem_chave_gera_aviso():
    achados = rodar("C2-TR-EF", ["DT11/300 1-TR125 1-EF3H"])
    assert ids(achados) == ["C2-TR-EF"] and achados[0]["severidade"] == "aviso"


@pytest.mark.parametrize("linha_", [
    "DT11/300 1-TR125 1-EF3H 1-CFU",    # chave libera o elo
    "DT11/300 1-TR125 1-EF3H 1-CFUR",
    "DT11/300 1-TR125",                  # trafo sem elo
    "DT11/300 1-CFU 1-EF3H",             # elo sem trafo (é a regra CFU->EF, não esta)
    "DT11/300 1-TR330 1-PR15",
])
def test_trafo_e_elo_sem_problema(linha_):
    assert rodar("C2-TR-EF", [linha_]) == []


@pytest.mark.parametrize("outra", ["1-CFA", "1-CL", "1-RCFU", "1-ACFU"])
def test_so_cfu_e_cfur_liberam_o_elo(outra):
    assert ids(rodar("C2-TR-EF", [f"DT11/300 1-TR125 1-EF3H {outra}"])) == ["C2-TR-EF"]


def test_trafo_de_qualquer_tipo_e_elo_de_qualquer_corrente():
    for trafo in ("TR105", "TR215", "TR330A"):
        assert rodar("C2-TR-EF", [f"DT11/300 1-{trafo} 1-EF10K"]) != []


# ── estrutura SI exige RA2, salvo estrutura S# ────────────────────────────────
@pytest.mark.parametrize("si", ["SI1", "SI2", "SI3", "SI4"])
def test_estrutura_si_sem_ra2_gera_aviso(si):
    achados = rodar("C2-BT-EXT", [f"DT11/300 1-{si}"])
    assert ids(achados) == ["C2-BT-EXT"] and achados[0]["severidade"] == "aviso"


@pytest.mark.parametrize("linha_", [
    "DT11/300 1-SI3 1-RA2",              # exemplo do prompt: SI3 + RA2
    "DT11/300 1-SI4 1-SI3 1-RA2",        # exemplo do prompt: SI4 + SI3 + RA2
    "DT11/300 1-SI3 1-S2",               # estrutura S# dispensa o RA2
    "DT11/300 1-SI4 1-S4",
    "DT11/300 1-SI1 1-S44",
    "DT11/300 1-RA2",                    # RA2 sem SI: nada a exigir
    "DT11/300 1-S2",
    "DT11/300 1-U1",
])
def test_estrutura_si_com_ra2_ou_s_esta_ok(linha_):
    assert rodar("C2-BT-EXT", [linha_]) == []


def test_si_nao_e_confundido_com_s_numerico_nem_com_supl():
    assert rodar("C2-BT-EXT", ["DT11/300 1-SI3 1-SUPL"]) != []      # SUPL não é estrutura S#
    assert rodar("C2-BT-EXT", ["DT11/300 1-SI3 1-SI4"]) != []       # SI não é S#


# ── schema (v2) ──────────────────────────────────────────────────────────────
def erros_de(mutacao):
    regras = copy.deepcopy(SEMENTE)
    mutacao(regras)
    return validar_conjunto(regras, GRUPOS)


def regra(regras, rid):
    return next(x for x in regras if x["id"] == rid)


@pytest.mark.parametrize("mutacao,trecho", [
    (lambda r: r[0].update(severidade="grave"), "'severidade'"),
    (lambda r: r[0].update(ativa="sim"), "'ativa'"),
    (lambda r: r[0].update(mensagem=""), "'mensagem'"),
    (lambda r: r[0].update(id=""), "'id' é obrigatório"),
    (lambda r: r[1].update(id=r[0]["id"]), "duplicado"),
    (lambda r: r[0].update(operacoes=["X"]), "'operacoes'"),
    (lambda r: r[0].update(operacoes=[]), "'operacoes'"),
    (lambda r: r[0].update(versao=3), "'versao'"),
    (lambda r: r[0].update(escopo="tudo"), "'escopo'"),
    (lambda r: r[0].update(quandu={"tem": "X"}), "campo desconhecido 'quandu'"),
    (lambda r: (r[0].pop("quando"), r[0].pop("entao")), "informe 'quando' e/ou 'entao'"),
    (lambda r: r[0].update(quando={"tem": "@NAO_EXISTE"}), "grupo '@NAO_EXISTE' não existe"),
    (lambda r: r[0].update(quando={"tem": {"regex": "(["}}), "regex inválida"),
    (lambda r: r[0].update(quando={"tem": "CFU", "algum": []}), "condição inválida"),
    (lambda r: r[0].update(quando={"inventada": "X"}), "condição inválida"),
    (lambda r: r[0].update(quando={"todos": []}), "ao menos uma condição"),
    (lambda r: r[0].update(entao={"soma": "SUPL"}), "'soma' exige 'qtd'"),
    (lambda r: r[0].update(entao={"soma": "SUPL", "qtd": {"parecido": 1}}), "comparador 'parecido' inválido"),
    (lambda r: r[0].update(entao={"soma": "SUPL", "qtd": {">=": "dois"}}), "número"),
    (lambda r: r[0].update(entao={"soma": "SUPL", "qtd": {"entre": [1]}}), "[mínimo, máximo]"),
    (lambda r: r[0].update(entao={"soma": "SUPL", "qtd": {">=": {"base": 1, "vezes": 2}}}), "só têm efeito com 'se'"),
    (lambda r: r[0].update(quando={"texto": ""}), "regex obrigatória"),
    (lambda r: r[0].update(quando={"texto": "x" * 501}), "longa demais"),
    (lambda r: r[0].update(deve_ser={"esq": 1, "cmp": ">=", "dir": 2}), "é só de regra de planilha"),
    (lambda r: regra(r, "C2-P50").update(quando={"tem": "X"}), "é só de regra de linha"),
    (lambda r: regra(r, "C2-P50")["deve_ser"].update(cmp="entre"), "deve_ser.cmp"),
    (lambda r: regra(r, "C2-P50")["deve_ser"].update(dir={"soma_metrosX": "P50"}), "desconhecida"),
    (lambda r: regra(r, "C2-P50")["deve_ser"].update(dir={"vezes": ["dois", {"soma_qtd": "PR15"}]}), "[número, expressão]"),
    (lambda r: regra(r, "C2-P50").pop("deve_ser"), "'deve_ser' precisa ser"),
])
def test_schema_rejeita(mutacao, trecho):
    assert any(trecho in e for e in erros_de(mutacao)), erros_de(mutacao)


def test_schema_rejeita_payload_que_nao_e_lista():
    from services.regras_dominio import validar_regras
    assert validar_regras({"a": 1}) == ["O payload de regras precisa ser uma lista."]


def test_regra_v1_antiga_continua_valida_e_e_validada_como_v1():
    v1 = {"id": "V1", "escopo": "outros", "tipo": "requer", "severidade": "aviso", "mensagem": "m", "ativa": True,
          "parametros": {"se_regex": "^CFU$", "exige_regex": "^SUPL$", "qtd_min": 1}}
    assert validar_conjunto([v1], {}) == []
    v1["parametros"]["se_regex"] = "(["
    assert any("regex inválida" in e for e in validar_conjunto([v1], {}))


# ── rotas (banco temporário) ─────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(role):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


def test_get_devolve_semente_default(client):
    r = client.get("/api/validacao/regras", headers=cab("operador")).json()
    assert r["versao_de"] == "DEFAULT" and r["personalizado"] is False and len(r["regras"]) == 11


def test_escrita_e_testar_so_admin(client):
    for metodo, url, corpo in [
        ("post", "/api/validacao/regras", {"projeto_codigo": "229", "regras": SEMENTE}),
        ("post", "/api/validacao/regras/testar", {"regras": SEMENTE, "grupos": GRUPOS}),
        ("post", "/api/validacao/regras/restaurar-semente", {"projeto_codigo": "229"}),
        ("get", "/api/validacao/regras/historico", None),
    ]:
        resp = getattr(client, metodo)(url, headers=cab("operador"), **({"json": corpo} if corpo else {}))
        assert resp.status_code == 403, url


def test_admin_salva_regra_ligada_por_projeto_sem_tocar_no_default(client):
    regras = ativas("C2-CFU-SUPL")
    assert client.post("/api/validacao/regras", json={"projeto_codigo": "229", "regras": regras}, headers=cab("admin")).status_code == 200
    do_projeto = client.get("/api/validacao/regras?projeto_codigo=229", headers=cab("operador")).json()
    assert do_projeto["personalizado"] is True and [r["id"] for r in do_projeto["regras"] if r["ativa"]] == ["C2-CFU-SUPL"]
    outro = client.get("/api/validacao/regras?projeto_codigo=027", headers=cab("operador")).json()
    assert outro["personalizado"] is False and not any(r["ativa"] for r in outro["regras"])


def test_regras_invalidas_nao_sao_salvas(client):
    regras = copy.deepcopy(SEMENTE)
    regras[0]["quando"] = {"tem": {"regex": "(["}}
    r = client.post("/api/validacao/regras", json={"projeto_codigo": "229", "regras": regras}, headers=cab("admin"))
    assert r.status_code == 400 and any("regex" in e for e in r.json()["detail"]["erros"])
    assert client.get("/api/validacao/regras?projeto_codigo=229", headers=cab("admin")).json()["personalizado"] is False


def test_historico_reversao_e_semente(client):
    a, b = ativas("C2-CFU-SUPL"), ativas("C2-CFU-EF")
    for regras in (a, b):
        client.post("/api/validacao/regras", json={"projeto_codigo": "229", "regras": regras}, headers=cab("admin"))
    hist = client.get("/api/validacao/regras/historico?projeto_codigo=229", headers=cab("admin")).json()["historico"]
    assert len(hist) == 1 and hist[0]["criado_por"] == "admin@x.com"
    assert client.post("/api/validacao/regras/reverter", json={"projeto_codigo": "229", "historico_id": hist[0]["id"]},
                       headers=cab("admin")).status_code == 200
    atual = client.get("/api/validacao/regras?projeto_codigo=229", headers=cab("admin")).json()["regras"]
    assert [r["id"] for r in atual if r["ativa"]] == ["C2-CFU-SUPL"]
    assert client.post("/api/validacao/regras/reverter", json={"projeto_codigo": "229", "historico_id": 999},
                       headers=cab("admin")).status_code == 404
    assert client.post("/api/validacao/regras/restaurar-semente", json={"projeto_codigo": "229"}, headers=cab("admin")).status_code == 200
    assert not any(r["ativa"] for r in client.get("/api/validacao/regras?projeto_codigo=229", headers=cab("admin")).json()["regras"])


def test_testar_roda_rascunho_sem_salvar(client):
    r = client.post("/api/validacao/regras/testar", headers=cab("admin"),
                    json={"regras": ativas("C2-CFU-SUPL"), "grupos": GRUPOS, "cabos": [], "outros": [linha("DT11/300 1-CFU")]})
    assert r.status_code == 200 and ids(r.json()["achados"]) == ["C2-CFU-SUPL"] and r.json()["resumo"]["aviso"] == 1
    assert client.get("/api/validacao/regras", headers=cab("admin")).json()["personalizado"] is False
    ruim = copy.deepcopy(SEMENTE)
    ruim[0]["severidade"] = "grave"
    assert client.post("/api/validacao/regras/testar", json={"regras": ruim, "grupos": GRUPOS}, headers=cab("admin")).status_code == 400


def test_rota_de_planilhas_inclui_a_camada_2_do_projeto(client):
    client.post("/api/validacao/regras", json={"projeto_codigo": "229", "regras": ativas("C2-CFU-SUPL")}, headers=cab("admin"))
    corpo = {"cabos": [], "outros": [linha("DT11/300 1-CFU")], "projeto_codigo": "229"}
    com = client.post("/api/validacao/planilhas", json=corpo, headers=cab("operador")).json()
    assert ids(com["achados"]) == ["C2-CFU-SUPL"]
    sem_dominio = client.post("/api/validacao/planilhas", json={**corpo, "incluir_dominio": False}, headers=cab("operador")).json()
    assert sem_dominio["achados"] == []
    outro_projeto = client.post("/api/validacao/planilhas", json={**corpo, "projeto_codigo": "027"}, headers=cab("operador")).json()
    assert outro_projeto["achados"] == []  # 027 herda o DEFAULT, que vem desligado


def test_seed_e_idempotente_e_nao_sobrescreve_edicao(client):
    client.post("/api/validacao/regras", json={"projeto_codigo": "DEFAULT", "regras": ativas("C2-CFU-EF")}, headers=cab("admin"))
    database.init_db()
    atual = client.get("/api/validacao/regras", headers=cab("admin")).json()["regras"]
    assert [r["id"] for r in atual if r["ativa"]] == ["C2-CFU-EF"]


# ── adicionar regras novas da semente ────────────────────────────────────────
def instalacao_antiga(client, projeto="DEFAULT"):
    """Simula quem já usa o sistema: só as 9 primeiras regras, uma delas ligada e editada pelo admin."""
    regras = copy.deepcopy(SEMENTE)[:9]
    regras[0]["ativa"] = True
    regras[0]["mensagem"] = "Mensagem editada pelo admin"
    assert client.post("/api/validacao/regras", json={"projeto_codigo": projeto, "regras": regras}, headers=cab("admin")).status_code == 200


def test_adicionar_novas_acrescenta_so_o_que_falta_desligado_e_preserva_o_resto(client):
    instalacao_antiga(client)
    r = client.post("/api/validacao/regras/adicionar-novas", json={"projeto_codigo": "DEFAULT"}, headers=cab("admin"))
    assert r.status_code == 200 and r.json()["adicionadas"] == ["C2-TR-EF", "C2-BT-EXT"]
    regras = client.get("/api/validacao/regras", headers=cab("admin")).json()["regras"]
    assert len(regras) == 11
    por_id = {x["id"]: x for x in regras}
    assert por_id["C2-CFU-SUPL"]["ativa"] is True and por_id["C2-CFU-SUPL"]["mensagem"] == "Mensagem editada pelo admin"
    assert por_id["C2-TR-EF"]["ativa"] is False and por_id["C2-BT-EXT"]["ativa"] is False
    hist = client.get("/api/validacao/regras/historico", headers=cab("admin")).json()["historico"]
    # o banco de teste nasce com as 11 da semente; simular a instalação antiga gerou 1 entrada e adicionar-novas outra
    assert len(hist) == 2 and len(hist[0]["regras"]) == 9          # a versão anterior (9 regras) ficou no histórico


def test_adicionar_novas_e_idempotente(client):
    instalacao_antiga(client)
    client.post("/api/validacao/regras/adicionar-novas", json={"projeto_codigo": "DEFAULT"}, headers=cab("admin"))
    r = client.post("/api/validacao/regras/adicionar-novas", json={"projeto_codigo": "DEFAULT"}, headers=cab("admin"))
    assert r.json()["adicionadas"] == []
    assert len(client.get("/api/validacao/regras/historico", headers=cab("admin")).json()["historico"]) == 2  # sem nova versão


def test_adicionar_novas_na_versao_propria_do_projeto_nao_toca_no_default(client):
    instalacao_antiga(client, "229")
    client.post("/api/validacao/regras/adicionar-novas", json={"projeto_codigo": "229"}, headers=cab("admin"))
    do_projeto = client.get("/api/validacao/regras?projeto_codigo=229", headers=cab("admin")).json()
    assert do_projeto["personalizado"] is True and len(do_projeto["regras"]) == 11
    padrao = client.get("/api/validacao/regras", headers=cab("admin")).json()
    assert padrao["personalizado"] is False and len(padrao["regras"]) == 11   # o DEFAULT do banco novo já tinha as 11


def test_adicionar_novas_em_projeto_herdado_cria_a_versao_do_projeto(client):
    r = client.post("/api/validacao/regras/adicionar-novas", json={"projeto_codigo": "027"}, headers=cab("admin"))
    assert r.json()["adicionadas"] == []        # herda o DEFAULT completo: nada a acrescentar, nada é criado
    assert client.get("/api/validacao/regras?projeto_codigo=027", headers=cab("admin")).json()["personalizado"] is False


def test_adicionar_novas_so_admin(client):
    assert client.post("/api/validacao/regras/adicionar-novas", json={"projeto_codigo": "DEFAULT"}, headers=cab("operador")).status_code == 403


def test_regra_nova_da_semente_entra_sempre_desligada_mesmo_que_a_semente_a_traga_ligada(client, tmp_path, monkeypatch):
    import routers.validacao_regras as rv
    instalacao_antiga(client)
    semente = copy.deepcopy(SEED)
    semente["regras"][-1]["ativa"] = True          # C2-BT-EXT ligada na semente (não deve valer no destino)
    arquivo = tmp_path / "semente.json"
    arquivo.write_text(json.dumps(semente), encoding="utf-8")
    monkeypatch.setattr(rv, "REGRAS_DOMINIO_SEED_PATH", str(arquivo))
    client.post("/api/validacao/regras/adicionar-novas", json={"projeto_codigo": "DEFAULT"}, headers=cab("admin"))
    regras = {r["id"]: r for r in client.get("/api/validacao/regras", headers=cab("admin")).json()["regras"]}
    assert regras["C2-BT-EXT"]["ativa"] is False
