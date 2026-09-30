"""tests/test_regras_camadas.py — regras em camadas: padrão + ajustes do projeto (TASK-018).

Decisões do usuário: 1) camadas; 2) vale a do projeto quando ele sobrescreve; 3) grupos globais, redefiníveis por
projeto; 5) botão −/+ oculta/reexibe regra herdada (continua listada)."""
import copy
import json

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from config import REGRAS_DOMINIO_SEED_PATH
from middleware.auth_middleware import create_jwt_token
from services.regras_camadas import (efetivo, overlay_de_bruto, overlay_de_efetivo, overlay_vazio)
from services.regras_dominio import avaliar, normalizar_regra

with open(REGRAS_DOMINIO_SEED_PATH, encoding="utf-8") as f:
    SEED = json.load(f)
PADRAO = {"grupos": SEED["grupos"], "regras": [normalizar_regra(r) for r in SEED["regras"]]}


def linha(ativo, op="I"):
    return {"ativo": ativo, "operacao": op}


def ids(achados):
    return sorted(a["regra_id"] for a in achados)


def por_id(regras):
    return {r["id"]: r for r in regras}


# ── funções puras ────────────────────────────────────────────────────────────
def test_overlay_vazio_e_o_padrao_puro():
    regras, grupos, origem_g, avisos = efetivo(PADRAO, overlay_vazio())
    assert [r["id"] for r in regras] == [r["id"] for r in PADRAO["regras"]]
    assert all(r["origem"] == "padrao" and r["oculta"] is False for r in regras)
    assert grupos == PADRAO["grupos"] and set(origem_g.values()) == {"padrao"} and avisos == []


def test_sobrescrita_vale_a_do_projeto_e_marca_origem():
    o = overlay_vazio()
    o["sobrescritas"]["C2-CFU-SUPL"] = {"severidade": "erro", "ativa": True}
    r = por_id(efetivo(PADRAO, o)[0])["C2-CFU-SUPL"]
    assert r["severidade"] == "erro" and r["ativa"] is True and r["origem"] == "sobrescrita"
    assert por_id(efetivo(PADRAO, o)[0])["C2-CFU-EF"]["origem"] == "padrao"


def test_valor_none_remove_o_campo_sobrescrito():
    o = overlay_vazio()
    o["sobrescritas"]["C2-CFU-SUPL"] = {"operacoes": None}
    assert "operacoes" not in por_id(efetivo(PADRAO, o)[0])["C2-CFU-SUPL"]


def test_padrao_evolui_projeto_acompanha_exceto_onde_sobrescreveu():
    padrao2 = copy.deepcopy(PADRAO)
    for r in padrao2["regras"]:
        r["mensagem"] = "NOVA " + r["id"]
    padrao2["regras"].append({**copy.deepcopy(PADRAO["regras"][0]), "id": "C2-NOVA"})
    o = overlay_vazio()
    o["sobrescritas"]["C2-CFU-SUPL"] = {"mensagem": "Minha mensagem"}
    m = por_id(efetivo(padrao2, o)[0])
    assert m["C2-CFU-SUPL"]["mensagem"] == "Minha mensagem"          # vale a do projeto
    assert m["C2-CFU-EF"]["mensagem"] == "NOVA C2-CFU-EF"           # acompanha o padrão
    assert "C2-NOVA" in m                                            # regra nova do padrão chega


def test_ocultas_continuam_listadas_e_nao_disparam():
    o = overlay_vazio()
    o["sobrescritas"]["C2-CFU-SUPL"] = {"ativa": True}
    o["ocultas"] = ["C2-CFU-SUPL"]
    regras, grupos, _, _ = efetivo(PADRAO, o)
    r = por_id(regras)["C2-CFU-SUPL"]
    assert r["oculta"] is True and r["ativa"] is True
    assert avaliar(regras, [], [linha("DT11/300 1-CFU")], grupos) == []
    o["ocultas"] = []                                               # botão +
    regras, grupos, _, _ = efetivo(PADRAO, o)
    assert ids(avaliar(regras, [], [linha("DT11/300 1-CFU")], grupos)) == ["C2-CFU-SUPL"]


def test_regra_propria_e_grupo_do_projeto():
    o = overlay_vazio()
    o["adicionadas"] = [{**copy.deepcopy(PADRAO["regras"][0]), "id": "P229-X"}]
    o["grupos"] = {"MEUGRUPO": ["ABC*"], PADRAO["grupos"] and next(iter(PADRAO["grupos"])): ["ZZZ"]}
    regras, grupos, origem_g, _ = efetivo(PADRAO, o)
    assert por_id(regras)["P229-X"]["origem"] == "projeto"
    assert origem_g["MEUGRUPO"] == "projeto" and "sobrescrita" in origem_g.values()


def test_orfas_e_id_repetido_viram_aviso():
    o = overlay_vazio()
    o["sobrescritas"]["NAO-EXISTE"] = {"ativa": True}
    o["ocultas"] = ["OUTRA"]
    o["adicionadas"] = [{**copy.deepcopy(PADRAO["regras"][0]), "id": "C2-CFU-SUPL"}]
    regras, _, _, avisos = efetivo(PADRAO, o)
    assert len(avisos) == 3 and len(regras) == len(PADRAO["regras"])


def test_diff_ida_e_volta_e_identica():
    editado = [copy.deepcopy(r) for r in PADRAO["regras"]]
    editado[0]["severidade"] = "erro"
    editado[1]["oculta"] = True
    editado.append({**copy.deepcopy(PADRAO["regras"][2]), "id": "P-NOVA"})
    editado.pop(2)                                                   # regra do padrão ausente = oculta
    o = overlay_de_efetivo(PADRAO, editado, PADRAO["grupos"])
    assert set(o["sobrescritas"]) == {editado[0]["id"]} and o["grupos"] == {}
    assert set(o["ocultas"]) == {editado[1]["id"], PADRAO["regras"][2]["id"]}
    assert [r["id"] for r in o["adicionadas"]] == ["P-NOVA"]
    regras = efetivo(PADRAO, o)[0]
    assert sorted(r["id"] for r in regras) == sorted([r["id"] for r in PADRAO["regras"]] + ["P-NOVA"])
    assert por_id(regras)[PADRAO["regras"][2]["id"]]["oculta"] is True


def test_lista_v1_antiga_vira_overlay_equivalente():
    """Cópia integral antiga (lista) → overlay; o efetivo avalia igual à cópia."""
    copia = copy.deepcopy(SEED["regras"])
    copia[0]["ativa"] = True
    copia[0]["mensagem"] = "editada"
    copia = [r for r in copia if r["id"] != "C2-BT-EXT"]             # a cópia não tinha a regra mais nova
    o = overlay_de_bruto(PADRAO, {"versao": 2, "grupos": SEED["grupos"], "regras": copia})
    assert o["ocultas"] == ["C2-BT-EXT"] and list(o["sobrescritas"]) == [copia[0]["id"]] and o["adicionadas"] == []
    regras, grupos, _, _ = efetivo(PADRAO, o)
    amostra = [linha("DT11/300 1-CFU"), linha("SI4 1-DT"), linha("SI3")]
    assert ids(avaliar(regras, [], amostra, grupos)) == ids(avaliar(copia, [], amostra, SEED["grupos"]))


# ── rotas ────────────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(role):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


def get(client, projeto="DEFAULT"):
    return client.get(f"/api/validacao/regras?projeto_codigo={projeto}", headers=cab("admin")).json()


def salvar(client, projeto, regras, grupos=None):
    corpo = {"projeto_codigo": projeto, "regras": regras}
    if grupos is not None:
        corpo["grupos"] = grupos
    return client.post("/api/validacao/regras", json=corpo, headers=cab("admin"))


def editar(client, projeto, mut):
    regras = get(client, projeto)["regras"]
    mut(por_id(regras), regras)
    return salvar(client, projeto, regras)


def linhas_db(projeto):
    conn = database.get_connection()
    row = conn.execute("SELECT regras_json FROM regras_dominio WHERE projeto_codigo = ?", (projeto,)).fetchone()
    conn.close()
    return json.loads(row[0]) if row else None


def test_projeto_guarda_so_o_overlay_e_o_padrao_e_o_outro_projeto_ficam_intactos(client):
    padrao_antes = linhas_db("DEFAULT")

    def mut(m, regras):
        m["C2-CFU-SUPL"]["severidade"] = "erro"
        m["C2-CFU-SUPL"]["ativa"] = True
        regras.append({**copy.deepcopy(m["C2-CFU-EF"]), "id": "P229-POSTE"})
    assert editar(client, "229", mut).status_code == 200
    guardado = linhas_db("229")
    assert guardado["overlay"] is True and list(guardado["sobrescritas"]) == ["C2-CFU-SUPL"]
    assert [r["id"] for r in guardado["adicionadas"]] == ["P229-POSTE"] and guardado["ocultas"] == []
    assert linhas_db("DEFAULT") == padrao_antes
    p229, p027 = por_id(get(client, "229")["regras"]), por_id(get(client, "027")["regras"])
    assert p229["C2-CFU-SUPL"]["origem"] == "sobrescrita" and p229["P229-POSTE"]["origem"] == "projeto"
    assert "P229-POSTE" not in p027 and p027["C2-CFU-SUPL"]["severidade"] == "aviso" and not p027["C2-CFU-SUPL"]["ativa"]
    assert "P229-POSTE" not in por_id(get(client)["regras"])
    assert get(client, "229")["personalizado"] is True and get(client, "027")["personalizado"] is False


def test_mudanca_no_padrao_chega_ao_projeto_exceto_sobrescrita(client):
    editar(client, "229", lambda m, r: m["C2-CFU-SUPL"].update(mensagem="Do projeto"))
    def mut(m, regras):
        m["C2-CFU-SUPL"]["mensagem"] = "Do padrão v2"
        m["C2-CFU-EF"]["mensagem"] = "EF do padrão v2"
    editar(client, "DEFAULT", mut)
    p = por_id(get(client, "229")["regras"])
    assert p["C2-CFU-SUPL"]["mensagem"] == "Do projeto" and p["C2-CFU-EF"]["mensagem"] == "EF do padrão v2"
    assert por_id(get(client, "027")["regras"])["C2-CFU-EF"]["mensagem"] == "EF do padrão v2"


def test_ocultar_e_reexibir_por_projeto_e_validacao_respeita(client):
    editar(client, "229", lambda m, r: m["C2-CFU-SUPL"].update(ativa=True))
    corpo = {"cabos": [], "outros": [linha("DT11/300 1-CFU")], "projeto_codigo": "229"}
    post = lambda: client.post("/api/validacao/planilhas", json=corpo, headers=cab("operador")).json()["achados"]
    assert ids(post()) == ["C2-CFU-SUPL"]
    editar(client, "229", lambda m, r: m["C2-CFU-SUPL"].update(oculta=True))         # botão −
    r = por_id(get(client, "229")["regras"])["C2-CFU-SUPL"]
    assert r["oculta"] is True and r["ativa"] is True and post() == []              # continua listada, não dispara
    editar(client, "229", lambda m, r: m["C2-CFU-SUPL"].update(oculta=False))        # botão +
    assert ids(post()) == ["C2-CFU-SUPL"]


def test_voltar_ao_padrao_desfaz_a_sobrescrita_de_uma_regra(client):
    editar(client, "229", lambda m, r: m["C2-CFU-SUPL"].update(severidade="erro"))
    padrao = por_id(get(client)["regras"])["C2-CFU-SUPL"]
    editar(client, "229", lambda m, r: m["C2-CFU-SUPL"].update(severidade=padrao["severidade"]))
    assert linhas_db("229")["sobrescritas"] == {}
    assert por_id(get(client, "229")["regras"])["C2-CFU-SUPL"]["origem"] == "padrao"


def test_grupo_do_projeto_redefine_so_ali(client):
    grupos = get(client)["grupos"]
    nome = next(iter(grupos))
    r = salvar(client, "229", get(client, "229")["regras"], {**grupos, nome: ["ZZZ*"], "NOVO": ["AB"]})
    assert r.status_code == 200
    g = get(client, "229")
    assert g["grupos"][nome] == ["ZZZ*"] and g["grupos_origem"][nome] == "sobrescrita" and g["grupos_origem"]["NOVO"] == "projeto"
    assert get(client, "027")["grupos"][nome] == grupos[nome] and "NOVO" not in get(client, "027")["grupos"]


def test_regra_invalida_no_projeto_nao_salva(client):
    def mut(m, regras):
        regras.append({**copy.deepcopy(m["C2-CFU-EF"]), "id": "P-RUIM", "quando": {"tem": {"regex": "(["}}})
    assert editar(client, "229", mut).status_code == 400
    assert linhas_db("229") is None


def test_historico_e_reversao_do_overlay(client):
    editar(client, "229", lambda m, r: m["C2-CFU-SUPL"].update(ativa=True))
    editar(client, "229", lambda m, r: m["C2-CFU-EF"].update(ativa=True))
    hist = client.get("/api/validacao/regras/historico?projeto_codigo=229", headers=cab("admin")).json()["historico"]
    assert len(hist) == 1 and por_id(hist[0]["regras"])["C2-CFU-SUPL"]["ativa"] and not por_id(hist[0]["regras"])["C2-CFU-EF"]["ativa"]
    assert client.post("/api/validacao/regras/reverter", json={"projeto_codigo": "229", "historico_id": hist[0]["id"]},
                       headers=cab("admin")).status_code == 200
    p = por_id(get(client, "229")["regras"])
    assert p["C2-CFU-SUPL"]["ativa"] and not p["C2-CFU-EF"]["ativa"]


def test_restaurar_semente_no_projeto_descarta_ajustes_e_guarda_historico(client):
    editar(client, "229", lambda m, r: m["C2-CFU-SUPL"].update(ativa=True))
    assert client.post("/api/validacao/regras/restaurar-semente", json={"projeto_codigo": "229"}, headers=cab("admin")).status_code == 200
    g = get(client, "229")
    assert g["personalizado"] is False and not any(r["ativa"] for r in g["regras"])
    assert len(client.get("/api/validacao/regras/historico?projeto_codigo=229", headers=cab("admin")).json()["historico"]) == 1


def test_adicionar_novas_em_projeto_nao_faz_nada(client):
    r = client.post("/api/validacao/regras/adicionar-novas", json={"projeto_codigo": "229"}, headers=cab("admin")).json()
    assert r["adicionadas"] == [] and "automaticamente" in r["mensagem"] and linhas_db("229") is None


def test_migracao_de_copia_integral_antiga_nao_muda_o_resultado(client):
    """Projeto que já tinha a lista inteira salva (formato antigo) segue validando igual após a leitura em camadas."""
    copia = copy.deepcopy(SEED["regras"])
    for r in copia:
        r["ativa"] = r["id"] in ("C2-CFU-SUPL", "C2-CFU-EF")
    copia = [r for r in copia if r["id"] != "C2-BT-EXT"]
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO regras_dominio (projeto_codigo, regras_json) VALUES (?, ?)",
                 ("229", json.dumps(copia, ensure_ascii=False)))
    conn.commit()
    conn.close()
    amostra = [linha("DT11/300 1-CFU"), linha("SI4 1-DT"), linha("CH 1-CFUR")]
    esperado = avaliar(copia, [], amostra, SEED["grupos"])
    corpo = {"cabos": [], "outros": amostra, "projeto_codigo": "229"}
    real = client.post("/api/validacao/planilhas", json=corpo, headers=cab("operador")).json()["achados"]
    assert [i for i in ids(real) if i.startswith("C2-")] == ids(esperado)
    assert get(client, "229")["personalizado"] is True
    assert por_id(get(client, "229")["regras"])["C2-BT-EXT"]["oculta"] is True
