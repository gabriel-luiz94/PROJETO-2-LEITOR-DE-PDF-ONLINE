"""tests/test_ajustes_cadastro.py — cadastro de ajustes (receitas) em camadas por projeto (TASK-023, ADR-006).

Decisões do usuário: salva no Supabase (tabelas próprias) com botão Salvar; ajustes sempre com pré-visualização;
mesmo modelo de camadas das regras (padrão + ajustes do projeto, −/+ oculta)."""
import copy
import json

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from config import AJUSTES_SEED_PATH
from middleware.auth_middleware import create_jwt_token
from services.ajustes_camadas import efetivo, lista_de_bruto, overlay_de_bruto, overlay_de_efetivo, overlay_vazio
from services.ajustes_planilhas import validar_receitas

with open(AJUSTES_SEED_PATH, encoding="utf-8") as f:
    SEED = json.load(f)
PADRAO = SEED["ajustes"]


def por_id(lista):
    return {r["id"]: r for r in lista}


def receita(rid="AJ-X", **kw):
    base = {"id": rid, "nome": "Meu ajuste", "ativa": True, "regras": [],
            "acoes": [{"acao": "normalizar", "tabela": "outros", "regras": ["espacos"]}]}
    base.update(kw)
    return base


# ── semente ──────────────────────────────────────────────────────────────────
def test_semente_e_valida_e_vem_toda_desligada():
    assert validar_receitas(PADRAO, {"CHAVE_MT": ["CFU"]}) == []
    assert len(PADRAO) == 5 and not any(r["ativa"] for r in PADRAO)
    assert por_id(PADRAO)["AJ-CFU-SUPL"]["regras"] == ["C2-CFU-SUPL"]


# ── camadas (funções puras) ──────────────────────────────────────────────────
def test_overlay_vazio_e_o_padrao_puro_e_sobrescrita_oculta_e_propria():
    lista, avisos = efetivo(PADRAO, overlay_vazio())
    assert [r["id"] for r in lista] == [r["id"] for r in PADRAO] and avisos == []
    assert all(r["origem"] == "padrao" and r["oculta"] is False for r in lista)
    o = overlay_vazio()
    o["sobrescritas"]["AJ-SUP-L"] = {"ativa": True, "nome": "Meu nome"}
    o["ocultas"] = ["AJ-EXCLUIR-VAZIAS"]
    o["adicionadas"] = [receita("P229-A")]
    m = por_id(efetivo(PADRAO, o)[0])
    assert m["AJ-SUP-L"]["origem"] == "sobrescrita" and m["AJ-SUP-L"]["ativa"] is True and m["AJ-SUP-L"]["nome"] == "Meu nome"
    assert m["AJ-EXCLUIR-VAZIAS"]["oculta"] is True and m["P229-A"]["origem"] == "projeto"


def test_none_remove_campo_padrao_evolui_e_orfas_viram_aviso():
    o = overlay_vazio()
    o["sobrescritas"]["AJ-CFU-SUPL"] = {"regras": None}
    assert "regras" not in por_id(efetivo(PADRAO, o)[0])["AJ-CFU-SUPL"]
    novo = copy.deepcopy(PADRAO)
    for r in novo:
        r["descricao"] = "NOVA " + r["id"]
    o = overlay_vazio()
    o["sobrescritas"]["AJ-SUP-L"] = {"descricao": "minha"}
    m = por_id(efetivo(novo, o)[0])
    assert m["AJ-SUP-L"]["descricao"] == "minha" and m["AJ-ORDENAR-OUTROS"]["descricao"] == "NOVA AJ-ORDENAR-OUTROS"
    o["sobrescritas"]["SUMIU"] = {"ativa": True}
    o["ocultas"] = ["OUTRA"]
    o["adicionadas"] = [receita("AJ-SUP-L")]
    lista, avisos = efetivo(PADRAO, o)
    assert len(avisos) == 3 and len(lista) == len(PADRAO)


def test_diff_ida_e_volta_e_copia_integral_antiga_vira_overlay_equivalente():
    editado = copy.deepcopy(PADRAO)
    editado[0]["ativa"] = True
    editado[1]["oculta"] = True
    editado.append(receita("P-NOVA"))
    editado.pop(2)                                               # ausente do padrão = oculta
    o = overlay_de_efetivo(PADRAO, editado)
    assert list(o["sobrescritas"]) == [PADRAO[0]["id"]] and set(o["ocultas"]) == {PADRAO[1]["id"], PADRAO[2]["id"]}
    assert [r["id"] for r in o["adicionadas"]] == ["P-NOVA"]
    m = por_id(efetivo(PADRAO, o)[0])
    assert m[PADRAO[0]["id"]]["ativa"] is True and m[PADRAO[1]["id"]]["oculta"] and m["P-NOVA"]["origem"] == "projeto"
    copia = copy.deepcopy(PADRAO)
    copia[3]["ativa"] = True
    o2 = overlay_de_bruto(PADRAO, {"versao": 1, "ajustes": copia[:4]})        # a cópia não tinha o último
    assert o2["ocultas"] == [PADRAO[4]["id"]] and list(o2["sobrescritas"]) == [PADRAO[3]["id"]]
    assert lista_de_bruto(None) == [] and overlay_de_bruto(PADRAO, "x") == overlay_vazio()


# ── schema do cadastro ───────────────────────────────────────────────────────
@pytest.mark.parametrize("mut,trecho", [
    (lambda r: r.pop("id"), "'id' obrigatório"),
    (lambda r: r.update(id="com espaço"), "'id' obrigatório"),
    (lambda r: r.pop("nome"), "'nome' obrigatório"),
    (lambda r: r.update(nome="x" * 121), "'nome' obrigatório"),
    (lambda r: r.update(ativa="sim"), "'ativa'"),
    (lambda r: r.update(descricao=1), "'descricao'"),
    (lambda r: r.update(regras="C2"), "'regras'"),
    (lambda r: r.update(regras=[""]), "'regras'"),
    (lambda r: r.update(acoes=[]), "'acoes'"),
    (lambda r: r.update(acoes=[{"acao": "x"}]), "'acao' precisa ser"),
    (lambda r: r.update(foo=1), "campo desconhecido"),
])
def test_schema_do_cadastro_rejeita(mut, trecho):
    r = receita()
    mut(r)
    assert any(trecho in e for e in validar_receitas([r])), validar_receitas([r])


def test_schema_do_cadastro_ids_duplicados_e_nao_lista():
    assert any("duplicado" in e for e in validar_receitas([receita("A"), receita("A")]))
    assert validar_receitas({"a": 1}) == ["O payload de ajustes precisa ser uma lista."]
    assert validar_receitas(["x"]) == ["Ajuste #1: precisa ser um objeto."]


# ── rotas ────────────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(role):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


def get(client, projeto="DEFAULT", role="admin"):
    return client.get(f"/api/validacao/ajustes?projeto_codigo={projeto}", headers=cab(role)).json()


def salvar(client, projeto, ajustes):
    return client.post("/api/validacao/ajustes", json={"projeto_codigo": projeto, "ajustes": ajustes}, headers=cab("admin"))


def editar(client, projeto, mut):
    lista = get(client, projeto)["ajustes"]
    mut(por_id(lista), lista)
    return salvar(client, projeto, lista)


def linhas_db(projeto):
    conn = database.get_connection()
    row = conn.execute("SELECT ajustes_json FROM ajustes_planilhas WHERE projeto_codigo = ?", (projeto,)).fetchone()
    conn.close()
    return json.loads(row[0]) if row else None


def test_get_devolve_semente_com_frases_e_operador_le_mas_nao_escreve(client):
    g = get(client, role="operador")
    assert len(g["ajustes"]) == 5 and g["versao_de"] == "DEFAULT" and g["personalizado"] is False
    assert por_id(g["ajustes"])["AJ-SUP-L"]["frases"] == ["Em Outros: substituir 'SUP-L' por 'SUPL'."]
    for metodo, url, corpo in [("post", "/api/validacao/ajustes", {"ajustes": []}), ("get", "/api/validacao/ajustes/historico", None),
                               ("post", "/api/validacao/ajustes/reverter", {"historico_id": 1}),
                               ("post", "/api/validacao/ajustes/restaurar-semente", {})]:
        r = getattr(client, metodo)(url, headers=cab("operador"), **({"json": corpo} if corpo is not None else {}))
        assert r.status_code == 403, url
    assert client.get("/api/validacao/ajustes").status_code == 401


def test_projeto_guarda_so_o_overlay_e_o_padrao_e_o_outro_projeto_ficam_intactos(client):
    padrao_antes = linhas_db("DEFAULT")

    def mut(m, lista):
        m["AJ-SUP-L"]["ativa"] = True
        m["AJ-EXCLUIR-VAZIAS"]["oculta"] = True
        lista.append(receita("P229-X"))
    assert editar(client, "229", mut).status_code == 200
    guardado = linhas_db("229")
    assert guardado["overlay"] is True and list(guardado["sobrescritas"]) == ["AJ-SUP-L"] and guardado["ocultas"] == ["AJ-EXCLUIR-VAZIAS"]
    assert [r["id"] for r in guardado["adicionadas"]] == ["P229-X"] and linhas_db("DEFAULT") == padrao_antes
    p229, p027 = por_id(get(client, "229")["ajustes"]), por_id(get(client, "027")["ajustes"])
    assert p229["AJ-SUP-L"]["origem"] == "sobrescrita" and p229["AJ-EXCLUIR-VAZIAS"]["oculta"] and p229["P229-X"]["origem"] == "projeto"
    assert "P229-X" not in p027 and not p027["AJ-SUP-L"]["ativa"] and "P229-X" not in por_id(get(client)["ajustes"])
    assert get(client, "229")["personalizado"] is True and get(client, "027")["personalizado"] is False


def test_padrao_muda_e_projeto_acompanha_exceto_sobrescrita(client):
    editar(client, "229", lambda m, l: m["AJ-SUP-L"].update(nome="Do projeto"))
    editar(client, "DEFAULT", lambda m, l: (m["AJ-SUP-L"].update(nome="Padrão v2"), m["AJ-CFU-SUPL"].update(nome="CFU v2")))
    p = por_id(get(client, "229")["ajustes"])
    assert p["AJ-SUP-L"]["nome"] == "Do projeto" and p["AJ-CFU-SUPL"]["nome"] == "CFU v2"


def test_ajuste_invalido_nao_salva_e_nada_e_gravado(client):
    def mut(m, lista):
        lista.append(receita("P-RUIM", acoes=[{"acao": "substituir", "tabela": "outros", "de": {"regex": "(["}, "para": "x"}]))
    r = editar(client, "229", mut)
    assert r.status_code == 400 and any("regex inválida" in e for e in r.json()["detail"]["erros"]) and linhas_db("229") is None
    assert salvar(client, "DEFAULT", [receita("A"), receita("A")]).status_code == 400


def test_historico_reversao_e_restaurar_semente(client):
    editar(client, "229", lambda m, l: m["AJ-SUP-L"].update(ativa=True))
    editar(client, "229", lambda m, l: m["AJ-ORDENAR-OUTROS"].update(ativa=True))
    hist = client.get("/api/validacao/ajustes/historico?projeto_codigo=229", headers=cab("admin")).json()["historico"]
    assert len(hist) == 1 and por_id(hist[0]["ajustes"])["AJ-SUP-L"]["ativa"] and not por_id(hist[0]["ajustes"])["AJ-ORDENAR-OUTROS"]["ativa"]
    assert client.post("/api/validacao/ajustes/reverter", json={"projeto_codigo": "229", "historico_id": hist[0]["id"]}, headers=cab("admin")).status_code == 200
    p = por_id(get(client, "229")["ajustes"])
    assert p["AJ-SUP-L"]["ativa"] and not p["AJ-ORDENAR-OUTROS"]["ativa"]
    assert client.post("/api/validacao/ajustes/reverter", json={"projeto_codigo": "229", "historico_id": 999}, headers=cab("admin")).status_code == 404
    assert client.post("/api/validacao/ajustes/restaurar-semente", json={"projeto_codigo": "229"}, headers=cab("admin")).status_code == 200
    g = get(client, "229")
    assert g["personalizado"] is False and not any(r["ativa"] for r in g["ajustes"])
    # DEFAULT: editar e restaurar a semente
    editar(client, "DEFAULT", lambda m, l: m["AJ-SUP-L"].update(ativa=True))
    assert client.post("/api/validacao/ajustes/restaurar-semente", json={"projeto_codigo": "DEFAULT"}, headers=cab("admin")).status_code == 200
    assert not any(r["ativa"] for r in get(client)["ajustes"])


def test_por_regra_so_lista_ajustes_ligados_e_nao_ocultos(client):
    assert get(client)["por_regra"] == {}
    editar(client, "229", lambda m, l: m["AJ-CFU-SUPL"].update(ativa=True))
    assert get(client, "229")["por_regra"] == {"C2-CFU-SUPL": ["AJ-CFU-SUPL"]}
    editar(client, "229", lambda m, l: m["AJ-CFU-SUPL"].update(oculta=True))
    assert get(client, "229")["por_regra"] == {}


def test_preview_por_receita_inclui_inativa_explicita_e_barra_oculta_e_inexistente(client):
    corpo = {"receitas": ["AJ-SUP-L"], "outros": [{"operacao": "I", "ativo": "DT11/300 1-SUP-L"}], "projeto_codigo": "229"}
    r = client.post("/api/validacao/ajustes/preview", json=corpo, headers=cab("operador")).json()
    assert r["outros"][0]["ativo"] == "DT11/300 1-SUPL" and r["resumo"]["editar"] == 1
    r = client.post("/api/validacao/ajustes/preview", headers=cab("operador"), json={
        "receitas": ["AJ-SUP-L"], "acoes": [{"acao": "mesclar_duplicadas", "tabela": "outros"}],
        "outros": [{"operacao": "I", "ativo": "1-SUP-L 1-SUPL"}]}).json()
    assert r["outros"][0]["ativo"] == "2-SUPL"                     # receitas primeiro, depois as ações avulsas
    assert client.post("/api/validacao/ajustes/preview", json={**corpo, "receitas": ["NAO-EXISTE"]}, headers=cab("operador")).status_code == 404
    editar(client, "229", lambda m, l: m["AJ-SUP-L"].update(oculta=True))
    r = client.post("/api/validacao/ajustes/preview", json=corpo, headers=cab("operador"))
    assert r.status_code == 400 and "oculto" in r.json()["detail"]


def test_receita_usa_os_grupos_de_ativos_das_regras(client):
    editar(client, "DEFAULT", lambda m, l: l.append(receita("AJ-G", acoes=[{"acao": "remover_ativo", "tabela": "outros", "ativo": "@CHAVE_MT"}])))
    r = client.post("/api/validacao/ajustes/preview", headers=cab("operador"), json={
        "receitas": ["AJ-G"], "outros": [{"operacao": "I", "ativo": "DT11/300 1-CFU 1-X"}]}).json()
    assert r["outros"][0]["ativo"] == "DT11/300 1-X"
    ruim = receita("AJ-H", acoes=[{"acao": "remover_ativo", "tabela": "outros", "ativo": "@NAO_EXISTE"}])
    assert salvar(client, "DEFAULT", [ruim]).status_code == 400


# ── Supabase (simulado) ──────────────────────────────────────────────────────
class FakeTabela:
    def __init__(self, nuvem, nome, falhar):
        self.nuvem, self.nome, self.falhar, self._f = nuvem, nome, falhar, {}

    def select(self, *_):
        return self

    def eq(self, c, v):
        self._f[c] = v
        return self

    def order(self, *_a, **_k):
        return self

    def insert(self, d):
        self.nuvem.setdefault(self.nome + "_ins", []).append(d)
        return self

    def upsert(self, d):
        if self.falhar:
            raise RuntimeError("nuvem fora")
        self.nuvem[(self.nome, d["projeto_codigo"])] = d
        return self

    def execute(self):
        class R: data = []
        r = R()
        chave = (self.nome, self._f.get("projeto_codigo"))
        r.data = [self.nuvem[chave]] if chave in self.nuvem else []
        return r


class FakeSupabase:
    def __init__(self, falhar=False):
        self.nuvem, self.falhar = {}, falhar

    def table(self, nome):
        return FakeTabela(self.nuvem, nome, self.falhar)


def test_salva_na_nuvem_e_le_de_la_primeiro(client, monkeypatch):
    import services.repo_json as rj
    fake = FakeSupabase()
    monkeypatch.setattr(rj, "get_supabase", lambda: fake)
    assert editar(client, "229", lambda m, l: m["AJ-SUP-L"].update(ativa=True)).status_code == 200
    nuvem = json.loads(fake.nuvem[("ajustes_planilhas", "229")]["ajustes_json"])
    assert nuvem["overlay"] is True and nuvem["sobrescritas"] == {"AJ-SUP-L": {"ativa": True}}
    assert linhas_db("229") == nuvem                                  # local e nuvem iguais
    fake.nuvem[("ajustes_planilhas", "229")]["ajustes_json"] = json.dumps({**nuvem, "sobrescritas": {"AJ-SUP-L": {"ativa": False, "nome": "da nuvem"}}})
    assert por_id(get(client, "229")["ajustes"])["AJ-SUP-L"]["nome"] == "da nuvem"   # nuvem primeiro


def test_falha_de_sincronizacao_avisa_mas_o_local_ficou_salvo(client, monkeypatch):
    import services.repo_json as rj
    monkeypatch.setattr(rj, "get_supabase", lambda: FakeSupabase(falhar=True))
    r = editar(client, "229", lambda m, l: m["AJ-SUP-L"].update(ativa=True))
    assert r.status_code == 500 and "Supabase" in r.json()["detail"] and "nuvem fora" in r.json()["detail"]
    assert linhas_db("229")["sobrescritas"] == {"AJ-SUP-L": {"ativa": True}}


def test_barra_do_resumo_so_tem_os_botoes_de_validacao_e_as_opcoes_vivem_na_gaveta():
    """TASK-026: caixas de modo e 'Validar ao montar' saem da barra; ficam na aba Execução da gaveta."""
    resumo = open("static/resumo.html", encoding="utf-8").read()
    for botao in ("btn-regras-validacao", "btn-ajustar", "btn-validar", "btn-validacao-opcoes", "validar-estado"):
        assert f'id="{botao}"' in resumo, botao
    for velho in ('id="vmodo-', 'name="validacao_auto"', 'id="validacao-modos"'):
        assert velho not in resumo, velho
    gaveta = open("static/painel_regras.js", encoding="utf-8").read()
    for marca in ("montarPainelExecucao", "vmodo-det", "vmodo-pular-ia", "validacao_auto", "rpAbrirNaAba"):
        assert marca in gaveta, marca
    assert "window.validacaoPrefs" in open("static/resumo.js", encoding="utf-8").read()
