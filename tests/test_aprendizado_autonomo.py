"""TASK-059 — aprendizado supervisionado: comparação (só diferenças), padrões, propostas, aprovação, IA opcional, privacidade."""
import json
import uuid

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.aprendizado as rota_apr
import routers.obras as rota_obras
from middleware.auth_middleware import create_jwt_token
from services.ajustes_planilhas import ajustar
from services.autonomo import aprendizado as ap
from services.autonomo import aprendizado_repo as repo
from services.autonomo import execucoes


def L(a, op="I"):
    return {"entidade": "X", "operacao": op, "ativo": a}


# ── comparar ────────────────────────────────────────────────────────────────
def tipos(r):
    return sorted(e["tipo"] for e in r["eventos"])


def test_sem_diferenca_nao_gera_evento_e_conta_a_base():
    t = {"cabos": [L("CAA 2 ABC 35 m")], "outros": [L("2-U4"), L("1-DT11/300 1-U3")]}
    r = ap.comparar(t, json.loads(json.dumps(t)))
    assert r["eventos"] == [] and r["base"]["item"] == {"U4": 1, "DT11/300": 1, "U3": 1} and r["base"]["linha"]["outros|2-U4"] == 1
    assert r["linhas"] == {"agente": 3, "usuario": 3}


def test_detecta_cada_tipo_de_correcao():
    ag = {"cabos": [L("CA 2 ABC 35 m"), L("CAA 2 ABC 10 m")], "outros": [L("2-U4"), L("1-DT11/300 1-U3"), L("LIXO"), L("1-P1", "I")]}
    us = {"cabos": [L("CAA 2 ABC 35 m"), L("CAA 2 ABC 10 m")], "outros": [L("2-U3"), L("1-DT11/300 1-U3 1-X9"), L("1-P1", "R")]}
    r = ap.comparar(ag, us)
    por = {e["tipo"]: e for e in r["eventos"]}
    assert por["texto_trocado"] == {"tipo": "texto_trocado", "tabela": "cabos", "de": "CA", "para": "CAA"}
    assert por["item_trocado"]["de"] == "U4" and por["item_trocado"]["para"] == "U3"
    assert {(e["ativo"], e["contexto"]) for e in r["eventos"] if e["tipo"] == "item_adicionado"} == {("X9", "DT11/300"), ("X9", "U3")}
    assert por["linha_excluida"]["ativo"] == "LIXO"
    assert por["operacao_alterada"]["de"] == "I" and por["operacao_alterada"]["para"] == "R"


def test_linha_que_so_mudou_de_lugar_nao_e_correcao():
    ag = {"cabos": [], "outros": [L("1-A"), L("1-B"), L("1-C")]}
    us = {"cabos": [], "outros": [L("1-C"), L("1-A"), L("1-B")]}
    assert ap.comparar(ag, us)["eventos"] == []


def test_quantidade_alterada_e_linha_adicionada():
    r = ap.comparar({"cabos": [], "outros": [L("2-U4")]}, {"cabos": [], "outros": [L("3-U4"), L("1-NOVO")]})
    assert tipos(r) == ["linha_adicionada", "qtd_alterada"]


def test_eventos_nao_levam_coordenadas_nem_textos_do_desenho():
    ag = {"cabos": [], "outros": [{"entidade": "X", "operacao": "I", "ativo": "2-U4", "_x": 1.5, "_y": 2.5, "texto": "SEGREDO"}]}
    us = {"cabos": [], "outros": [{"entidade": "X", "operacao": "I", "ativo": "2-U3", "_x": 1.5, "_y": 2.5, "texto": "SEGREDO"}]}
    s = json.dumps(ap.comparar(ag, us))
    assert "SEGREDO" not in s and "1.5" not in s


def test_tabela_enorme_e_ignorada_sem_travar():
    grande = [L(f"1-X{i}") for i in range(ap.LIMITE_LINHAS_COMPARAR + 1)]
    r = ap.comparar({"cabos": [], "outros": grande}, {"cabos": [], "outros": grande[:-1]})
    assert r["eventos"] == [] and r["ignoradas"]


# ── padrões ─────────────────────────────────────────────────────────────────
def registros(n, antes, depois):
    out = []
    for i in range(n):
        r = ap.comparar(antes, depois)
        out.append({"execucao_id": f"e{i}", "eventos": r["eventos"], "base": r["base"]})
    return out


def test_so_vira_proposta_com_repeticao_obras_e_consistencia():
    ag = {"cabos": [], "outros": [L("2-U4")]}
    us = {"cabos": [], "outros": [L("2-U3")]}
    assert ap.detectar_padroes(registros(2, ag, us)) == []                                   # poucas vezes
    um = registros(1, {"cabos": [], "outros": [L("2-U4"), L("3-U4"), L("4-U4")]}, {"cabos": [], "outros": [L("2-U3"), L("3-U3"), L("4-U3")]})
    assert ap.detectar_padroes(um) == []                                                      # 3 vezes mas em UMA obra só
    p = ap.detectar_padroes(registros(3, ag, us))
    assert len(p) == 1 and p[0]["tipo"] == "item_trocado" and p[0]["consistencia"] == 1.0 and p[0]["receita"]["acoes"][0]["modo"] == "item"
    # o mesmo ativo aparece muito SEM ser trocado -> consistência baixa -> não propõe
    ruido = registros(3, ag, us) + [{"execucao_id": f"ok{i}", "eventos": [], "base": {"item": {"U4": 3}}} for i in range(4)]
    assert ap.detectar_padroes(ruido) == []


def test_receitas_propostas_reproduzem_a_correcao_no_motor():
    ag = {"cabos": [L("CA 2 ABC 35 m")], "outros": [L("2-U4"), L("1-DT11/300 1-U3"), L("LIXO"), L("1-U5 1-Z1")]}
    us = {"cabos": [L("CAA 2 ABC 35 m")], "outros": [L("2-U3"), L("1-DT11/300 1-U3"), L("1-U5")]}
    p = ap.detectar_padroes(registros(3, ag, us))
    assert {x["tipo"] for x in p} == {"texto_trocado", "item_trocado", "linha_excluida", "item_removido"}
    acoes = [x["receita"]["acoes"][0] for x in p if x["receita"]]
    assert len(acoes) == 4
    diff = ajustar(acoes, ag["cabos"], ag["outros"], {})
    ativos = lambda t: [l["ativo"] for l in t]
    assert ativos(diff["cabos"]) == ["CAA 2 ABC 35 m"] and ativos(diff["outros"]) == ["2-U3", "1-DT11/300 1-U3", "1-U5"]      # igual ao do usuário


def test_item_adicionado_vira_ajuste_com_condicao():
    ag = {"cabos": [], "outros": [L("2-U3")]}
    us = {"cabos": [], "outros": [L("2-U3 1-90277")]}
    p = ap.detectar_padroes(registros(3, ag, us))
    a = p[0]["receita"]["acoes"][0]
    assert a["acao"] == "adicionar_ativo" and a["ativo"] == "90277" and a["quando"] == {"tem": "U3"} and a["qtd"] == 1
    assert [l["ativo"] for l in ajustar([a], ag["cabos"], ag["outros"], {})["outros"]] == ["2-U3 1-90277"]


def test_operacao_alterada_e_informativa_sem_receita():
    ag = {"cabos": [], "outros": [L("1-P1", "I")]}
    us = {"cabos": [], "outros": [L("1-P1", "R")]}
    p = ap.detectar_padroes(registros(3, ag, us))
    assert p[0]["tipo"] == "operacao_alterada" and p[0]["receita"] is None


def test_ativo_com_caractere_especial_na_exclusao():
    ag = {"cabos": [], "outros": [L("1-DT11/300 (X)")]}
    us = {"cabos": [], "outros": []}
    p = ap.detectar_padroes(registros(3, ag, us))
    assert p[0]["receita"] and ajustar([p[0]["receita"]["acoes"][0]], [], ag["outros"], {})["outros"] == []


# ── API / banco ─────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(uid="adm", role="admin"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@x.com', role)}"}


def criar_execucao(user="adm", projeto="229", cabos=None, outros=None, status="ok"):
    eid = execucoes.criar("arq.dxf", uuid.uuid4().hex, projeto, user)
    oid = f"auto_{eid}"
    snap = {"cabos": {"bodyId": "body-cabos", "data": cabos or []}, "outros": {"bodyId": "body-outros", "data": outros or []},
            "autonomo": {"execucao_id": eid, "arquivo": "arq.dxf"}}
    rota_obras.gravar_obra({"user_id": user}, {"id": oid, "nome": "[Autônomo] arq", "data": "d", "dados_json": json.dumps(snap), "projeto": projeto, "publica": False})
    execucoes.atualizar(eid, status=status, obra_id=oid)
    return eid


def salvar_corrigida(client, eid, cabos, outros, user="adm", role="admin", projeto="229", oid=None):
    dados = {"cabos": {"bodyId": "body-cabos", "data": cabos}, "outros": {"bodyId": "body-outros", "data": outros}}
    r = client.post("/api/obras", json={"id": oid or f"obra_{uuid.uuid4().hex[:8]}", "nome": "Corrigida", "data": "d", "dados_json": json.dumps(dados),
                                        "projeto": projeto, "origem_execucao": eid}, headers=cab(user, role))
    assert r.status_code == 200, r.text
    return r


def tres_obras_com_a_mesma_correcao(client, projeto="229"):
    for _ in range(3):
        eid = criar_execucao(projeto=projeto, cabos=[L("CA 2 ABC 35 m")], outros=[L("2-U4"), L("LIXO")])
        salvar_corrigida(client, eid, [L("CAA 2 ABC 35 m")], [L("2-U3")], projeto=projeto)


def test_salvar_obra_do_autonomo_registra_so_as_diferencas(client):
    eid = criar_execucao(cabos=[L("CA 2 ABC 35 m")], outros=[L("2-U4")])
    salvar_corrigida(client, eid, [L("CAA 2 ABC 35 m")], [L("2-U3")])
    regs = repo.listar_registros("adm", "229")
    assert len(regs) == 1 and {e["tipo"] for e in regs[0]["eventos"]} == {"texto_trocado", "item_trocado"} and regs[0]["obra_salva"].startswith("obra_")
    r = client.get("/api/aprendizado/resumo?projeto=229", headers=cab()).json()
    assert r["obras_corrigidas"] == 1 and r["correcoes_total"] == 2 and r["ia_disponivel"] is False


def test_salvar_de_novo_substitui_e_obra_sem_correcao_conta_como_acerto(client):
    eid = criar_execucao(outros=[L("2-U4")])
    salvar_corrigida(client, eid, [], [L("2-U3")])
    salvar_corrigida(client, eid, [], [L("2-U4")])                       # voltou ao que o autônomo entregou
    regs = repo.listar_registros("adm", "229")
    assert len(regs) == 1 and regs[0]["eventos"] == []
    assert client.get("/api/aprendizado/resumo?projeto=229", headers=cab()).json()["obras_sem_correcao"] == 1


@pytest.mark.parametrize("quem,status", [("outro", "ok"), ("adm", "revertida"), ("adm", "erro")])
def test_execucao_alheia_ou_nao_concluida_nao_registra(client, quem, status):
    eid = criar_execucao(user=quem, outros=[L("2-U4")], status=status)
    salvar_corrigida(client, eid, [], [L("2-U3")])
    assert repo.listar_registros("adm", "229") == [] and repo.listar_registros("outro", "229") == []


def test_origem_inexistente_ou_modelo_nao_atrapalha_o_salvar(client):
    r = client.post("/api/obras", json={"id": "x1", "nome": "n", "data": "d", "dados_json": "{}", "projeto": "229", "origem_execucao": "nao_existe"}, headers=cab())
    assert r.status_code == 200
    assert repo.listar_registros("adm", "229") == []


def test_falha_no_aprendizado_nunca_derruba_o_salvar(client, monkeypatch):
    eid = criar_execucao(outros=[L("2-U4")])
    monkeypatch.setattr(repo, "registrar_correcao", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("banco")))
    salvar_corrigida(client, eid, [], [L("2-U3")])


def test_propostas_aparecem_e_aprovar_cria_o_ajuste_no_projeto(client):
    tres_obras_com_a_mesma_correcao(client)
    props = client.get("/api/aprendizado/propostas?projeto=229", headers=cab()).json()
    tipos_ = {p["tipo"] for p in props}
    assert {"texto_trocado", "item_trocado", "linha_excluida"} <= tipos_ and all(p["status"] == "pendente" for p in props)
    assert all("user_id" not in p for p in props)
    p = next(p for p in props if p["tipo"] == "item_trocado")
    r = client.post(f"/api/aprendizado/propostas/{p['id']}/aprovar", headers=cab())
    assert r.status_code == 200
    ajustes = client.get("/api/validacao/ajustes?projeto_codigo=229", headers=cab()).json()["ajustes"]
    criado = next(a for a in ajustes if a["id"] == r.json()["ajuste_id"])
    assert criado["ativa"] is True and criado["acoes"][0]["de"] == "U4" and criado["acoes"][0]["para"] == "U3"
    assert client.post(f"/api/aprendizado/propostas/{p['id']}/aprovar", headers=cab()).status_code == 400           # já decidida
    # reanalisar não recria nem reabre as decididas
    client.post("/api/aprendizado/analisar?projeto=229", headers=cab())
    assert sum(1 for x in client.get("/api/aprendizado/propostas?projeto=229", headers=cab()).json() if x["tipo"] == "item_trocado") == 1


def test_recusada_nao_volta_e_informativa_nao_aprova(client):
    for _ in range(3):
        eid = criar_execucao(outros=[L("1-P1", "I")])
        salvar_corrigida(client, eid, [], [L("1-P1", "R")])
    props = client.get("/api/aprendizado/propostas?projeto=229", headers=cab()).json()
    assert len(props) == 1 and props[0]["pode_aprovar"] is False
    assert client.post(f"/api/aprendizado/propostas/{props[0]['id']}/aprovar", headers=cab()).status_code == 400
    assert client.post(f"/api/aprendizado/propostas/{props[0]['id']}/recusar", headers=cab()).status_code == 200
    eid = criar_execucao(outros=[L("1-P1", "I")]); salvar_corrigida(client, eid, [], [L("1-P1", "R")])
    again = client.get("/api/aprendizado/propostas?projeto=229", headers=cab()).json()
    assert len(again) == 1 and again[0]["status"] == "recusada"


def test_so_administrador_e_so_o_proprio_usuario(client):
    tres_obras_com_a_mesma_correcao(client)
    assert client.get("/api/aprendizado/resumo?projeto=229", headers=cab("op", "operador")).status_code == 403
    assert client.get("/api/aprendizado/resumo?projeto=229").status_code == 401
    # outro administrador não vê nem aprova as propostas do primeiro ("só com o meu trabalho")
    assert client.get("/api/aprendizado/propostas?projeto=229", headers=cab("adm2")).json() == []
    pid = client.get("/api/aprendizado/propostas?projeto=229", headers=cab()).json()[0]["id"]
    assert client.post(f"/api/aprendizado/propostas/{pid}/aprovar", headers=cab("adm2")).status_code == 404
    assert client.get("/api/aprendizado/resumo?projeto=229", headers=cab("adm2")).json()["obras_corrigidas"] == 0


def test_aprendizado_e_separado_por_projeto_e_apagar_historico(client):
    tres_obras_com_a_mesma_correcao(client)
    assert client.get("/api/aprendizado/propostas?projeto=027", headers=cab()).json() == []
    pid = next(p["id"] for p in client.get("/api/aprendizado/propostas?projeto=229", headers=cab()).json() if p["tipo"] == "item_trocado")
    client.post(f"/api/aprendizado/propostas/{pid}/aprovar", headers=cab())
    r = client.delete("/api/aprendizado/historico?projeto=229", headers=cab()).json()
    assert r["correcoes"] == 3 and r["propostas"] >= 1
    rest = client.get("/api/aprendizado/propostas?projeto=229", headers=cab()).json()
    assert [p["status"] for p in rest] == ["aprovada"]                       # a aprovada fica (já virou ajuste)
    assert client.get("/api/aprendizado/resumo?projeto=229", headers=cab()).json()["obras_corrigidas"] == 0


def test_correcoes_listadas_sem_dados_brutos(client):
    tres_obras_com_a_mesma_correcao(client)
    c = client.get("/api/aprendizado/correcoes?projeto=229", headers=cab()).json()
    assert len(c) == 3 and c[0]["correcoes"] >= 3 and "base" not in c[0] and any("trocou" in d for d in c[0]["descricoes"])


# ── IA opcional ─────────────────────────────────────────────────────────────
def test_ia_so_depois_de_n_obras_e_resultado_vira_proposta_pendente(client, monkeypatch):
    for _ in range(ap.IA_MINIMO_OBRAS - 1):
        eid = criar_execucao(outros=[L("2-U4")]); salvar_corrigida(client, eid, [], [L("2-U3")])
    r = client.post("/api/aprendizado/ia/sugerir?projeto=229", headers=cab())
    assert r.status_code == 400 and "liberada" in r.json()["detail"]
    assert client.get("/api/aprendizado/resumo?projeto=229", headers=cab()).json()["ia_disponivel"] is False
    eid = criar_execucao(outros=[L("2-U4")]); salvar_corrigida(client, eid, [], [L("2-U3")])
    assert client.get("/api/aprendizado/resumo?projeto=229", headers=cab()).json()["ia_disponivel"] is True

    prompts = []

    async def chamar(texto):
        prompts.append(texto)
        return json.dumps({"acoes": [
            {"acao": "adicionar_ativo", "tabela": "outros", "ativo": "90277", "qtd": 1, "quando": {"tem": "U3"}, "motivo": "sempre que tem U3"},
            {"acao": "inventada", "tabela": "outros", "motivo": "x"},
            {"acao": "substituir", "tabela": "outros", "de": "A", "para": "B", "modo": "item", "operacao_nova": "R", "motivo": "muda operação"}]})
    monkeypatch.setattr(rota_apr, "_chamador", lambda request: (chamar, None))
    r = client.post("/api/aprendizado/ia/sugerir?projeto=229", headers=cab()).json()
    assert r["status"] == "ok" and r["novas"] == 1 and len(r["descartadas"]) == 2
    assert "2-U3" not in prompts[0] and "U4" in prompts[0] and "arq.dxf" not in prompts[0]        # só ativos/contagens, nada de arquivo
    ia = [p for p in client.get("/api/aprendizado/propostas?projeto=229", headers=cab()).json() if p["origem"] == "ia"]
    assert len(ia) == 1 and ia[0]["status"] == "pendente" and ia[0]["pode_aprovar"]
    assert client.get("/api/validacao/ajustes?projeto_codigo=229", headers=cab()).json()["ajustes"] and not any(a["id"] == ia[0]["receita"]["id"] for a in client.get("/api/validacao/ajustes?projeto_codigo=229", headers=cab()).json()["ajustes"])   # nada aplicado
    assert client.post("/api/aprendizado/ia/sugerir?projeto=229", headers=cab()).json()["novas"] == 0     # mesma sugestão não duplica


def test_ia_falhas_voltam_como_status(client, monkeypatch):
    for _ in range(ap.IA_MINIMO_OBRAS):
        eid = criar_execucao(outros=[L("2-U4")]); salvar_corrigida(client, eid, [], [L("2-U3")])
    monkeypatch.setattr(rota_apr, "_chamador", lambda request: (None, "Sem chave de IA"))
    assert client.post("/api/aprendizado/ia/sugerir?projeto=229", headers=cab()).json()["status"] == "sem_chave"

    async def quebra(texto):
        raise RuntimeError("provedor fora")
    monkeypatch.setattr(rota_apr, "_chamador", lambda request: (quebra, None))
    assert client.post("/api/aprendizado/ia/sugerir?projeto=229", headers=cab()).json()["status"] == "erro"

    async def lixo(texto):
        return "não é json"
    monkeypatch.setattr(rota_apr, "_chamador", lambda request: (lixo, None))
    assert client.post("/api/aprendizado/ia/sugerir?projeto=229", headers=cab()).json()["status"] == "erro"


def test_linhas_sem_parecenca_nao_viram_edicao():
    r = ap.comparar({"cabos": [], "outros": [L("LIXO")]}, {"cabos": [], "outros": [L("1-ABCDEFGH")]})
    assert tipos(r) == ["linha_adicionada", "linha_excluida"]


def test_proposta_decidida_nao_muda_de_ideia(client):
    tres_obras_com_a_mesma_correcao(client)
    props = client.get("/api/aprendizado/propostas?projeto=229", headers=cab()).json()
    a, b = props[0]["id"], props[1]["id"]
    assert client.post(f"/api/aprendizado/propostas/{a}/recusar", headers=cab()).status_code == 200
    assert client.post(f"/api/aprendizado/propostas/{a}/aprovar", headers=cab()).status_code == 400
    assert client.post(f"/api/aprendizado/propostas/{a}/recusar", headers=cab()).status_code == 400
    assert client.post(f"/api/aprendizado/propostas/{b}/aprovar", headers=cab()).status_code == 200
    assert client.post(f"/api/aprendizado/propostas/{b}/recusar", headers=cab()).status_code == 400
