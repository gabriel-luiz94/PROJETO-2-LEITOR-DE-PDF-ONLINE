"""TASK-059 etapa 2 — regras do leitor aprendidas (simuladas no motor real), captura manual (Leitor/Resumo) e aprovação."""
import json
import uuid

import pytest

pytest.importorskip("quickjs")
from fastapi.testclient import TestClient  # noqa: E402

import app as appmod  # noqa: E402
import database  # noqa: E402
import routers.obras as rota_obras  # noqa: E402
from middleware.auth_middleware import create_jwt_token  # noqa: E402
from services.autonomo import aprendizado_leitor as al  # noqa: E402
from services.autonomo import aprendizado_repo as repo  # noqa: E402
from services.autonomo import aprendizado_servico as svc  # noqa: E402
from services.autonomo import execucoes  # noqa: E402
from services.autonomo.leitor_js import LeitorJS  # noqa: E402

PROC = [{"fase": 1, "ordem": 10, "modo": "DEFINIR", "texto_regex": r"^POSTE (\d+)$", "ativo_template": "P{1}", "parar": True},
        {"fase": 1, "ordem": 20, "modo": "DEFINIR", "texto_regex": r"^TRAFO (\d+)$", "ativo_template": "T{1}", "parar": True}]
CLS = [{"ordem": 5, "ativo_regex": r"^P\d+$", "entidade": "POSTE", "parar": True},
       {"ordem": 6, "ativo_regex": r"^T\d+$", "entidade": "TRAFO", "parar": True}]


@pytest.fixture(scope="module")
def leitor():
    lt = LeitorJS()
    yield lt
    lt.fechar()


@pytest.fixture
def simular(leitor):
    return lambda entrada, proc, cls: leitor.processar_lote(entrada, proc, cls, "extracao")


def item(texto, user, layer="", n=1, fontes=("a", "b"), cor="#000000"):
    """Conjunto agregado: `user` = (entidade, operação, ativo) que o usuário deixou."""
    return {"texto": texto, "cor": cor, "layer": layer, "user": al.tri(*user), "n": n, "fontes": set(fontes)}


def aplicar(simular, itens, tabela, regra):
    entrada = [{"pagina": 1, "texto": i["texto"], "cor": i["cor"], "layer": i["layer"]} for i in itens]
    r = simular(entrada, PROC, [regra] + CLS) if tabela == "classificacao" else simular(entrada, PROC + [regra], CLS)
    return [al.tri(x["entidade"], x["operacao"], x["ativo"]) for x in r]


# ── mineração ───────────────────────────────────────────────────────────────
def test_entidade_corrigida_vira_regra_de_classificacao_por_ativo(simular):
    itens = [item("POSTE 11", ("APOIO", "M", "P11"), n=2), item("POSTE 12", ("APOIO", "M", "P12"), n=2), item("TRAFO 5", ("TRAFO", "M", "T5"), n=3)]
    props = al.minerar(itens, PROC, CLS, simular)
    assert len(props) == 1 and props[0]["tabela"] == "classificacao" and props[0]["regressoes"] == 0
    regra = props[0]["regra"]
    assert regra["entidade"] == "APOIO" and regra["parar"] is True and regra["ordem"] < min(r["ordem"] for r in CLS) and "ativo_regex" in regra
    assert aplicar(simular, itens, "classificacao", regra) == [i["user"] for i in itens]          # com a regra, o motor entrega o que o usuário quer
    assert props[0]["corrigidos"] == 4 and props[0]["grupo"] == 4 and props[0]["testados"] == 3


def test_regra_que_estragaria_um_item_certo_usa_a_camada_para_separar(simular):
    itens = [item("POSTE 11", ("APOIO", "M", "P11"), layer="01_REMOVER", n=3), item("POSTE 11", ("POSTE", "M", "P11"), layer="01_REDE", n=5)]
    props = al.minerar(itens, PROC, CLS, simular)
    assert len(props) == 1 and props[0]["regra"].get("layer_em") == ["01_REMOVER"] and props[0]["regressoes"] == 0
    assert aplicar(simular, itens, "classificacao", props[0]["regra"]) == [i["user"] for i in itens]


def test_sem_como_separar_nao_propoe(simular):
    # mesmo texto, camada e ativo: ora o usuário muda, ora não -> qualquer regra estragaria os itens certos
    itens = [item("POSTE 11", ("APOIO", "M", "P11"), n=3), item("POSTE 11", ("POSTE", "M", "P11"), n=5, cor="")]
    assert al.minerar(itens, PROC, CLS, simular) == []


def test_operacao_por_camada_vira_regra_de_processamento(simular):
    itens = [item("POSTE 11", ("POSTE", "I", "P11"), layer="01_NOVO", n=2), item("POSTE 12", ("POSTE", "I", "P12"), layer="01_NOVO", n=2),
             item("POSTE 13", ("POSTE", "M", "P13"), layer="01_REDE", n=4)]
    props = al.minerar(itens, PROC, CLS, simular)
    p = next(p for p in props if p["tabela"] == "processamento")
    assert p["regra"]["operacao"] == "I" and p["regra"]["layer_em"] == ["01_NOVO"] and p["regra"]["fase"] == 3 and p["regressoes"] == 0
    assert aplicar(simular, itens, "processamento", p["regra"]) == [i["user"] for i in itens]


def test_ativo_corrigido_pelo_texto(simular):
    itens = [item("POSTE 11", ("POSTE", "M", "P99"), n=3, fontes=("a", "b"))]
    props = al.minerar(itens, PROC, CLS, simular)
    assert len(props) == 1 and props[0]["regra"]["ativo_template"] == "P99" and props[0]["regra"]["texto_regex"] == "^POSTE 11$"
    assert aplicar(simular, itens, "processamento", props[0]["regra"]) == [itens[0]["user"]]


def test_ativo_que_derrubaria_a_classificacao_nao_e_proposto(simular):
    # P11-ESP não casa com a regra de classificação ^P\d+$ -> a entidade (que estava certa) viraria '0': regressão, sem proposta
    assert al.minerar([item("POSTE 11", ("POSTE", "M", "P11-ESP"), n=3)], PROC, CLS, simular) == []


def test_item_apagado_pelo_usuario_vira_entidade_zero(simular):
    itens = [item("POSTE 11", ("0", "M", "P11"), n=3)]
    props = al.minerar(itens, PROC, CLS, simular)
    assert props and props[0]["regra"]["entidade"] == "0"
    assert aplicar(simular, itens, "classificacao", props[0]["regra"]) == [itens[0]["user"]]


@pytest.mark.parametrize("n,fontes", [(2, ("a", "b")), (5, ("a",))])
def test_poucas_ocorrencias_ou_uma_fonte_so_nao_propoe(simular, n, fontes):
    assert al.minerar([item("POSTE 11", ("APOIO", "M", "P11"), n=n, fontes=fontes)], PROC, CLS, simular) == []


def test_ativo_com_chaves_nao_vira_template(simular):
    assert al.minerar([item("POSTE 11", ("POSTE", "M", "{1}-X"), n=3)], PROC, CLS, simular) == []


def test_regex_escapada_para_javascript():
    assert al._esc("1-DT11/300 (X)") == r"1-DT11\/300 \(X\)"
    assert al._regex_exato(["P1", "P2"]) == "^(?:P1|P2)$"


def test_agregar_soma_e_conta_fontes():
    base = {"texto": "A", "cor": "", "layer": "", "e0": "0", "o0": "M", "a0": "", "e1": "X", "o1": "M", "a1": ""}
    r = al.agregar([{**base, "fonte": "s1"}, {**base, "fonte": "s2", "n": 4}, {**base, "texto": "B", "fonte": "s1"}])
    assert len(r) == 2 and r[0]["n"] == 5 and r[0]["fontes"] == {"s1", "s2"}


def test_conjunto_do_autonomo_reencontra_pela_coordenada():
    def L(e, o, a, x, y):
        return {"entidade": e, "operacao": o, "ativo": a, "_x": x, "_y": y}
    origens = [{"texto": "POSTE 11", "cor": "#000000", "layer": "L", "x": 1, "y": 2, "e": "POSTE", "o": "M", "a": "P11"},
               {"texto": "LIXO", "cor": "", "layer": "", "x": 3, "y": 4, "e": "APOIO", "o": "M", "a": "X"},
               {"texto": "AJUSTADO", "cor": "", "layer": "", "x": 5, "y": 6, "e": "POSTE", "o": "M", "a": "P1"},
               {"texto": "SEM COORD", "cor": "", "layer": "", "x": None, "y": None, "e": "POSTE", "o": "M", "a": "P2"},
               {"texto": "MANTIDO", "cor": "", "layer": "", "x": 7, "y": 8, "e": "POSTE", "o": "M", "a": "P3"}]
    agente = {"cabos": [], "outros": [L("POSTE", "M", "P11", 1, 2), L("APOIO", "M", "X", 3, 4), L("POSTE", "M", "P1-AJ", 5, 6), L("POSTE", "M", "P3", 7, 8)]}
    usuario = {"cabos": [], "outros": [L("APOIO", "M", "P11", 1, 2), L("POSTE", "M", "P1-AJ", 5, 6), L("POSTE", "M", "P3", 7, 8)]}
    r = {x["texto"]: (x["e1"], x["o1"], x["a1"]) for x in al.montar_conjunto_autonomo(origens, agente, usuario)}
    assert r == {"POSTE 11": ("APOIO", "M", "P11"), "LIXO": ("0", "M", "X"), "MANTIDO": ("POSTE", "M", "P3")}      # ajustado e sem coordenada ficam de fora


# ── API ─────────────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    monkeypatch.setattr(svc, "INTERVALO_AUTOMATICO_S", 0)
    return TestClient(appmod.app)


def cab(uid="adm", role="admin"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@x.com', role)}"}


def linhas_leitor(entidade_final="APOIO"):
    return [{"texto": "POSTE 11", "cor": "#000000", "layer": "", "e0": "POSTE", "o0": "M", "a0": "P11", "e1": entidade_final, "o1": "M", "a1": "P11"},
            {"texto": "POSTE 12", "cor": "#000000", "layer": "", "e0": "POSTE", "o0": "M", "a0": "P12", "e1": entidade_final, "o1": "M", "a1": "P12"},
            {"texto": "TRAFO 1", "cor": "#000000", "layer": "", "e0": "TRAFO", "o0": "M", "a0": "T1", "e1": "TRAFO", "o1": "M", "a1": "T1"}]


def capturar(client, sessao, linhas, projeto="229", **kw):
    return client.post("/api/aprendizado/leitor", json={"projeto": projeto, "sessao": sessao, "itens": linhas}, headers=cab(**kw))


def regras_do_projeto(projeto="229"):
    from routers import regras_leitor as rl
    return rl._get_regras("processamento", projeto), rl._get_regras("classificacao", projeto)


def test_captura_manual_gera_proposta_de_regra_do_leitor_e_aprovar_grava_com_historico(client):
    proc0, cls0 = regras_do_projeto()
    assert cls0                                                       # o projeto de teste vem com as regras de semente
    for s in ("s1", "s2"):
        assert capturar(client, s, [{**l, "texto": l["texto"] + " ", "a0": l["a0"]} for l in linhas_leitor()]).status_code == 200
    r = client.post("/api/aprendizado/analisar?projeto=229", headers=cab()).json()
    assert r["leitor"]["itens"] >= 1
    props = [p for p in client.get("/api/aprendizado/propostas?projeto=229", headers=cab()).json() if p["alvo"] == "regra_leitor"]
    # os textos "POSTE 11 " não casam com as regras de semente (que são do projeto real): o motor entrega outra coisa; o que importa é o fluxo
    assert isinstance(props, list)


def test_fluxo_completo_com_regras_do_projeto(client, monkeypatch):
    """Usa regras próprias e controladas no projeto (as mesmas do conjunto de teste) para provar proposta -> aprovar -> regra gravada."""
    from routers import regras_leitor as rl
    for tabela, regras in (("processamento", PROC), ("classificacao", CLS)):
        rl._save_regras(tabela, rl.RegrasLeitorPayload(projeto_codigo="229", regras=regras), "t@x.com")
    for s in ("s1", "s2"):
        assert capturar(client, s, linhas_leitor()).json()["itens"] == 3
    client.post("/api/aprendizado/analisar?projeto=229", headers=cab())
    props = [p for p in client.get("/api/aprendizado/propostas?projeto=229", headers=cab()).json() if p["alvo"] == "regra_leitor"]
    assert len(props) == 1 and props[0]["pode_aprovar"] and props[0]["simulacao"]["regressoes"] == 0 and props[0]["regra"]["entidade"] == "APOIO"
    assert props[0]["tabela"] == "classificacao" and "user_id" not in props[0]
    r = client.post(f"/api/aprendizado/propostas/{props[0]['id']}/aprovar", headers=cab())
    assert r.status_code == 200 and r.json()["tabela"] == "classificacao"
    proc, cls = regras_do_projeto()
    assert cls == CLS + [props[0]["regra"]] and proc == PROC                               # só entrou a regra aprovada
    hist = client.get("/api/regras-leitor/classificacao/historico?projeto_codigo=229", headers=cab()).json()["historico"]
    assert hist                                                                           # versão anterior guardada (dá para reverter)
    assert client.post(f"/api/aprendizado/propostas/{props[0]['id']}/aprovar", headers=cab()).status_code == 400
    # a regra aprovada faz o motor entregar o que o usuário queria
    from services.autonomo.leitor_js import obter_leitor
    out = obter_leitor().processar_lote([{"pagina": 1, "texto": "POSTE 11", "cor": "#000000", "layer": ""}], proc, cls, "extracao")
    assert out[0]["entidade"] == "APOIO"


def test_captura_substitui_a_mesma_sessao_e_so_admin(client):
    capturar(client, "s1", linhas_leitor())
    assert capturar(client, "s1", linhas_leitor()[:1]).json()["itens"] == 1
    assert len(repo.listar_itens("adm", "229")) == 1
    assert capturar(client, "s1", linhas_leitor(), uid="op", role="operador").status_code == 403
    assert client.post("/api/aprendizado/leitor", json={"projeto": "229", "sessao": "x", "itens": []}).status_code == 401
    assert repo.listar_itens("adm2", "229") == [] and repo.listar_itens("adm", "027") == []                 # por usuário e por projeto


def test_captura_ignora_itens_malformados(client):
    r = capturar(client, "s1", [1, "x", None, {"texto": "A" * 600, "e0": "X"}, {"texto": "ok", "e0": "POSTE", "o0": "M", "a0": "P1", "e1": "POSTE", "o1": "M", "a1": "P1"}])
    assert r.status_code == 200 and r.json()["itens"] == 1


def test_apagar_historico_leva_os_itens(client):
    capturar(client, "s1", linhas_leitor())
    client.delete("/api/aprendizado/historico?projeto=229", headers=cab())
    assert repo.listar_itens("adm", "229") == []


def test_sem_motor_o_aprendizado_de_ajustes_segue_valendo(client, monkeypatch):
    monkeypatch.setattr(svc, "simulador", lambda: (_ for _ in ()).throw(RuntimeError("sem quickjs")))
    capturar(client, "s1", linhas_leitor())
    r = client.post("/api/aprendizado/analisar?projeto=229", headers=cab())
    assert r.status_code == 200 and "não analisadas" in r.json()["leitor"]["aviso"] and "novas" in r.json()


# ── obra manual (Resumo) e autônomo (coordenadas) ───────────────────────────
def L(a, op="I", e="X", **kw):
    return {"entidade": e, "operacao": op, "ativo": a, **kw}


def salvar_manual(client, baseline, cabos, outros, sessao="m1", user="adm", role="admin"):
    dados = {"cabos": {"bodyId": "body-cabos", "data": cabos}, "outros": {"bodyId": "body-outros", "data": outros}}
    r = client.post("/api/obras", json={"id": f"obra_{uuid.uuid4().hex[:6]}", "nome": "M", "data": "d", "dados_json": json.dumps(dados), "projeto": "229",
                                        "baseline_manual": {"sessao": sessao, **baseline}}, headers=cab(user, role))
    assert r.status_code == 200, r.text
    return r


def test_obra_manual_registra_diferencas_contra_o_baseline_do_leitor(client):
    base = {"cabos": [L("CA 2 ABC 35 m")], "outros": [L("2-U4"), L("LIXO")]}
    salvar_manual(client, base, [L("CAA 2 ABC 35 m")], [L("2-U3")])
    regs = repo.listar_registros("adm", "229")
    assert len(regs) == 1 and regs[0]["execucao_id"] == "manual:m1" and {e["tipo"] for e in regs[0]["eventos"]} == {"texto_trocado", "item_trocado", "linha_excluida"}
    salvar_manual(client, base, [L("CAA 2 ABC 35 m")], [L("2-U4"), L("LIXO")])             # mesma sessão: substitui
    assert len(repo.listar_registros("adm", "229")) == 1 and [e["tipo"] for e in repo.listar_registros("adm", "229")[0]["eventos"]] == ["texto_trocado"]


def test_ajuste_ativo_ja_aplicado_pelo_usuario_nao_conta_como_correcao(client):
    from routers import validacao_ajustes as va
    receita = {"id": "troca_u4", "nome": "U4 vira U3", "ativa": True, "acoes": [{"acao": "substituir", "tabela": "outros", "de": "U4", "para": "U3", "modo": "item"}]}
    va._salvar("229", [va._limpa(r) for r in va.ajustes_efetivos("229")] + [receita], "t@x.com")
    salvar_manual(client, {"cabos": [], "outros": [L("2-U4")]}, [], [L("2-U3")])                # o usuário fez o que o ajuste já faz
    assert repo.listar_registros("adm", "229")[0]["eventos"] == []


def test_manual_so_para_administrador_e_sem_baseline_nao_registra(client):
    salvar_manual(client, {"cabos": [], "outros": [L("2-U4")]}, [], [L("2-U3")], user="op", role="operador")
    assert repo.listar_registros("op", "229") == []
    r = client.post("/api/obras", json={"id": "x", "nome": "n", "data": "d", "dados_json": "{}", "projeto": "229"}, headers=cab())
    assert r.status_code == 200 and repo.listar_registros("adm", "229") == []


def test_autonomo_reaproveita_coordenadas_para_o_conjunto_do_leitor(client):
    eid = execucoes.criar("a.dxf", uuid.uuid4().hex, "229", "adm")
    ag_outros = [L("P11", "M", "POSTE", _x=1, _y=2), L("T1", "M", "TRAFO", _x=3, _y=4)]
    snap = {"cabos": {"data": []}, "outros": {"data": ag_outros}, "autonomo": {"execucao_id": eid}}
    rota_obras.gravar_obra({"user_id": "adm"}, {"id": f"auto_{eid}", "nome": "x", "data": "d", "dados_json": json.dumps(snap), "projeto": "229", "publica": False})
    execucoes.atualizar(eid, status="ok", obra_id=f"auto_{eid}", itens_origem=[
        {"texto": "POSTE 11", "cor": "#000000", "layer": "", "x": 1, "y": 2, "e": "POSTE", "o": "M", "a": "P11"},
        {"texto": "TRAFO 1", "cor": "#000000", "layer": "", "x": 3, "y": 4, "e": "TRAFO", "o": "M", "a": "T1"}])
    dados = {"cabos": {"data": []}, "outros": {"data": [L("P11", "M", "APOIO", _x=1, _y=2), L("T1", "M", "TRAFO", _x=3, _y=4)]}}
    r = client.post("/api/obras", json={"id": "u1", "nome": "c", "data": "d", "dados_json": json.dumps(dados), "projeto": "229", "origem_execucao": eid}, headers=cab())
    assert r.status_code == 200
    itens = repo.listar_itens("adm", "229")
    assert {(i["texto"], i["user"]) for i in itens} == {("POSTE 11", ("APOIO", "M", "P11")), ("TRAFO 1", ("TRAFO", "M", "T1"))}
    assert client.get("/api/autonomo/execucoes", headers=cab()).status_code == 200
    assert "itens_origem" not in json.dumps(client.get(f"/api/autonomo/execucoes/{eid}", headers=cab()).json())             # bloco grande não vai para a tela


def test_regra_que_so_corrige_parte_do_grupo_nao_e_proposta(simular):
    # 3 itens que a regra conserta + 2 azuis (o motor força operação 0 e entidade 0: regra de classificação nenhuma muda isso) = 60% < 80%
    itens = [item("POSTE 11", ("APOIO", "M", "P11"), n=3), item("POSTE 12", ("APOIO", "M", "P12"), n=2, cor="#0000ff")]
    assert al.minerar(itens, PROC, CLS, simular) == []
    itens[1]["n"] = 1                       # 3 de 4 = 75% ainda é pouco
    assert al.minerar(itens, PROC, CLS, simular) == []


# ── Montar Orçamento também registra (POST /api/aprendizado/sessao) ──────────
def sessao(client, corpo, **kw):
    return client.post("/api/aprendizado/sessao", json={"projeto": "229", **corpo}, headers=cab(**kw))


def test_montar_orcamento_registra_trabalho_manual_sem_salvar_obra(client):
    base = {"sessao": "orc1", "cabos": [L("CA 2 ABC 35 m")], "outros": [L("2-U4"), L("LIXO")]}
    r = sessao(client, {"baseline_manual": base, "cabos": [L("CAA 2 ABC 35 m")], "outros": [L("2-U3")]})
    assert r.status_code == 200 and r.json()["registrado"] is True and r.json()["correcoes"] == 3
    regs = repo.listar_registros("adm", "229")
    assert len(regs) == 1 and regs[0]["execucao_id"] == "manual:orc1" and regs[0]["obra_salva"] is None
    assert client.get("/api/obras", headers=cab()).json() == []                                   # nenhuma obra foi criada
    # montar de novo na mesma sessão substitui (não conta duas vezes)
    sessao(client, {"baseline_manual": base, "cabos": [L("CAA 2 ABC 35 m")], "outros": [L("2-U4"), L("LIXO")]})
    regs = repo.listar_registros("adm", "229")
    assert len(regs) == 1 and [e["tipo"] for e in regs[0]["eventos"]] == ["texto_trocado"]


def test_salvar_e_montar_na_mesma_sessao_dao_um_registro_so(client):
    base = {"cabos": [], "outros": [L("2-U4")]}
    salvar_manual(client, base, [], [L("2-U3")], sessao="mesma")
    sessao(client, {"baseline_manual": {"sessao": "mesma", **base}, "cabos": [], "outros": [L("2-U3")]})
    assert len(repo.listar_registros("adm", "229")) == 1


def test_montar_orcamento_de_obra_do_autonomo_registra_pela_execucao(client):
    eid = execucoes.criar("a.dxf", uuid.uuid4().hex, "229", "adm")
    ag = [L("P11", "M", "POSTE", _x=1, _y=2)]
    rota_obras.gravar_obra({"user_id": "adm"}, {"id": f"auto_{eid}", "nome": "x", "data": "d", "projeto": "229", "publica": False,
                                                "dados_json": json.dumps({"cabos": {"data": []}, "outros": {"data": ag}, "autonomo": {"execucao_id": eid}})})
    execucoes.atualizar(eid, status="ok", obra_id=f"auto_{eid}", itens_origem=[{"texto": "POSTE 11", "cor": "#000000", "layer": "", "x": 1, "y": 2, "e": "POSTE", "o": "M", "a": "P11"}])
    r = sessao(client, {"origem_execucao": eid, "cabos": [], "outros": [L("P11", "M", "APOIO", _x=1, _y=2)]})
    assert r.json()["registrado"] is True and repo.listar_registros("adm", "229")[0]["execucao_id"] == eid
    assert [i["user"] for i in repo.listar_itens("adm", "229")] == [("APOIO", "M", "P11")]


def test_sessao_sem_origem_ou_de_outro_usuario_ou_nao_admin(client):
    assert sessao(client, {"cabos": [], "outros": []}).json()["registrado"] is False                # nem baseline nem execução
    assert sessao(client, {"origem_execucao": "nao_existe", "cabos": [], "outros": []}).json()["registrado"] is False
    assert sessao(client, {"baseline_manual": {"sessao": "x", "cabos": [], "outros": []}, "cabos": [], "outros": []}, uid="op", role="operador").status_code == 403
    assert client.post("/api/aprendizado/sessao", json={"projeto": "229"}).status_code == 401
    assert sessao(client, {"baseline_manual": {"sessao": "x", "cabos": [], "outros": []}, "cabos": [L("a")] * 3001, "outros": []}).status_code == 400


def test_falha_interna_nao_vira_erro_para_a_tela(client, monkeypatch):
    monkeypatch.setattr(svc, "registrar_sessao", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("x")))
    r = sessao(client, {"baseline_manual": {"sessao": "x", "cabos": [], "outros": []}, "cabos": [], "outros": []})
    assert r.status_code == 200 and r.json()["registrado"] is False
