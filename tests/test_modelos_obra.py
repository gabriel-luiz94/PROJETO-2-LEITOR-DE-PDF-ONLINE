"""TASK-058 — modelos de obra com variável V: detecção, geração, API (permissões), exportação/importação, validação, IA."""
import json

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.ai_chat as chat
import routers.obras as rota
from middleware.auth_middleware import create_jwt_token
from services import modelos_obra as mo
from services import obras_arquivo as oa
from services import obras_contexto as oc
from services.validacao_planilhas import validar_planilhas


def L(ativo, op="I"):
    return {"entidade": "X", "operacao": op, "ativo": ativo}


# ── unidades ────────────────────────────────────────────────────────────────
def test_detecta_so_o_token_exato():
    cab = [L("CAA 2 ABC V m"), L("CAA 2 ABC V(vão) m"), L("CAV 3 ABC 10 m"), L("CAA 2 ABC V1 m"), L("CAA 2 ABC 10 m"), L("V")]
    out = [L("V-U4"), L("*V(postes)-CFU 2-U3"), L("1-VA"), L("V1-X"), L("CAV-Y"), L("2-U4 V(vão)-DT11/300")]
    v = mo.detectar_variaveis(cab, out)
    assert [x["chave"] for x in v] == ["#1", "vão", "#2", "postes"]
    assert [len(x["ocorrencias"]) for x in v] == [1, 2, 1, 1]            # V(vão) em Cabos e em Outros = um campo
    assert mo.detectar_variaveis([L("CAA 2 ABC 10 m")], [L("2-U4")]) == []
    assert mo.tem_variavel("CAA 2 ABC V m", "cabos") and not mo.tem_variavel("CAV 2 ABC 3 m", "cabos") and mo.tem_variavel("*V-CFU", "outros")


def test_nomeados_sao_um_campo_sem_distinguir_maiusculas():
    v = mo.detectar_variaveis([L("CAA 2 ABC V(Vão) m")], [L("V(VÃO)-U4")])
    assert len(v) == 1 and v[0]["chave"] == "vão" and len(v[0]["ocorrencias"]) == 2


def test_gera_cabos_outros_negativas_e_decimais():
    cab = [L("CAA 2 ABC V(vão) m"), L("CAA 2 ABC V(vão) m"), L("CAA 2 ABC 10 m")]
    out = [L("2-U4 V-CFU"), L("*V(postes)-DT11/300"), L("1-VA")]
    g = mo.gerar(cab, out, [], {"vão": "85,5", "#1": 3, "postes": "2"})
    assert [x["ativo"] for x in g["cabos"]] == ["CAA 2 ABC 85.5 m"] * 2 + ["CAA 2 ABC 10 m"]
    assert [x["ativo"] for x in g["outros"]] == ["2-U4 3-CFU", "*2-DT11/300", "1-VA"]
    assert g["cabos"][0]["operacao"] == "I" and g["cabos"][0]["entidade"] == "X"
    assert not any(mo.tem_variavel(x["ativo"], "cabos") for x in g["cabos"]) and cab[0]["ativo"] == "CAA 2 ABC V(vão) m"   # original intacto
    assert mo.formatar(3.0) == "3" and mo.formatar(0.12345) == "0.1235" and mo.formatar(85.50) == "85.5"


def test_cabo_sem_m_e_usa_padrao_quando_nao_informado():
    cfg = [{"chave": "#1", "rotulo": "Vão", "padrao": 40}]
    g = mo.gerar([L("CAA 2 ABC V")], [], cfg, {})
    assert g["cabos"][0]["ativo"] == "CAA 2 ABC 40"
    assert mo.gerar([L("CAA 2 ABC V m")], [], cfg, {"#1": "7"})["cabos"][0]["ativo"] == "CAA 2 ABC 7 m"        # informado vence o padrão


@pytest.mark.parametrize("valor", ["", None, "0", "-5", "abc", "1e3", "5 m", "1.2.3", "9999999", True, "nan", "inf"])
def test_valores_invalidos_sao_recusados(valor):
    with pytest.raises(mo.ErroModelo):
        mo.gerar([L("CAA 2 ABC V m")], [], [], {"#1": valor})


def test_valor_desconhecido_e_erro_e_mensagem_cita_o_rotulo():
    with pytest.raises(mo.ErroModelo) as e:
        mo.gerar([L("CAA 2 ABC V(vão) m")], [], [{"chave": "vão", "rotulo": "Comprimento do vão"}], {"outra": 1})
    assert any("Comprimento do vão" in m for m in e.value.mensagens) and any("desconhecida" in m for m in e.value.mensagens)


def test_resultado_e_validado_pela_camada_1():
    # a linha tocada fica inválida (operação) mesmo com V preenchido -> a geração falha com a lista de erros
    with pytest.raises(mo.ErroModelo) as e:
        mo.gerar([L("CAA 2 ABC V m", op="ZZ")], [], [], {"#1": 5})
    assert "inválid" in e.value.mensagens[0]
    # erro em linha que NÃO tem V não impede a geração
    assert mo.gerar([L("CAA 2 ABC V m"), L("CAA 2 ABC 5 m", op="ZZ")], [], [], {"#1": 5})


def test_validacao_aponta_variavel_sobrando_com_mensagem_clara():
    ach = validar_planilhas([L("CAA 2 ABC V m")], [L("*V(x)-CFU"), L("2-U4")])
    assert [(a["tabela"], a["regra_id"], a["severidade"]) for a in ach] == [("cabos", "C1-VAR", "erro"), ("outros", "C1-VAR", "erro")]
    assert "variável de modelo" in ach[0]["mensagem"]
    assert validar_planilhas([L("CAA 2 ABC 5 m")], [L("2-U4"), L("1-VA")]) == []


def test_configuracao_de_parametros():
    assert mo.validar_configuracao([{"chave": "#1", "rotulo": "A", "padrao": "40,5"}]) == [{"chave": "#1", "rotulo": "A", "padrao": 40.5}]
    for ruim in ({"chave": "#1", "padrao": "-1"}, {"chave": "#1", "rotulo": "x" * 81}, "x", [1], [{"rotulo": "a"}]):
        with pytest.raises(mo.ErroModelo):
            mo.validar_configuracao(ruim if isinstance(ruim, (str, list)) else [ruim])


# ── API ─────────────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(uid, role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@empresa.com', role)}"}


def dados(cabos=("CAA 2 ABC V(vão) m", "CAA 2 ABC 10 m"), outros=("2-U4", "V(postes)-DT11/300")):
    return json.dumps({"cabos": {"bodyId": "body-cabos", "data": [L(a) for a in cabos]},
                       "outros": {"bodyId": "body-outros", "data": [L(a) for a in outros]},
                       "totalizadora": {"bodyId": "body-totalizadora", "data": [{"id": "T"}]}})


def salvar_modelo(client, uid="a", oid="M1", nome="Extensão BT", projeto="229", dj=None, params=None, **extra):
    corpo = {"id": oid, "nome": nome, "data": "d", "dados_json": dj or dados(), "projeto": projeto, "tipo": "modelo", **extra}
    if params is not None:
        corpo["parametros"] = params
    return client.post("/api/obras", json=corpo, headers=cab(uid))


def test_salvar_modelo_e_sempre_publico_sem_totalizadora_e_guarda_parametros(client):
    r = salvar_modelo(client, params=[{"chave": "vão", "rotulo": "Vão da rede", "padrao": "40"}], publica=False)
    assert r.status_code == 200 and r.json()["tipo"] == "modelo" and r.json()["publica"] is True
    o = client.get("/api/obras/M1?projeto=229", headers=cab("b")).json()                 # outro usuário do projeto enxerga
    assert o["tipo"] == "modelo" and o["publica"] is True
    snap = json.loads(o["dados_json"])
    assert "totalizadora" not in snap and snap["modelo"]["parametros"][0] == {"chave": "vão", "rotulo": "Vão da rede", "padrao": 40.0}
    assert client.get("/api/obras/M1?projeto=027", headers=cab("b")).status_code == 404       # outro projeto não vê


def test_modelo_exige_variavel_e_parametros_validos(client):
    sem_v = dados(cabos=("CAA 2 ABC 10 m",), outros=("2-U4",))
    r = salvar_modelo(client, dj=sem_v)
    assert r.status_code == 400 and "ao menos uma variável" in r.json()["detail"]["erros"][0]
    assert salvar_modelo(client, params=[{"chave": "vão", "padrao": "-3"}]).status_code == 400
    assert salvar_modelo(client, dj="não é json").status_code == 400
    assert client.post("/api/obras", json={"id": "x", "nome": "n", "data": "d", "dados_json": "{}", "projeto": "229", "tipo": "xx"}, headers=cab("a")).status_code == 400
    assert client.get("/api/obras?projeto=229", headers=cab("a")).json() == []                 # nada foi gravado


def test_so_o_criador_edita_ou_exclui_o_modelo(client):
    salvar_modelo(client)
    assert salvar_modelo(client, uid="b", nome="Sequestrado").status_code == 403
    assert client.delete("/api/obras/M1", headers=cab("b")).status_code == 200                # a rota é silenciosa para quem não é dono...
    assert client.get("/api/obras/M1?projeto=229", headers=cab("a")).json()["nome"] == "Extensão BT"       # ...e nada acontece
    assert client.put("/api/obras/M1/visibilidade", json={"publica": False}, headers=cab("a")).status_code == 400   # sempre público
    assert salvar_modelo(client, nome="Nova versão").status_code == 200                       # criador sobrescreve
    assert client.get("/api/obras/M1?projeto=229", headers=cab("b")).json()["nome"] == "Nova versão"
    client.delete("/api/obras/M1", headers=cab("a"))
    assert client.get("/api/obras/M1?projeto=229", headers=cab("b")).status_code == 404


def test_editar_mantem_os_parametros_quando_nao_enviados(client):
    salvar_modelo(client, params=[{"chave": "vão", "rotulo": "Vão da rede", "padrao": "40"}])
    salvar_modelo(client, nome="Editado")                                                      # sem `parametros`: mantém
    p = client.get("/api/obras/M1/modelo?projeto=229", headers=cab("b")).json()["parametros"]
    assert {x["chave"]: x["rotulo"] for x in p}["vão"] == "Vão da rede"
    assert next(x for x in p if x["chave"] == "vão")["onde"][0].startswith("Cabos linha 1")


def test_gerar_aplica_os_valores_sem_gravar_e_respeita_o_acesso(client):
    salvar_modelo(client, params=[{"chave": "vão", "rotulo": "Vão", "padrao": 40}])
    url = "/api/obras/M1/gerar?projeto=229"
    r = client.post(url, json={"valores": {"vão": "85", "postes": "3"}}, headers=cab("b"))
    assert r.status_code == 200
    g = r.json()
    assert [x["ativo"] for x in g["cabos"]] == ["CAA 2 ABC 85 m", "CAA 2 ABC 10 m"] and [x["ativo"] for x in g["outros"]] == ["2-U4", "3-DT11/300"]
    assert client.post(url, json={"valores": {"postes": "3"}}, headers=cab("b")).json()["cabos"][0]["ativo"] == "CAA 2 ABC 40 m"   # padrão
    r = client.post(url, json={"valores": {"vão": "0"}}, headers=cab("b"))
    assert r.status_code == 400 and r.json()["detail"]["erros"]
    assert [o["nome"] for o in client.get("/api/obras?projeto=229&origem=minhas", headers=cab("b")).json()] == []     # nada gravado para B
    assert client.post("/api/obras/M1/gerar", json={"valores": {}}, headers=cab("b")).status_code == 404               # sem projeto: invisível
    assert client.post("/api/obras/M1/gerar?projeto=027", json={"valores": {}}, headers=cab("b")).status_code == 404
    assert client.post("/api/obras/nao/gerar?projeto=229", json={"valores": {}}, headers=cab("b")).status_code == 404
    assert json.loads(client.get("/api/obras/M1?projeto=229", headers=cab("a")).json()["dados_json"])["cabos"]["data"][0]["ativo"] == "CAA 2 ABC V(vão) m"


def test_gerar_em_obra_comum_e_recusado_e_exige_login(client):
    client.post("/api/obras", json={"id": "O1", "nome": "Comum", "data": "d", "dados_json": dados(), "projeto": "229"}, headers=cab("a"))
    r = client.post("/api/obras/O1/gerar?projeto=229", json={"valores": {"vão": 5, "postes": 2}}, headers=cab("a"))
    assert r.status_code == 400 and "não é um modelo" in r.json()["detail"]
    assert client.post("/api/obras/O1/gerar?projeto=229", json={"valores": {}}).status_code == 401


def test_detectar_para_a_janela_de_salvar(client):
    r = client.post("/api/obras/modelo/detectar", json={"cabos": [L("CAA 2 ABC V m")], "outros": [L("*V-CFU")]}, headers=cab("a"))
    assert r.status_code == 200 and [p["chave"] for p in r.json()["parametros"]] == ["#1", "#2"]
    assert r.json()["parametros"][0]["rotulo"] == "Valor 1"
    assert client.post("/api/obras/modelo/detectar", json={"cabos": [], "outros": []}).status_code == 401


def test_filtro_por_tipo_e_selo(client):
    salvar_modelo(client)
    client.post("/api/obras", json={"id": "O1", "nome": "Comum", "data": "d", "dados_json": dados(cabos=("CAA 2 ABC 5 m",), outros=("2-U4",)), "projeto": "229"}, headers=cab("a"))
    ids = lambda t, u="a": sorted(o["id"] for o in client.get(f"/api/obras?projeto=229&tipo={t}", headers=cab(u)).json())
    assert ids("todos") == ["M1", "O1"] and ids("modelos") == ["M1"] and ids("obras") == ["O1"]
    assert ids("todos", "b") == ["M1"] and ids("obras", "b") == []                             # B só enxerga o modelo (público)
    assert client.get("/api/obras/indice?projeto=229", headers=cab("b")).json()[0]["tipo"] == "modelo"


def test_obra_comum_continua_igual(client):
    client.post("/api/obras", json={"id": "O1", "nome": "Comum", "data": "d", "dados_json": dados(), "projeto": "229"}, headers=cab("a"))
    o = client.get("/api/obras/O1", headers=cab("a")).json()
    assert o["tipo"] == "obra" and o["publica"] is False and "totalizadora" in json.loads(o["dados_json"])        # nada mudou para obras comuns


# ── exportar / importar ─────────────────────────────────────────────────────
def test_exportar_e_importar_modelo_vira_modelo_do_importador(client):
    salvar_modelo(client, params=[{"chave": "vão", "rotulo": "Vão da rede", "padrao": "40"}])
    arq = client.get("/api/obras/M1/exportar?projeto=229", headers=cab("b")).json()                     # público: B exporta
    assert arq["tipo"] == "modelo" and "totalizadora" not in arq["dados"] and "a@" not in json.dumps(arq)
    assert {p["chave"]: p["rotulo"] for p in arq["parametros"]}["vão"] == "Vão da rede"
    r = client.post("/api/obras/importar?projeto=229", content=json.dumps(arq), headers={**cab("b"), "Content-Type": "application/json"})
    assert r.status_code == 200 and r.json()["tipo"] == "modelo"
    novo = client.get(f"/api/obras/{r.json()['id']}?projeto=229", headers=cab("b")).json()
    assert novo["tipo"] == "modelo" and novo["minha"] is True and novo["nome"] == "Extensão BT"
    assert client.delete(f"/api/obras/{r.json()['id']}", headers=cab("b")).status_code == 200                # B é o dono do novo
    assert client.get("/api/obras/M1?projeto=229", headers=cab("a")).status_code == 200


def test_importar_modelo_invalido(client):
    base = {"formato": "leitor-obra", "versao": 1, "tipo": "modelo", "nome": "M", "projeto": "229",
            "dados": {"cabos": [L("CAA 2 ABC 5 m")], "outros": [L("2-U4")]}, "parametros": []}
    imp = lambda c: client.post("/api/obras/importar?projeto=229", content=json.dumps(c), headers={**cab("b"), "Content-Type": "application/json"})
    r = imp(base)
    assert r.status_code == 400 and "nenhuma variável" in r.json()["detail"]["erros"][0]
    base["dados"]["cabos"] = [L("CAA 2 ABC V m")]
    base["parametros"] = [{"chave": "#1", "padrao": "-2"}]
    assert imp(base).status_code == 400
    base["parametros"] = [{"chave": "#1", "rotulo": "Vão", "padrao": 30}]
    assert imp(base).status_code == 200


# ── IA ──────────────────────────────────────────────────────────────────────
def test_indice_da_ia_marca_modelos():
    t = oc.montar_indice([{"id": "1", "nome": "Meu", "data": "d", "minha": True, "tipo": "modelo", "publica": True},
                          {"id": "2", "nome": "Dele", "data": "d", "minha": False, "tipo": "modelo", "publica": True, "dono": "ana"},
                          {"id": "3", "nome": "Comum", "data": "d", "minha": True, "tipo": "obra", "publica": False}])
    assert '| MODELO\n' in t + "\n" and "MODELO de ana" in t and t.endswith('nome="Comum" | data=d')


def test_ia_ve_modelos_do_projeto_e_o_prompt_proibe_inventar_valores(client):
    import types
    salvar_modelo(client)
    req = types.SimpleNamespace(prompt="quais obras eu tenho?", projeto_codigo="229")
    extra, instr = chat._contexto_obras(req, types.SimpleNamespace(state=types.SimpleNamespace(user={"user_id": "b", "email": "b@x.com", "role": "operador"})))
    assert "MODELO de a" in extra and "NUNCA invente valores" in instr


# ── nuvem sem a coluna tipo ─────────────────────────────────────────────────
def test_nuvem_sem_coluna_tipo_ainda_grava_e_reconhece_o_modelo(client, monkeypatch):
    from tests.test_obras_visibilidade import NuvemFalsa
    nuvem = NuvemFalsa()
    original_checa = nuvem._checa
    nuvem._checa = lambda nomes: (_ for _ in ()).throw(RuntimeError("no column")) if "tipo" in set(nomes) else original_checa(nomes)
    monkeypatch.setattr(rota, "get_supabase", lambda: nuvem)
    assert salvar_modelo(client).status_code == 200
    assert nuvem.linhas and "tipo" not in nuvem.linhas[0] and nuvem.linhas[0]["publica"] is True       # recuou para o formato da TASK-057
    o = client.get("/api/obras/M1?projeto=229", headers=cab("b")).json()
    assert o["tipo"] == "modelo"                                                                     # reconhecido pela marca no dados_json
