"""TASK-056 — exportar/importar obra (`.obra.json`): formato, validação, cópia particular e isolamento entre usuários."""
import json

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.obras as rota
from middleware.auth_middleware import create_jwt_token
from services import obras_arquivo as oa


# ── unidades ────────────────────────────────────────────────────────────────
SNAP = {
    "cabos": {"bodyId": "body-cabos", "data": [{"entidade": "CABO", "operacao": "I", "ativo": "CAA 2 ABC 35 m", "qtdAtivos": 3, "_raw": {"x": 1}}, None]},
    "outros": {"bodyId": "body-outros", "data": [{"entidade": "ESTRUTURA", "operacao": "*R", "ativo": "2-U4 1-CFU"},
                                                  {"entidade": "0", "operacao": "I", "ativo": "1-AF", "texto": "RAMAIS (GERADO)", "lixo": [1]}]},
    "totalizadora": {"bodyId": "body-totalizadora", "data": [{"id": "TOT-1", "obs": "CABO", "operacao": "I", "ativo": "CAA2", "qtd": 35.5,
                                                              "desc": "d", "naoEncontrado": False, "origem": "CABOS", "baseId": "TOT-1"}]},
    "regras": {"bodyId": "body-regras", "data": [{"origem": "CABOS", "fator": 1.05}]},
    "autonomo": {"execucao_id": "abc", "arquivo": "x.dxf"},
}


def arquivo(**mudar):
    base = {"formato": "leitor-obra", "versao": 1, "tipo": "obra", "nome": "Rede Norte", "projeto": "229",
            "dados": oa.normalizar_dados(SNAP)}
    base.update(mudar)
    return base


def test_normaliza_as_formas_antigas_e_descarta_o_que_nao_e_da_obra():
    d = oa.normalizar_dados(json.dumps(SNAP))
    assert d["cabos"] == [{"entidade": "CABO", "operacao": "I", "ativo": "CAA 2 ABC 35 m", "qtdAtivos": 3}]
    assert d["outros"][1] == {"entidade": "0", "operacao": "I", "ativo": "1-AF", "texto": "RAMAIS (GERADO)"}
    assert d["totalizadora"][0]["naoEncontrado"] is False and d["totalizadora"][0]["qtd"] == 35.5
    assert oa.normalizar_dados({"cabos": [{"ativo": "P 50", "operacao": "I", "entidade": "CABO"}], "outros": []})["cabos"][0]["ativo"] == "P 50"
    assert oa.normalizar_dados("isto não é json") == {"cabos": [], "outros": [], "totalizadora": []}
    assert oa.normalizar_dados({"cabos": [{"ativo": True, "qtdAtivos": True}]})["cabos"] == [{}]       # bool nunca vira número/texto


def test_exportacao_leva_so_os_dados_da_obra():
    env = oa.montar_exportacao({"nome": "Rede Norte", "projeto": "229", "dados_json": json.dumps(SNAP), "user_id": "u1", "id": "obra_1"})
    assert env["formato"] == "leitor-obra" and env["versao"] == 1 and env["tipo"] == "obra" and env["projeto"] == "229"
    texto = json.dumps(env)
    assert "autonomo" not in texto and "regras" not in texto and "u1" not in texto and "obra_1" not in texto and "_raw" not in texto
    assert set(env["dados"]) == {"cabos", "outros", "totalizadora"}


def test_nome_de_arquivo_e_seguro():
    assert oa.nome_de_arquivo("Rede Norte") == "Rede Norte.obra.json"
    assert oa.nome_de_arquivo("../../etc/passwd") == "etc_passwd.obra.json"
    assert oa.nome_de_arquivo("") == "obra.obra.json" and oa.nome_de_arquivo("a" * 300).endswith(".obra.json") and len(oa.nome_de_arquivo("a" * 300)) < 100


def test_validacao_aceita_arquivo_correto():
    pronta = oa.validar_importacao(arquivo(), "229")
    assert pronta["nome"] == "Rede Norte" and pronta["projeto"] == "229" and len(pronta["dados"]["outros"]) == 2
    assert json.loads(oa.dados_json_para_gravar(pronta["dados"]))["cabos"]["bodyId"] == "body-cabos"      # forma que o Carregar entende



@pytest.mark.parametrize("carga,trecho", [
    ("texto", "objeto JSON"), ({"formato": "outro"}, "formato desconhecido"),
    (arquivo(versao=2), "versão mais nova"), (arquivo(versao="1"), "Versão do arquivo inválida"), (arquivo(versao=True), "Versão do arquivo inválida"),
    (arquivo(tipo="modelo"), "outro tipo"), (arquivo(nome="  "), "nome da obra"), (arquivo(nome="x" * 121), "nome da obra"),
    (arquivo(projeto=""), "não informa o projeto"), (arquivo(projeto="027"), "projeto 027, mas o projeto selecionado é 229"),
    (arquivo(dados="x"), "seção 'dados'"), (arquivo(dados={"cabos": []}), "Outros: seção ausente"),
    (arquivo(dados={"cabos": "x", "outros": []}), "Cabos: precisa ser uma lista"),
    (arquivo(dados={"cabos": [{"ativo": "x" * 501, "operacao": "I"}], "outros": []}), "até 500 caracteres"),
    (arquivo(dados={"cabos": [{"ativo": "P 50", "operacao": "XX"}], "outros": []}), "operação 'XX' inválida"),
    (arquivo(dados={"cabos": [], "outros": [{"ativo": "1-U4", "operacao": "I", "entidade": "<script>"}]}), "'entidade' inválida"),
    (arquivo(dados={"cabos": [1], "outros": []}), "precisa ser um objeto"),
    (arquivo(dados={"cabos": [{"ativo": "a", "operacao": "I"}] * 5001, "outros": []}), "linhas demais"),
])
def test_validacao_recusa_com_mensagem_clara(carga, trecho):
    with pytest.raises(oa.ErroArquivoObra) as e:
        oa.validar_importacao(carga, "229")
    assert any(trecho in m for m in e.value.mensagens), e.value.mensagens


def test_projeto_selecionado_ausente_e_recusado_e_erros_sao_limitados():
    with pytest.raises(oa.ErroArquivoObra, match="nenhum"):
        oa.validar_importacao(arquivo(), None)
    ruins = [{"ativo": "a", "operacao": "ZZ"}] * 100
    with pytest.raises(oa.ErroArquivoObra) as e:
        oa.validar_importacao(arquivo(dados={"cabos": ruins, "outros": []}), "229")
    assert len(e.value.mensagens) <= oa.LIMITE_ERROS


# ── API ─────────────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(uid, role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@x.com', role)}"}


def salvar(client, uid, obra_id="obra_1", nome="Rede Norte", projeto="229", snap=SNAP):
    r = client.post("/api/obras", json={"id": obra_id, "nome": nome, "data": "01/10/2026, 10:00:00", "dados_json": json.dumps(snap), "projeto": projeto}, headers=cab(uid))
    assert r.status_code == 200
    return obra_id


def importar(client, uid, carga, projeto="229"):
    corpo = carga if isinstance(carga, (bytes, str)) else json.dumps(carga)
    url = "/api/obras/importar" + (f"?projeto={projeto}" if projeto else "")
    return client.post(url, content=corpo, headers={**cab(uid), "Content-Type": "application/json"})


def nomes(client, uid, projeto="229"):
    return [o["nome"] for o in client.get(f"/api/obras?projeto={projeto}", headers=cab(uid)).json()]


def test_exige_login(client):
    assert client.get("/api/obras/x/exportar").status_code == 401
    assert client.post("/api/obras/importar?projeto=229", content="{}").status_code == 401


def test_exportar_devolve_arquivo_so_da_obra_do_dono(client):
    salvar(client, "a")
    r = client.get("/api/obras/obra_1/exportar", headers=cab("a"))
    assert r.status_code == 200 and "attachment" in r.headers["content-disposition"] and "Rede%20Norte.obra.json" in r.headers["content-disposition"]
    env = r.json()
    assert env["nome"] == "Rede Norte" and env["projeto"] == "229" and env["dados"]["cabos"][0]["ativo"] == "CAA 2 ABC 35 m"
    assert client.get("/api/obras/obra_1/exportar", headers=cab("b")).status_code == 404       # a obra é só de A
    assert client.get("/api/obras/nao_existe/exportar", headers=cab("a")).status_code == 404


def test_importar_cria_copia_particular_do_importador_e_nao_mexe_na_original(client):
    salvar(client, "a")
    arq = client.get("/api/obras/obra_1/exportar", headers=cab("a")).json()
    r = importar(client, "b", arq)
    assert r.status_code == 200 and r.json()["nome"] == "Rede Norte" and r.json()["id"] != "obra_1" and r.json()["outros"] == 2
    assert nomes(client, "b") == ["Rede Norte"] and nomes(client, "a") == ["Rede Norte"]
    nova = client.get(f"/api/obras/{r.json()['id']}", headers=cab("b")).json()
    assert client.get(f"/api/obras/{r.json()['id']}", headers=cab("a")).status_code == 404       # particular: A não enxerga a cópia de B
    assert client.get("/api/obras/obra_1", headers=cab("b")).status_code == 404                   # nem B a original de A
    # ida e volta: as tabelas da cópia são iguais às da original (Cabos, Outros e Totalizadora)
    orig = oa.normalizar_dados(client.get("/api/obras/obra_1", headers=cab("a")).json()["dados_json"])
    assert oa.normalizar_dados(nova["dados_json"]) == orig and orig["totalizadora"]
    assert nova["projeto"] == "229"


def test_reimportar_nao_sobrescreve_e_numera_o_nome(client):
    salvar(client, "a")
    arq = client.get("/api/obras/obra_1/exportar", headers=cab("a")).json()
    ids = {importar(client, "a", arq).json()["id"] for _ in range(3)}
    assert len(ids) == 3 and "obra_1" not in ids
    assert sorted(nomes(client, "a")) == sorted(["Rede Norte", "Rede Norte (importada)", "Rede Norte (importada 2)", "Rede Norte (importada 3)"])


def test_recusa_projeto_diferente_do_selecionado_sem_gravar(client):
    arq = oa.montar_exportacao({"nome": "Obra PB", "projeto": "027", "dados_json": json.dumps(SNAP)})
    r = importar(client, "b", arq, projeto="229")
    assert r.status_code == 400 and "projeto 027, mas o projeto selecionado é 229" in r.json()["detail"]["erros"][0]
    assert nomes(client, "b", "229") == [] and nomes(client, "b", "027") == []
    assert importar(client, "b", arq, projeto=None).status_code == 400            # sem projeto selecionado
    assert importar(client, "b", arq, projeto="027").status_code == 200           # selecionando o projeto certo, entra


def test_recusa_projeto_nao_cadastrado(client):
    arq = oa.montar_exportacao({"nome": "X", "projeto": "999", "dados_json": json.dumps(SNAP)})
    r = importar(client, "b", arq, projeto="999")
    assert r.status_code == 400 and "não está cadastrado" in r.json()["detail"]["erros"][0]


def test_recusa_arquivo_invalido_grande_ou_estranho(client):
    assert importar(client, "b", "isto não é json").status_code == 400
    assert importar(client, "b", b"\xff\xfe\x00").status_code == 400
    assert importar(client, "b", {"formato": "x"}).status_code == 400
    grande = json.dumps(arquivo()) + " " * (oa.LIMITE_BYTES + 1)
    r = importar(client, "b", grande)
    assert r.status_code == 413 and "5 MB" in r.json()["detail"]["erros"][0]
    assert nomes(client, "b") == []


def test_qualquer_usuario_importa_inclusive_operador_e_admin(client):
    for uid, role in (("op", "operador"), ("adm", "admin")):
        r = client.post("/api/obras/importar?projeto=229", content=json.dumps(arquivo()), headers={**cab(uid, role), "Content-Type": "application/json"})
        assert r.status_code == 200, r.text


def test_texto_perigoso_do_arquivo_e_so_texto(client):
    carga = arquivo(nome="<img src=x onerror=alert(1)>", dados={"cabos": [], "outros": [{"entidade": "0", "operacao": "I", "ativo": "<script>alert(1)</script>"}]})
    r = importar(client, "b", carga)
    assert r.status_code == 200 and r.json()["nome"] == "<img src=x onerror=alert(1)>"          # guardado como texto; a tela usa textContent


# ── nuvem (cliente FALSO: os testes nunca tocam o Supabase real) ───────────────────────────────────────
class NuvemFalsa:
    def __init__(self):
        self.upserts, self.linhas = [], []

    def table(self, nome):
        assert nome == "obras"
        return self

    def upsert(self, payload):
        self.upserts.append(payload)
        self.linhas = [x for x in self.linhas if x["id"] != payload["id"]] + [payload]
        return self

    def select(self, *_):
        self._f = []
        return self

    def eq(self, k, v):
        self._f.append((k, v))
        return self

    def order(self, *_, **__):
        return self

    def execute(self):
        dados = [x for x in self.linhas if all(x.get(k) == v for k, v in getattr(self, "_f", []))]
        self._f = []
        return type("R", (), {"data": dados})()


def test_na_nuvem_grava_so_as_colunas_existentes_em_nome_do_importador(client, monkeypatch):
    nuvem = NuvemFalsa()
    monkeypatch.setattr(rota, "get_supabase", lambda: nuvem)
    salvar(client, "a")
    arq = client.get("/api/obras/obra_1/exportar", headers=cab("a")).json()
    r = importar(client, "b", arq)
    assert r.status_code == 200
    ultimo = nuvem.upserts[-1]
    assert set(ultimo) == {"id", "nome", "data", "dados_json", "user_id", "projeto", "publica", "dono_nome"} and ultimo["publica"] is False and ultimo["user_id"] == "b" and ultimo["id"] == r.json()["id"]
    assert [o["nome"] for o in client.get("/api/obras?projeto=229", headers=cab("b")).json()] == ["Rede Norte"]
