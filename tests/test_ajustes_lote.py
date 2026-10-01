"""tests/test_ajustes_lote.py — Ajustar roda todos os ajustes habilitados (TASK-029).

Decisões do usuário: 'Executar todas' em cadeia na ordem do cadastro; ajustes que não mudam nada ficam fora dos cartões;
confirmação extra para exclusões; só simula (o frontend aplica)."""
import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.validacao_ajustes as rota
from middleware.auth_middleware import create_jwt_token


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


def receita(rid, acoes, ativa=True, nome=None):
    return {"id": rid, "nome": nome or rid, "ativa": ativa, "regras": [], "acoes": acoes}


def padrao(client, receitas):
    r = client.post("/api/validacao/ajustes", json={"projeto_codigo": "DEFAULT", "ajustes": receitas}, headers=cab("admin"))
    assert r.status_code == 200, r.text


def lote(client, outros, cabos=(), projeto=None):
    corpo = {"outros": [{"operacao": "I", "ativo": a} for a in outros], "cabos": list(cabos)}
    if projeto:
        corpo["projeto_codigo"] = projeto
    r = client.post("/api/validacao/ajustes/preview-lote", json=corpo, headers=cab())
    assert r.status_code == 200, r.text
    return r.json()


CFUU = {"acao": "substituir", "tabela": "outros", "de": "1-CFUU", "para": "1-CFU"}
ADD_SUPL = {"acao": "adicionar_ativo", "tabela": "outros", "ativo": "SUPL", "quando": {"todos": [{"tem": "CFU"}, {"nao": {"tem": "SUPL"}}]}}
VAZIAS = {"acao": "excluir_linhas", "tabela": "outros", "onde": {"vazias": True}}


def test_exige_login(client):
    assert client.post("/api/validacao/ajustes/preview-lote", json={}).status_code == 401


def test_sem_ajuste_habilitado_devolve_tudo_vazio(client):
    padrao(client, [receita("A", [CFUU], ativa=False)])
    r = lote(client, ["DT11/300 1-CFUU"])
    assert r == {"total_habilitados": 0, "itens": [], "sem_mudanca": [], "problemas": [], "cadeia": None, "cadeia_erro": None, "avisos": []}


def test_cartao_so_para_quem_muda_e_os_outros_vao_para_sem_mudanca(client):
    padrao(client, [receita("A", [CFUU], nome="Corrige CFUU"), receita("B", [VAZIAS], nome="Vazias")])
    r = lote(client, ["DT11/300 1-CFUU", "1-X"])
    assert r["total_habilitados"] == 2 and [i["id"] for i in r["itens"]] == ["A"] and [s["id"] for s in r["sem_mudanca"]] == ["B"]
    a = r["itens"][0]
    assert a["nome"] == "Corrige CFUU" and a["frases"] == ["Em Outros: substituir '1-CFUU' por '1-CFU'."] and a["destrutivo"] is False
    assert a["diff"]["resumo"]["editar"] == 1 and a["diff"]["operacoes"][0]["depois"]["ativo"] == "DT11/300 1-CFU"


def test_cadeia_e_igual_a_executar_em_sequencia_e_difere_dos_diffs_independentes(client):
    padrao(client, [receita("A", [CFUU]), receita("B", [ADD_SUPL])])
    r = lote(client, ["DT11/300 1-CFUU"])
    assert [i["id"] for i in r["itens"]] == ["A"] and [s["id"] for s in r["sem_mudanca"]] == ["B"]     # sozinho, B não muda nada
    cadeia = r["cadeia"]
    assert cadeia["outros"][0]["ativo"] == "DT11/300 1-CFU 1-SUPL"                                    # em cadeia, B enxerga o resultado de A
    manual = client.post("/api/validacao/ajustes/preview", headers=cab(), json={
        "receitas": ["A", "B"], "outros": [{"operacao": "I", "ativo": "DT11/300 1-CFUU"}]}).json()
    assert cadeia["operacoes"] == [dict(o, acoes=o["acoes"]) for o in manual["operacoes"]] and cadeia["destrutivo"] is False


def test_destrutivo_por_cartao_e_na_cadeia(client):
    padrao(client, [receita("A", [CFUU]), receita("B", [VAZIAS]), receita("C", [{"acao": "remover_ativo", "tabela": "outros", "ativo": "X"}])])
    r = lote(client, ["DT11/300 1-CFUU", "", "1-X 1-Y"])
    por_id = {i["id"]: i for i in r["itens"]}
    assert por_id["A"]["destrutivo"] is False and por_id["B"]["destrutivo"] is True and por_id["C"]["destrutivo"] is True
    assert r["cadeia"]["destrutivo"] is True and r["cadeia"]["resumo"]["excluir"] == 1


def test_oculto_no_projeto_e_desligado_ficam_de_fora_e_projeto_usa_o_seu_conjunto(client):
    padrao(client, [receita("A", [CFUU]), receita("B", [VAZIAS], ativa=False)])
    lista = client.get("/api/validacao/ajustes?projeto_codigo=229", headers=cab("admin")).json()["ajustes"]
    for a in lista:
        if a["id"] == "A":
            a["oculta"] = True
    client.post("/api/validacao/ajustes", json={"projeto_codigo": "229", "ajustes": lista}, headers=cab("admin"))
    assert lote(client, ["DT11/300 1-CFUU"], projeto="229")["total_habilitados"] == 0      # A oculto no 229, B desligado
    assert lote(client, ["DT11/300 1-CFUU"], projeto="027")["total_habilitados"] == 1      # o 027 herda o padrão


def test_guarda_da_camada_1_aparece_como_descartada_e_nao_vira_cartao(client):
    ruim = {"acao": "substituir", "tabela": "outros", "de": "CFU", "para": "CFU EXTRA"}         # deixaria token sem par
    padrao(client, [receita("R", [ruim])])
    r = lote(client, ["DT11/300 1-CFU"])
    assert r["itens"] == [] and r["sem_mudanca"][0]["descartadas"][0]["motivo"].startswith("o ajuste deixaria a linha inválida")
    assert r["cadeia"] is None


def test_ajuste_invalido_vira_problema_sem_derrubar_os_outros(client, monkeypatch):
    bom = receita("A", [CFUU])
    quebrado = receita("Q", [{"acao": "fazer_magica", "tabela": "outros"}])
    monkeypatch.setattr(rota, "ajustes_efetivos", lambda projeto: [quebrado, bom])
    r = lote(client, ["DT11/300 1-CFUU"])
    assert [p["id"] for p in r["problemas"]] == ["Q"] and any("'acao' precisa ser" in e for e in r["problemas"][0]["erros"])
    assert [i["id"] for i in r["itens"]] == ["A"] and r["cadeia"]["resumo"]["editar"] == 1


def test_muitas_acoes_desabilitam_a_cadeia_mas_mantem_os_cartoes(client):
    acoes = [{"acao": "substituir", "tabela": "outros", "de": f"X{i}", "para": f"Y{i}"} for i in range(26)]
    padrao(client, [receita("A", acoes[:26]), receita("B", acoes[:26]), receita("C", [CFUU])])
    r = lote(client, ["DT11/300 1-CFUU 1-X1"])
    assert r["cadeia"] is None and "uma a uma" in r["cadeia_erro"] and [i["id"] for i in r["itens"]] == ["A", "B", "C"]


def test_limite_de_ajustes_por_pedido(client, monkeypatch):
    monkeypatch.setattr(rota, "ajustes_efetivos", lambda projeto: [receita(f"R{i}", [CFUU]) for i in range(35)])
    r = lote(client, ["DT11/300 1-CFUU"])
    assert r["total_habilitados"] == 30 and "30 primeiros" in r["avisos"][0]
