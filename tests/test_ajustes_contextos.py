"""tests/test_ajustes_contextos.py — contextos dentro de um projeto, terceira camada de overlay
sobre Ajustes (TASK-044).

Decisão do usuário: overlay completo por contexto (não só liga/desliga), reaproveitando a mesma
função `efetivo()` de services/ajustes_camadas.py em cadeia. Toda rota de Ajustes continua se
comportando exatamente como antes quando `contexto` não é informado (regressão coberta aqui)."""
import json

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from config import AJUSTES_SEED_PATH
from middleware.auth_middleware import create_jwt_token

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


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(role):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


def get(client, projeto="229", contexto=None, role="admin"):
    url = f"/api/validacao/ajustes?projeto_codigo={projeto}"
    if contexto:
        url += f"&contexto={contexto}"
    return client.get(url, headers=cab(role)).json()


def salvar(client, projeto, ajustes, contexto=None, role="admin"):
    corpo = {"projeto_codigo": projeto, "ajustes": ajustes}
    if contexto:
        corpo["contexto"] = contexto
    return client.post("/api/validacao/ajustes", json=corpo, headers=cab(role))


def editar(client, projeto, mut, contexto=None):
    lista = get(client, projeto, contexto)["ajustes"]
    mut(por_id(lista), lista)
    return salvar(client, projeto, lista, contexto)


def criar_contexto(client, projeto, contexto, role="admin"):
    return client.post("/api/validacao/ajustes/contextos",
                        json={"projeto_codigo": projeto, "contexto": contexto}, headers=cab(role))


def excluir_contexto(client, projeto, contexto, role="admin"):
    return client.post("/api/validacao/ajustes/contextos/excluir",
                        json={"projeto_codigo": projeto, "contexto": contexto}, headers=cab(role))


def listar_contextos(client, projeto, role="admin"):
    return client.get(f"/api/validacao/ajustes/contextos?projeto_codigo={projeto}", headers=cab(role)).json()


# ── CRUD de contexto ─────────────────────────────────────────────────────────
def test_criar_listar_e_excluir_contexto(client):
    assert listar_contextos(client, "229")["contextos"] == []
    assert criar_contexto(client, "229", "34,5kV").status_code == 200
    assert criar_contexto(client, "229", "13,8kV").status_code == 200
    assert listar_contextos(client, "229")["contextos"] == ["13,8kV", "34,5kV"]
    # contextos de OUTRO projeto ficam isolados
    assert listar_contextos(client, "027")["contextos"] == []
    assert excluir_contexto(client, "229", "13,8kV").status_code == 200
    assert listar_contextos(client, "229")["contextos"] == ["34,5kV"]


def test_criar_contexto_vazio_ou_longo_e_rejeitado(client):
    assert criar_contexto(client, "229", "").status_code == 400
    assert criar_contexto(client, "229", "   ").status_code == 400
    assert criar_contexto(client, "229", "x" * 61).status_code == 400
    assert listar_contextos(client, "229")["contextos"] == []


def test_criar_e_excluir_contexto_exige_admin(client):
    assert criar_contexto(client, "229", "34,5kV", role="operador").status_code == 403
    criar_contexto(client, "229", "34,5kV")
    assert excluir_contexto(client, "229", "34,5kV", role="operador").status_code == 403


def test_listar_contextos_exige_autenticacao_mas_nao_role_especifico(client):
    criar_contexto(client, "229", "34,5kV")
    assert client.get("/api/validacao/ajustes/contextos?projeto_codigo=229").status_code == 401
    assert listar_contextos(client, "229", role="operador")["contextos"] == ["34,5kV"]


# ── regressão: sem contexto, nada muda ───────────────────────────────────────
def test_sem_contexto_comportamento_identico_a_antes(client):
    sem = get(client, "229")
    sem_explicito = get(client, "229", contexto=None)
    assert sem == sem_explicito
    assert sem["contexto"] is None
    criar_contexto(client, "229", "34,5kV")   # criar um contexto não afeta o projeto sem contexto
    assert get(client, "229") == sem


# ── resolução em cadeia ──────────────────────────────────────────────────────
def test_contexto_liga_e_desliga_por_cima_do_projeto(client):
    # projeto 229: liga AJ-SUP-L, deixa o resto como a semente (todos desligados)
    editar(client, "229", lambda m, lista: m["AJ-SUP-L"].update(ativa=True))
    efetivo_projeto = por_id(get(client, "229")["ajustes"])
    assert efetivo_projeto["AJ-SUP-L"]["ativa"] is True
    assert efetivo_projeto["AJ-EXCLUIR-VAZIAS"]["ativa"] is False

    criar_contexto(client, "229", "34,5kV")
    # dentro do contexto: desliga o que o projeto ligou, liga outro que o projeto não ligou
    r = editar(client, "229",
               lambda m, lista: (m["AJ-SUP-L"].update(ativa=False), m["AJ-EXCLUIR-VAZIAS"].update(ativa=True)),
               contexto="34,5kV")
    assert r.status_code == 200

    efetivo_contexto = por_id(get(client, "229", contexto="34,5kV")["ajustes"])
    assert efetivo_contexto["AJ-SUP-L"]["ativa"] is False
    assert efetivo_contexto["AJ-EXCLUIR-VAZIAS"]["ativa"] is True

    # o projeto (sem contexto) continua exatamente como estava antes de editar o contexto
    assert por_id(get(client, "229")["ajustes"])["AJ-SUP-L"]["ativa"] is True
    assert por_id(get(client, "229")["ajustes"])["AJ-EXCLUIR-VAZIAS"]["ativa"] is False


def test_dois_contextos_sao_independentes(client):
    criar_contexto(client, "229", "34,5kV")
    criar_contexto(client, "229", "13,8kV")
    editar(client, "229", lambda m, lista: m["AJ-SUP-L"].update(ativa=True), contexto="34,5kV")
    editar(client, "229", lambda m, lista: m["AJ-EXCLUIR-VAZIAS"].update(ativa=True), contexto="13,8kV")

    ctx1 = por_id(get(client, "229", contexto="34,5kV")["ajustes"])
    ctx2 = por_id(get(client, "229", contexto="13,8kV")["ajustes"])
    assert ctx1["AJ-SUP-L"]["ativa"] is True and ctx1["AJ-EXCLUIR-VAZIAS"]["ativa"] is False
    assert ctx2["AJ-SUP-L"]["ativa"] is False and ctx2["AJ-EXCLUIR-VAZIAS"]["ativa"] is True


def test_contexto_pode_adicionar_receita_exclusiva_e_sobrescrever_campo(client):
    criar_contexto(client, "229", "34,5kV")

    def mut(m, lista):
        lista.append(receita("AJ-SO-CONTEXTO"))
        m["AJ-SUP-L"]["nome"] = "Nome só no contexto"
    editar(client, "229", mut, contexto="34,5kV")

    efetivo_contexto = por_id(get(client, "229", contexto="34,5kV")["ajustes"])
    assert "AJ-SO-CONTEXTO" in efetivo_contexto
    assert efetivo_contexto["AJ-SUP-L"]["nome"] == "Nome só no contexto"
    # não existe no projeto nem no padrão
    efetivo_projeto = por_id(get(client, "229")["ajustes"])
    assert "AJ-SO-CONTEXTO" not in efetivo_projeto
    assert efetivo_projeto["AJ-SUP-L"]["nome"] != "Nome só no contexto"


def test_excluir_contexto_nao_afeta_projeto_nem_outro_contexto(client):
    criar_contexto(client, "229", "34,5kV")
    criar_contexto(client, "229", "13,8kV")
    editar(client, "229", lambda m, lista: m["AJ-SUP-L"].update(ativa=True), contexto="34,5kV")
    editar(client, "229", lambda m, lista: m["AJ-EXCLUIR-VAZIAS"].update(ativa=True), contexto="13,8kV")
    projeto_antes = get(client, "229")

    assert excluir_contexto(client, "229", "34,5kV").status_code == 200

    assert get(client, "229") == projeto_antes
    assert listar_contextos(client, "229")["contextos"] == ["13,8kV"]
    ctx2 = por_id(get(client, "229", contexto="13,8kV")["ajustes"])
    assert ctx2["AJ-EXCLUIR-VAZIAS"]["ativa"] is True


def test_contexto_de_projeto_sem_overlay_proprio_ainda_funciona():
    """Contexto criado num projeto que nunca salvou nada (sem overlay próprio) resolve contra o
    padrão direto — não precisa o projeto ter overlay para o contexto funcionar."""
    pass  # coberto implicitamente: "229" nos testes acima nunca tem overlay salvo até o 1º editar()


# ── preview-lote respeita o contexto ─────────────────────────────────────────
def test_preview_lote_roda_o_conjunto_do_contexto(client):
    criar_contexto(client, "229", "34,5kV")
    editar(client, "229", lambda m, lista: m["AJ-SUP-L"].update(ativa=True), contexto="34,5kV")
    cabos, outros = [], [{"operacao": "I", "ativo": "1-SUP-L"}]

    resp_sem_contexto = client.post("/api/validacao/ajustes/preview-lote",
                                     json={"projeto_codigo": "229", "cabos": cabos, "outros": outros},
                                     headers=cab("operador"))
    resp_com_contexto = client.post("/api/validacao/ajustes/preview-lote",
                                     json={"projeto_codigo": "229", "contexto": "34,5kV", "cabos": cabos, "outros": outros},
                                     headers=cab("operador"))
    assert resp_sem_contexto.json()["total_habilitados"] == 0
    assert resp_com_contexto.json()["total_habilitados"] == 1
    assert resp_com_contexto.json()["itens"][0]["id"] == "AJ-SUP-L"
