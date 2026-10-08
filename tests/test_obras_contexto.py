"""tests/test_obras_contexto.py — IA do chat consulta e comanda as obras salvas (TASK-028).

Decisões do usuário: obra referida pelo NOME; só obras do projeto selecionado; confirmação sempre em popup (frontend);
a IA não salva nem exclui obras; sem regressão e sem lentidão (conversas sem a palavra "obra" ficam idênticas)."""
import json
import types

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.ai_chat as chat
from middleware.auth_middleware import create_jwt_token
from services.obras_contexto import (contar_obra, menciona_obra, montar_indice, normalizar, obras_citadas, resumir_obra)


def obra(i, nome, **kw):
    return {"id": f"obra_{i}", "nome": nome, "data": f"0{i}/10/2026", "projeto": "229", **kw}


def snap(cabos=(), outros=()):
    return json.dumps({"cabos": {"data": [{"operacao": "I", "ativo": a} for a in cabos]},
                       "outros": {"data": [{"operacao": "I", "ativo": a} for a in outros]}})


# ── funções puras ────────────────────────────────────────────────────────────
@pytest.mark.parametrize("texto,esperado", [
    ("adicione a obra 2 ao projeto", True), ("Quais OBRAS eu tenho?", True), ("carregue a Obra Rede Norte", True),
    ("gere uma rede com 3 postes", False), ("adicionar 2 CFU", False), ("faça a manobra", False), ("", False),
])
def test_menciona_obra(texto, esperado):
    assert menciona_obra(texto) is esperado


def test_indice_enxuto_nao_traz_conteudo_e_tem_teto():
    obras = [obra(i, f"Obra {i}", dados_json=snap(outros=["1-SECRETO"])) for i in range(1, 61)]
    ind = montar_indice(obras)
    assert "SECRETO" not in ind and 'nome="Obra 1"' in ind and "id=obra_50" in ind and "id=obra_51" not in ind
    assert "+10 obras mais antigas" in ind
    assert "nenhuma obra salva" in montar_indice([])


def test_busca_por_nome_sem_acento_e_so_quando_pede_conteudo():
    obras = [obra(1, "Obra 2"), obra(2, "Obra 22"), obra(3, "Rede Norte")]
    assert [o["id"] for o in obras_citadas("o que tem na OBRA 22?", obras)] == ["obra_2", "obra_1"]   # nome mais específico primeiro
    assert [o["id"] for o in obras_citadas("mostre a rede norte", obras)] == ["obra_3"]
    assert obras_citadas("adicione a obra 2 ao projeto", obras) == []        # comando: o conteúdo não precisa ir
    assert obras_citadas("o que tem na obra sul", obras) == []
    assert len(obras_citadas("compare obra 2 com obra 22 e rede norte", obras)) == 2   # teto
    assert normalizar("  Obra  Ação ") == "obra acao"


def test_resumo_de_uma_obra_com_teto_de_linhas_e_dados_ilegiveis():
    r = resumir_obra(obra(1, "Obra 2", dados_json=snap(cabos=["CAA 2 A 3 m"], outros=[f"1-X{i}" for i in range(45)])))
    assert 'CONTEÚDO DA OBRA "Obra 2"' in r and "CABOS (1 linhas)" in r and "OUTROS (45 linhas)" in r and "(+5 linhas não mostradas)" in r
    assert "1-X39" in r and "1-X40" not in r
    assert "CABOS (0 linhas)" in resumir_obra(obra(1, "X", dados_json="isto não é json"))
    assert contar_obra(snap(["A"], ["B", "C"])) == (1, 2) and contar_obra("lixo") == (0, 0) and contar_obra("[]") == (0, 0)


# ── rotas ────────────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    database.init_db()
    return TestClient(appmod.app)


def cab(uid="u1", role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@x.com', role)}"}


def salvar(client, uid, i, nome, projeto="229", **kw):
    corpo = {"id": f"obra_{i}", "nome": nome, "data": f"0{i}/10/2026", "dados_json": kw.get("dados", snap(["CAA 2 A 3 m"], ["1-SUPL"])), "projeto": projeto}
    assert client.post("/api/obras", json=corpo, headers=cab(uid)).status_code == 200


def test_indice_so_do_usuario_e_do_projeto_e_sem_dados_json(client):
    salvar(client, "u1", 1, "Obra 2"); salvar(client, "u1", 2, "Obra 3", projeto="027"); salvar(client, "u2", 3, "De outro usuário")
    r = client.get("/api/obras/indice?projeto=229", headers=cab("u1")).json()
    assert [o["nome"] for o in r] == ["Obra 2"] and set(r[0]) == {"id", "nome", "data", "projeto", "publica", "minha", "sem_dono", "dono"}
    assert [o["nome"] for o in client.get("/api/obras/indice?projeto=027", headers=cab("u1")).json()] == ["Obra 3"]
    assert client.get("/api/obras/indice?projeto=229").status_code == 401


def test_buscar_uma_obra_valida_o_dono(client):
    salvar(client, "u1", 1, "Obra 2")
    ok = client.get("/api/obras/obra_1", headers=cab("u1")).json()
    assert ok["nome"] == "Obra 2" and "dados_json" in ok
    assert client.get("/api/obras/obra_1", headers=cab("u2")).status_code == 404          # não é dele
    assert client.get("/api/obras/obra_999", headers=cab("u1")).status_code == 404         # inexistente (a IA não inventa obra)


def req_falso(prompt, projeto="229", uid="u1"):
    return types.SimpleNamespace(prompt=prompt, projeto_codigo=projeto), types.SimpleNamespace(state=types.SimpleNamespace(user={"user_id": uid}))


def test_sem_a_palavra_obra_nao_ha_contexto_nem_consulta_ao_banco(monkeypatch):
    def proibido(*a, **k):
        raise AssertionError("não devia consultar obras")
    monkeypatch.setattr(chat, "listar_visiveis", proibido); monkeypatch.setattr(chat, "buscar_obra", proibido)
    for prompt in ("gere uma rede com 3 postes", "adicionar 2 CFU", "analise a tabela", "ordene os cabos"):
        assert chat._contexto_obras(*req_falso(prompt)) == ("", ""), prompt


def test_com_a_palavra_obra_vai_indice_e_instrucoes_e_conteudo_so_se_pedido(client):
    salvar(client, "u1", 1, "Obra 2", dados=snap(outros=["1-CFU-ESPECIAL"])); salvar(client, "u1", 2, "Rede Norte")
    extra, instrucoes = chat._contexto_obras(*req_falso("quais obras eu tenho?"))
    assert 'nome="Obra 2"' in extra and 'nome="Rede Norte"' in extra and "CONTEÚDO DA OBRA" not in extra
    assert '"acao_ui": "obra"' in instrucoes and "NÃO exclui nem salva" in instrucoes
    extra, _ = chat._contexto_obras(*req_falso("o que tem na obra 2?"))
    assert 'CONTEÚDO DA OBRA "Obra 2"' in extra and "1-CFU-ESPECIAL" in extra and 'CONTEÚDO DA OBRA "Rede Norte"' not in extra
    extra, _ = chat._contexto_obras(*req_falso("quais obras eu tenho?", projeto="027"))
    assert "nenhuma obra salva" in extra                                                    # só o projeto selecionado
    extra, _ = chat._contexto_obras(*req_falso("quais obras eu tenho?", uid="u2"))
    assert "nenhuma obra salva" in extra                                                    # só do próprio usuário


def test_falha_ao_montar_o_contexto_nao_derruba_o_chat(monkeypatch):
    monkeypatch.setattr(chat, "listar_visiveis", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("banco fora")))
    assert chat._contexto_obras(*req_falso("quais obras eu tenho?")) == ("", "")


def test_fast_path_continua_igual_e_nao_intercepta_frases_com_obra(client, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    corpo = lambda p: {"prompt": p, "history": [], "projeto_codigo": "229"}
    r = client.post("/api/gemini/chat", json=corpo("adicionar 2 CFU"), headers=cab())
    assert r.status_code == 200 and "(Fast-Path)" in r.text and "2-CFU" in r.text

    class IASentinela:                                   # chegar aqui prova que a frase foi para a IA (sem rede nos testes)
        def __init__(self, *a, **k):
            raise RuntimeError("alcançou a IA")
    import google.genai as genai
    monkeypatch.setattr(genai, "Client", IASentinela)
    with pytest.raises(RuntimeError, match="alcançou a IA"):
        client.post("/api/gemini/chat", json=corpo("adicionar 2 obra norte"), headers=cab())
