"""tests/test_prompts_validacao.py — prompts de validação (TASK-012): parsing, validação e rotas."""
import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from config import VALIDACOES_SEED_DIR
from middleware.auth_middleware import create_jwt_token
from services.prompts_validacao import ler_sementes, parse_prompt, validar_prompt

SEMENTES, _ = ler_sementes(VALIDACOES_SEED_DIR)

VALIDO = """---
id: validar-planilhas
escopo: ambos
modo: checar
ordem: 10
modelo: gemini-3.1-flash-lite
ativo: true
saida: json
placeholders: LINHAS_CABOS
---
Corpo com {{LINHAS_CABOS}}.
"""


def com(texto, antigo, novo):
    assert antigo in texto
    return texto.replace(antigo, novo, 1)


# ── serviço ──────────────────────────────────────────────────────────────────
def test_sementes_do_repositorio_sao_validas():
    assert set(SEMENTES) == {"validar-planilhas", "corrigir-planilhas", "ajustar-planilhas"}
    for pid, texto in SEMENTES.items():
        assert validar_prompt(texto, pid) == []


def test_parse_separa_cabecalho_e_corpo():
    meta, corpo = parse_prompt(VALIDO)
    assert meta["id"] == "validar-planilhas" and corpo == "Corpo com {{LINHAS_CABOS}}."


@pytest.mark.parametrize("texto,trecho", [
    ("sem cabeçalho", "começar com o cabeçalho"),
    ("---\nid: x\nsem fechar", "sem fechamento"),
    (com(VALIDO, "id: validar-planilhas", "id: outro-id"), "não pode ser alterado"),
    (com(VALIDO, "escopo: ambos", "escopo: tudo"), "escopo"),
    (com(VALIDO, "modo: checar", "modo: inventar"), "modo"),
    (com(VALIDO, "saida: json", "saida: texto"), "saida"),
    (com(VALIDO, "ordem: 10", "ordem: dez"), "ordem"),
    (com(VALIDO, "ativo: true", "ativo: talvez"), "ativo"),
    (com(VALIDO, "ativo: true\n", ""), "'ativo' é obrigatório"),
    (com(VALIDO, "saida: json", "saida: json\ntemperatura: 5"), "temperatura"),
    (com(VALIDO, "{{LINHAS_CABOS}}", "{{LINHAS_QUALQUER}}"), "não é um placeholder conhecido"),
    (com(VALIDO, "Corpo com {{LINHAS_CABOS}}.", "Corpo sem nada."), "não usado no corpo"),
    (com(VALIDO, "placeholders: LINHAS_CABOS", "placeholders: "), "não declarado"),
    (com(VALIDO, "Corpo com {{LINHAS_CABOS}}.", ""), "corpo do prompt está vazio"),
    (com(VALIDO, "Corpo com {{LINHAS_CABOS}}.", "Corpo {{LINHAS_CABOS}} e {{ quebrado"), "não forma um placeholder"),
])
def test_prompt_invalido(texto, trecho):
    erros = validar_prompt(texto, "validar-planilhas")
    assert any(trecho in e for e in erros), erros


# ── rotas (banco temporário, nunca o de desenvolvimento) ─────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(role):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


def test_lista_as_sementes_para_qualquer_usuario(client):
    r = client.get("/api/validacao/prompts", headers=cab("operador"))
    assert r.status_code == 200
    assert [p["prompt_id"] for p in r.json()["prompts"]] == ["validar-planilhas", "corrigir-planilhas", "ajustar-planilhas"]
    assert all(not p["personalizado"] for p in r.json()["prompts"])


def test_exige_autenticacao_e_escrita_so_admin(client):
    assert client.get("/api/validacao/prompts").status_code == 401
    corpo = {"projeto_codigo": "229", "conteudo": SEMENTES["validar-planilhas"]}
    assert client.post("/api/validacao/prompts/validar-planilhas", json=corpo, headers=cab("operador")).status_code == 403
    assert client.get("/api/validacao/prompts/validar-planilhas/historico", headers=cab("operador")).status_code == 403


def test_admin_salva_versao_do_projeto_sem_tocar_no_default(client):
    novo = SEMENTES["validar-planilhas"].replace("Sua tarefa é APENAS", "Sua tarefa (editada) é APENAS")
    r = client.post("/api/validacao/prompts/validar-planilhas",
                    json={"projeto_codigo": "229", "conteudo": novo}, headers=cab("admin"))
    assert r.status_code == 200
    do_projeto = client.get("/api/validacao/prompts/validar-planilhas?projeto_codigo=229", headers=cab("operador")).json()
    assert do_projeto["personalizado"] is True and "(editada)" in do_projeto["conteudo"]
    outro = client.get("/api/validacao/prompts/validar-planilhas?projeto_codigo=027", headers=cab("operador")).json()
    assert outro["personalizado"] is False and "(editada)" not in outro["conteudo"]


def test_prompt_invalido_e_rejeitado_com_lista_de_erros(client):
    ruim = SEMENTES["validar-planilhas"].replace("escopo: ambos", "escopo: tudo")
    r = client.post("/api/validacao/prompts/validar-planilhas",
                    json={"projeto_codigo": "229", "conteudo": ruim}, headers=cab("admin"))
    assert r.status_code == 400 and any("escopo" in e for e in r.json()["detail"]["erros"])
    assert client.get("/api/validacao/prompts/validar-planilhas?projeto_codigo=229",
                      headers=cab("admin")).json()["personalizado"] is False


def test_historico_e_reversao(client):
    v1 = SEMENTES["validar-planilhas"] + "\nVERSAO-1"
    v2 = SEMENTES["validar-planilhas"] + "\nVERSAO-2"
    for texto in (v1, v2):
        assert client.post("/api/validacao/prompts/validar-planilhas",
                           json={"projeto_codigo": "229", "conteudo": texto}, headers=cab("admin")).status_code == 200
    hist = client.get("/api/validacao/prompts/validar-planilhas/historico?projeto_codigo=229", headers=cab("admin")).json()["historico"]
    assert len(hist) == 1 and "VERSAO-1" in hist[0]["conteudo"] and hist[0]["criado_por"] == "admin@x.com"
    r = client.post("/api/validacao/prompts/validar-planilhas/reverter",
                    json={"projeto_codigo": "229", "historico_id": hist[0]["id"]}, headers=cab("admin"))
    assert r.status_code == 200
    atual = client.get("/api/validacao/prompts/validar-planilhas?projeto_codigo=229", headers=cab("admin")).json()
    assert "VERSAO-1" in atual["conteudo"] and "VERSAO-2" not in atual["conteudo"]
    # a reversão também preserva a versão que foi substituída
    hist = client.get("/api/validacao/prompts/validar-planilhas/historico?projeto_codigo=229", headers=cab("admin")).json()["historico"]
    assert any("VERSAO-2" in h["conteudo"] for h in hist)


def test_reverter_versao_inexistente_e_prompt_desconhecido(client):
    assert client.post("/api/validacao/prompts/validar-planilhas/reverter",
                       json={"projeto_codigo": "229", "historico_id": 999}, headers=cab("admin")).status_code == 404
    assert client.get("/api/validacao/prompts/nao-existe", headers=cab("admin")).status_code == 404
    assert client.post("/api/validacao/prompts/nao-existe",
                       json={"conteudo": VALIDO}, headers=cab("admin")).status_code == 404


def test_restaurar_semente(client):
    editado = SEMENTES["validar-planilhas"] + "\nEDITADO"
    client.post("/api/validacao/prompts/validar-planilhas",
                json={"projeto_codigo": "229", "conteudo": editado}, headers=cab("admin"))
    r = client.post("/api/validacao/prompts/validar-planilhas/restaurar-semente",
                    json={"projeto_codigo": "229"}, headers=cab("admin"))
    assert r.status_code == 200
    atual = client.get("/api/validacao/prompts/validar-planilhas?projeto_codigo=229", headers=cab("admin")).json()
    assert atual["conteudo"] == SEMENTES["validar-planilhas"]


def test_seed_e_idempotente_e_nao_sobrescreve_edicao_do_default(client):
    editado = SEMENTES["validar-planilhas"] + "\nEDITADO-DEFAULT"
    assert client.post("/api/validacao/prompts/validar-planilhas",
                       json={"projeto_codigo": "DEFAULT", "conteudo": editado}, headers=cab("admin")).status_code == 200
    database.init_db()  # segunda inicialização
    atual = client.get("/api/validacao/prompts/validar-planilhas", headers=cab("admin")).json()
    assert "EDITADO-DEFAULT" in atual["conteudo"]
    assert len(client.get("/api/validacao/prompts", headers=cab("admin")).json()["prompts"]) == 3
