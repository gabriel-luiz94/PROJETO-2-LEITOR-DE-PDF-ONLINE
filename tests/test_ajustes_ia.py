"""tests/test_ajustes_ia.py — ajustes por IA com ações estruturadas (TASK-025, ADR-006).

O modelo é sempre simulado: nenhum teste faz chamada de rede nem usa chave real.
Decisão do usuário: a IA pode propor EXCLUSÃO de linhas, sempre com aviso e pré-visualização."""
import asyncio
import json

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.validacao as rota
from config import VALIDACOES_SEED_DIR
from middleware.auth_middleware import create_jwt_token
from services.correcao_ia import ajustar_com_ia, interpretar_acoes
from services.prompts_validacao import ler_sementes, validar_prompt

PROMPT = ler_sementes(VALIDACOES_SEED_DIR)[0]["ajustar-planilhas"]
GRUPOS = {"CHAVES": ["CFU", "CFUR"]}


def L(ativo, op="I"):
    return {"ativo": ativo, "operacao": op}


def resp(*acoes):
    return json.dumps({"acoes": list(acoes)})


def a(**kw):
    base = {"acao": "substituir", "tabela": "outros", "de": "1-CFUU", "para": "1-CFU", "motivo": "digitação de CFU"}
    base.update(kw)
    return base


ACHADOS = [{"linha_id": "OUTROS-0", "regra_id": "IA:digitacao", "mensagem": "parece CFU"}]


def rodar(coro):
    return asyncio.run(coro)


def com_resposta(texto, outros, achados=ACHADOS, cabos=(), grupos=None):
    async def chamar(_):
        return texto
    return rodar(ajustar_com_ia(chamar, PROMPT, list(cabos), outros, achados, grupos))


# ── semente do prompt ────────────────────────────────────────────────────────
def test_prompt_semente_e_valido_e_proibe_mudar_a_operacao():
    assert validar_prompt(PROMPT, "ajustar-planilhas") == []
    assert "NUNCA mude a operação" in PROMPT and "DESTRUTIVA" in PROMPT


# ── interpretação ────────────────────────────────────────────────────────────
def test_interpretar_acoes_tolera_cerca_e_separa_o_motivo():
    propostas, desc = interpretar_acoes("```json\n" + resp(a()) + "\n```")
    assert propostas == [{"acao": {"acao": "substituir", "tabela": "outros", "de": "1-CFUU", "para": "1-CFU"}, "motivo": "digitação de CFU"}] and desc == []
    propostas, desc = interpretar_acoes(resp("solto", a(motivo=5)))
    assert len(desc) == 1 and propostas[0]["motivo"] == ""


@pytest.mark.parametrize("texto", ["não é json", "[]", '{"correcoes": []}', '{"acoes": "x"}', ""])
def test_resposta_ilegivel_levanta(texto):
    with pytest.raises(ValueError):
        interpretar_acoes(texto)


# ── ajustar_com_ia ───────────────────────────────────────────────────────────
def test_proposta_valida_vira_diff_calculado_pelo_motor_sem_aplicar_nada():
    outros = [L("DT11/300 1-CFUU"), L("1-OUTRA")]
    r = com_resposta(resp(a(de="CFUU", para="CFU")), outros)
    assert r["status"] == "ok" and r["destrutivo"] is False and r["descartadas"] == []
    assert r["acoes"][0]["frase"] == "Em Outros: substituir 'CFUU' por 'CFU'." and r["acoes"][0]["motivo"] == "digitação de CFU"
    assert r["diff"]["resumo"]["editar"] == 1 and r["diff"]["operacoes"][0]["depois"]["ativo"] == "DT11/300 1-CFU"
    assert outros == [L("DT11/300 1-CFUU"), L("1-OUTRA")]                   # entrada intacta


def test_prompt_recebe_so_a_linha_citada():
    enviados = []

    async def chamar(texto):
        enviados.append(texto)
        return resp()

    r = rodar(ajustar_com_ia(chamar, PROMPT, [L("P50 A 2 m")], [L("DT11/300 1-CFUU"), L("1-OUTRA")], ACHADOS))
    assert r["status"] == "ok" and r["acoes"] == [] and r["diff"] is None
    assert "[OUTROS-0] | I | DT11/300 1-CFUU" in enviados[0] and "1-OUTRA" not in enviados[0] and "{{" not in enviados[0]


def test_exclusao_e_marcada_como_destrutiva_e_o_diff_traz_a_exclusao():
    outros = [L("DT11/300 1-CFUU"), L(""), L("")]
    r = com_resposta(resp({"acao": "excluir_linhas", "tabela": "outros", "onde": {"vazias": True}, "motivo": "linhas vazias"}), outros)
    assert r["destrutivo"] is True and r["acoes"][0]["destrutiva"] is True
    assert r["diff"]["resumo"]["excluir"] == 2 and r["diff"]["outros"][0]["ativo"] == "DT11/300 1-CFUU"


def test_remover_ativo_tambem_e_destrutivo_mesmo_sem_excluir_linha():
    r = com_resposta(resp({"acao": "remover_ativo", "tabela": "outros", "ativo": "CFUU"}), [L("DT11/300 1-CFUU 1-X")])
    assert r["destrutivo"] is True and r["diff"]["resumo"] == {"editar": 1, "inserir": 0, "excluir": 0, "mover": 0}


def test_acao_nao_destrutiva_nao_avisa():
    r = com_resposta(resp({"acao": "ordenar", "tabela": "outros", "por": [{"coluna": "ativo"}]}), [L("1-B"), L("1-A")])
    assert r["destrutivo"] is False and r["diff"]["resumo"]["mover"] == 1


@pytest.mark.parametrize("acao,trecho", [
    ({"acao": "substituir", "tabela": "outros", "de": "A", "para": "B", "operacao_nova": "R"}, "não pode mudar a operação"),
    ({"acao": "fazer_magica", "tabela": "outros"}, "'acao' precisa ser"),
    ({"acao": "substituir", "tabela": "outros", "de": {"regex": "(["}, "para": "x"}, "regex inválida"),
    ({"acao": "remover_ativo", "tabela": "outros", "ativo": "@NAO_EXISTE"}, "não existe"),
    ({"acao": "excluir_linhas", "tabela": "outros", "onde": {}}, "exatamente um"),
])
def test_acao_invalida_ou_proibida_e_descartada_com_motivo(acao, trecho):
    r = com_resposta(resp(acao), [L("DT11/300 1-CFUU")])
    assert r["diff"] is None and r["acoes"] == [] and any(trecho in d["motivo"] for d in r["descartadas"]), r["descartadas"]


def test_grupos_do_projeto_valem_para_a_acao_da_ia():
    r = com_resposta(resp({"acao": "remover_ativo", "tabela": "outros", "ativo": "@CHAVES"}), [L("DT11/300 1-CFU 1-X")], grupos=GRUPOS)
    assert r["diff"]["outros"][0]["ativo"] == "DT11/300 1-X"


def test_ajuste_que_deixa_a_linha_invalida_e_barrado_pela_camada_1_no_diff():
    r = com_resposta(resp(a(de="CFUU", para="CFU EXTRA")), [L("DT11/300 1-CFUU")])
    assert r["acoes"] and r["diff"]["operacoes"] == [] and "C1-OUT-ORFAO" in r["diff"]["descartadas"][0]["motivo"]


def test_limite_de_acoes_e_deduplicacao():
    acoes = [a(de=f"X{i}", para=f"Y{i}") for i in range(12)]
    r = com_resposta(resp(*acoes, a(de="X0", para="Y0")), [L("DT11/300 1-CFUU")])
    assert len(r["acoes"]) == 10 and sum("mais de 10" in d["motivo"] for d in r["descartadas"]) == 2


def test_falha_do_modelo_e_json_quebrado_viram_status():
    async def quebra(_):
        raise RuntimeError("429 quota exhausted")
    r = rodar(ajustar_com_ia(quebra, PROMPT, [], [L("DT11/300 1-CFUU")], ACHADOS))
    assert r["status"] == "erro" and "429" in r["mensagem"] and r["acoes"] == [] and r["diff"] is None
    r = com_resposta("```isso nao e json```", [L("DT11/300 1-CFUU")])
    assert r["status"] == "erro" and "JSON" in r["mensagem"]


def test_sem_linhas_citadas_nao_chama_o_modelo():
    async def chamar(_):
        raise AssertionError("não devia chamar")
    r = rodar(ajustar_com_ia(chamar, PROMPT, [], [L("1-A")], [{"linha_id": "GERAL"}]))
    assert r["status"] == "ok" and r["lotes"] == 0 and r["acoes"] == []


def test_tabela_muito_grande_vira_erro_e_nao_levanta():
    r = com_resposta(resp(a(de="CFUU", para="CFU")), [L("1-A")] * 5001, achados=[{"linha_id": "OUTROS-0", "regra_id": "x", "mensagem": "m"}])
    assert r["status"] == "erro" and "grandes demais" in r["mensagem"]


# ── rota ─────────────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    database.init_db()
    return TestClient(appmod.app)


@pytest.fixture
def modelo(monkeypatch):
    class Falso:
        resposta = resp(a(de="CFUU", para="CFU"))
        chamadas = []

    async def falso(api_key, modelos, texto, temperatura=0.0):
        Falso.chamadas.append(texto)
        return Falso.resposta

    Falso.chamadas = []
    monkeypatch.setattr(rota, "chamar_gemini", falso)
    return Falso


def cab(role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


CORPO = {"cabos": [], "outros": [L("DT11/300 1-CFUU")], "achados": ACHADOS, "projeto_codigo": "229"}


def test_rota_exige_login_e_chave(client, modelo):
    assert client.post("/api/validacao/ajustes-ia", json=CORPO).status_code == 401
    r = client.post("/api/validacao/ajustes-ia", json=CORPO, headers=cab())
    assert r.status_code == 200 and r.json()["status"] == "sem_chave" and modelo.chamadas == []


def test_rota_devolve_acoes_e_diff_sem_aplicar_nada(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    r = client.post("/api/validacao/ajustes-ia", json=CORPO, headers=cab()).json()
    assert r["status"] == "ok" and r["acoes"][0]["acao"]["acao"] == "substituir" and r["diff"]["outros"][0]["ativo"] == "DT11/300 1-CFU"
    assert "[OUTROS-0] | I | DT11/300 1-CFUU" in modelo.chamadas[0]


def test_rota_usa_os_grupos_do_projeto_e_respeita_prompt_desativado(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    modelo.resposta = resp({"acao": "remover_ativo", "tabela": "outros", "ativo": "@CHAVE_MT"})   # grupo da semente das regras
    corpo = {**CORPO, "outros": [L("DT11/300 1-CFU 1-X")]}
    assert client.post("/api/validacao/ajustes-ia", json=corpo, headers=cab()).json()["diff"]["outros"][0]["ativo"] == "DT11/300 1-X"
    desligado = PROMPT.replace("ativo: true", "ativo: false")
    client.post("/api/validacao/prompts/ajustar-planilhas", json={"projeto_codigo": "DEFAULT", "conteudo": desligado}, headers=cab("admin"))
    r = client.post("/api/validacao/ajustes-ia", json=corpo, headers=cab()).json()
    assert r["status"] == "desativado" and r["acoes"] == [] and r["diff"] is None
