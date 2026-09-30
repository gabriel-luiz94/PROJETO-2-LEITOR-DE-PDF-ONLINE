"""tests/test_correcao_ia.py — correção assistida (TASK-015): a IA só propõe; cada proposta é reconferida.

O modelo é sempre simulado: nenhum teste faz chamada de rede nem usa chave real.
"""
import asyncio
import json

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.validacao as rota
from config import VALIDACOES_SEED_DIR
from middleware.auth_middleware import create_jwt_token
from services.correcao_ia import corrigir_com_ia, interpretar_correcoes, linhas_citadas
from services.prompts_validacao import ler_sementes

PROMPT = ler_sementes(VALIDACOES_SEED_DIR)[0]["corrigir-planilhas"]


def L(ativo, op="I"):
    return {"ativo": ativo, "operacao": op}


def resp(*correcoes):
    return json.dumps({"correcoes": list(correcoes)})


def c(linha_id="OUTROS-0", depois="DT11/300 1-CFU", **kw):
    return {"linha_id": linha_id, "tabela": "outros", "antes": "qualquer coisa", "depois": depois, "motivo": "digitação", **kw}


REAIS = {"OUTROS-0": ("outros", "I", "DT11/300 1-CFUU"), "CABOS-1": ("cabos", "I", "3-CFU")}


# ── linhas citadas ───────────────────────────────────────────────────────────
def test_so_as_linhas_citadas_vao_para_a_ia_com_o_id_da_tela():
    cabos = [L("P50 A 2 m"), L("3-CFU"), L("")]
    outros = [L("1-A"), L("1-B"), L("1-C")]
    achados = [{"linha_id": "CABOS-1"}, {"linha_id": "OUTROS-2"}, {"linha_id": "GERAL"}, {"linha_id": "TOTALIZADORA-CABOS-0"}]
    cab, out = linhas_citadas(cabos, outros, achados)
    assert [l["id"] for l in cab] == ["CABOS-1"] and [l["id"] for l in out] == ["OUTROS-2"]


def test_linha_vazia_citada_nao_e_enviada():
    assert linhas_citadas([L("")], [], [{"linha_id": "CABOS-0"}]) == ([], [])


# ── interpretação e reconferência ────────────────────────────────────────────
def test_proposta_valida_usa_o_antes_real_e_nao_o_da_ia():
    correcoes, descartadas = interpretar_correcoes(resp(c()), REAIS)
    assert descartadas == [] and correcoes == [{
        "linha_id": "OUTROS-0", "tabela": "outros", "antes": "DT11/300 1-CFUU",
        "depois": "DT11/300 1-CFU", "motivo": "digitação", "avisos": []}]


def test_tolera_cerca_de_codigo():
    assert len(interpretar_correcoes("```json\n" + resp(c()) + "\n```", REAIS)[0]) == 1


@pytest.mark.parametrize("item,trecho", [
    (c(linha_id="OUTROS-99"), "não enviada"),
    (c(linha_id=None), "não enviada"),
    (c(depois=""), "sem o texto corrigido"),
    (c(depois=None), "sem o texto corrigido"),
    (c(depois="  dt11/300   1-cfuu "), "não muda"),        # só espaço e caixa diferentes
    (c(depois="DT11/300 1-CFUU"), "não muda"),
    (c(depois="3-IP RECAL"), "ainda reprova"),             # token sem par (C1-OUT-ORFAO)
    (c(depois="x-CFU"), "ainda reprova"),                  # quantidade não numérica
    (c(depois="CAA 2 35 m"), "ainda reprova"),             # formato de cabo em Outros
])
def test_proposta_descartada_com_motivo(item, trecho):
    correcoes, descartadas = interpretar_correcoes(resp(item), REAIS)
    assert correcoes == [] and len(descartadas) == 1 and trecho in descartadas[0]["motivo"]


def test_reprovada_na_camada_1_para_cabos():
    item = c(linha_id="CABOS-1", depois="4-CFU")           # continua no formato de Outros dentro de Cabos
    correcoes, descartadas = interpretar_correcoes(resp(item), REAIS)
    assert correcoes == [] and "ainda reprova" in descartadas[0]["motivo"]
    ok = c(linha_id="CABOS-1", depois="CAA 2 ABC 35 m")
    assert interpretar_correcoes(resp(ok), REAIS)[0][0]["tabela"] == "cabos"


def test_operacao_da_linha_nao_impede_nem_e_alterada():
    reais = {"OUTROS-0": ("outros", "", "DT11/300 1-CFUU")}     # operação vazia é outro achado (C1-OP), não do ativo
    correcoes, _ = interpretar_correcoes(resp(c()), reais)
    assert len(correcoes) == 1 and "operacao" not in correcoes[0]


def test_so_a_primeira_proposta_por_linha():
    correcoes, descartadas = interpretar_correcoes(resp(c(depois="DT11/300 1-CFU"), c(depois="DT11/300 2-CFU")), REAIS)
    assert [x["depois"] for x in correcoes] == ["DT11/300 1-CFU"] and "mais de uma" in descartadas[0]["motivo"]


def test_aviso_da_camada_1_nao_descarta_mas_acompanha_a_proposta():
    correcoes, _ = interpretar_correcoes(resp(c(depois="1-P50 1-P50")), REAIS)   # duplicidade é aviso, não erro
    assert len(correcoes) == 1 and "aparece 2 vezes" in correcoes[0]["avisos"][0]


def test_item_malformado_e_motivo_nao_texto():
    correcoes, descartadas = interpretar_correcoes(resp("solto", c(motivo=123)), REAIS)
    assert len(descartadas) == 1 and correcoes[0]["motivo"] == ""


@pytest.mark.parametrize("texto", ["não é json", "[]", '{"achados": []}', '{"correcoes": "x"}', ""])
def test_resposta_ilegivel_levanta(texto):
    with pytest.raises(ValueError):
        interpretar_correcoes(texto, REAIS)


# ── orquestração ─────────────────────────────────────────────────────────────
def rodar(coro):
    return asyncio.run(coro)


ACHADOS = [{"linha_id": "OUTROS-0", "regra_id": "IA:digitacao", "mensagem": "parece CFU", "sugestao": "DT11/300 1-CFU"}]


def test_prompt_recebe_so_a_linha_citada_e_a_sugestao_da_ia():
    enviados = []

    async def chamar(texto):
        enviados.append(texto)
        return resp(c())

    r = rodar(corrigir_com_ia(chamar, PROMPT, [L("P50 A 2 m")], [L("DT11/300 1-CFUU"), L("1-OUTRA")], ACHADOS))
    assert r["status"] == "ok" and len(r["correcoes"]) == 1 and r["correcoes"][0]["linha_id"] == "OUTROS-0"
    texto = enviados[0]
    assert "[OUTROS-0] | I | DT11/300 1-CFUU" in texto and "1-OUTRA" not in texto and "P50 A 2 m" not in texto
    assert "OUTROS-0 | IA:digitacao | parece CFU | sugestão: DT11/300 1-CFU" in texto and "{{" not in texto


def test_sem_linhas_citadas_nao_chama_o_modelo():
    async def chamar(texto):
        raise AssertionError("não devia chamar")

    r = rodar(corrigir_com_ia(chamar, PROMPT, [L("P50")], [L("1-A")], [{"linha_id": "GERAL"}]))
    assert r["status"] == "ok" and r["lotes"] == 0 and r["correcoes"] == []


def test_falha_do_modelo_vira_status_e_nao_levanta():
    async def chamar(texto):
        raise RuntimeError("429 quota exhausted")

    r = rodar(corrigir_com_ia(chamar, PROMPT, [], [L("DT11/300 1-CFUU")], ACHADOS))
    assert r["status"] == "erro" and "429" in r["mensagem"] and r["correcoes"] == []


def test_json_quebrado_vira_erro():
    async def chamar(texto):
        return "```isso nao e json```"

    r = rodar(corrigir_com_ia(chamar, PROMPT, [], [L("DT11/300 1-CFUU")], ACHADOS))
    assert r["status"] == "erro" and "JSON" in r["mensagem"]


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
        resposta = resp(c())
        erro = None
        chamadas = []

    async def falso(api_key, modelos, texto, temperatura=0.0):
        Falso.chamadas.append({"api_key": api_key, "modelos": modelos, "texto": texto})
        if Falso.erro:
            raise Falso.erro
        return Falso.resposta

    Falso.chamadas = []
    monkeypatch.setattr(rota, "chamar_gemini", falso)
    return Falso


def cab(role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


CORPO = {"cabos": [], "outros": [L("DT11/300 1-CFUU")], "achados": ACHADOS, "projeto_codigo": "229"}


def test_corrigir_exige_autenticacao(client):
    assert client.post("/api/validacao/corrigir", json=CORPO).status_code == 401


def test_corrigir_sem_chave_nao_chama_o_modelo(client, modelo):
    r = client.post("/api/validacao/corrigir", json=CORPO, headers=cab())
    assert r.status_code == 200 and r.json()["status"] == "sem_chave" and modelo.chamadas == []


def test_corrigir_devolve_proposta_sem_aplicar_nada(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    r = client.post("/api/validacao/corrigir", json=CORPO, headers=cab())
    corpo = r.json()
    assert r.status_code == 200 and corpo["status"] == "ok"
    assert corpo["correcoes"] == [{"linha_id": "OUTROS-0", "tabela": "outros", "antes": "DT11/300 1-CFUU",
                                   "depois": "DT11/300 1-CFU", "motivo": "digitação", "avisos": []}]
    assert "corrigir-planilhas" not in modelo.chamadas[0]["texto"] and "[OUTROS-0] | I | DT11/300 1-CFUU" in modelo.chamadas[0]["texto"]


def test_corrigir_descarta_o_que_ainda_reprova(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    modelo.resposta = resp(c(depois="3-IP RECAL"))
    corpo = client.post("/api/validacao/corrigir", json=CORPO, headers=cab()).json()
    assert corpo["correcoes"] == [] and "ainda reprova" in corpo["descartadas"][0]["motivo"]


def test_corrigir_falha_da_ia_nao_vira_erro_http(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    modelo.erro = RuntimeError("fora do ar")
    r = client.post("/api/validacao/corrigir", json=CORPO, headers=cab())
    assert r.status_code == 200 and r.json()["status"] == "erro" and r.json()["correcoes"] == []


def test_corrigir_respeita_prompt_desativado_e_editado_pelo_admin(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    editado = PROMPT.replace("Você propõe correções", "MARCADOR-CORRECAO propõe correções")
    assert client.post("/api/validacao/prompts/corrigir-planilhas", headers=cab("admin"),
                       json={"projeto_codigo": "229", "conteudo": editado}).status_code == 200
    client.post("/api/validacao/corrigir", json=CORPO, headers=cab())
    assert "MARCADOR-CORRECAO" in modelo.chamadas[-1]["texto"]
    desligado = editado.replace("ativo: true", "ativo: false")
    client.post("/api/validacao/prompts/corrigir-planilhas", headers=cab("admin"),
                json={"projeto_codigo": "229", "conteudo": desligado})
    antes = len(modelo.chamadas)
    assert client.post("/api/validacao/corrigir", json=CORPO, headers=cab()).json()["status"] == "desativado"
    assert len(modelo.chamadas) == antes
