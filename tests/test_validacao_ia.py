"""tests/test_validacao_ia.py — camada 3 (TASK-014): lotes, prompt, interpretação da resposta e rota /ia.

TASK-055 adiciona: `chamar_openai`/`chamar_claude` (mesmo contrato de `chamar_gemini`) e o dispatch
por provedor em `routers/validacao.py:_preparar_ia`.

O modelo é sempre simulado: nenhum teste faz chamada de rede nem usa chave real.
"""
import asyncio
import json
import types

import anthropic
import openai
import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.ai_chat as ai_chat
import routers.validacao as rota
from config import VALIDACOES_SEED_DIR
from middleware.auth_middleware import create_jwt_token
from services.prompts_validacao import ler_sementes
from services.validacao_ia import (
    MODELOS_RESERVA, MODELOS_RESERVA_CLAUDE, MODELOS_RESERVA_OPENAI, _ERROS_DE_FALLBACK,
    chamar_claude, chamar_openai, dividir_em_lotes, interpretar_resposta, montar_prompt, validar_com_ia,
)

PROMPT, _ = ler_sementes(VALIDACOES_SEED_DIR)
PROMPT = PROMPT["validar-planilhas"]


def L(ativo, op="I", id=None):
    return {"id": id, "ativo": ativo, "operacao": op} if id else {"ativo": ativo, "operacao": op}


# ── lotes ────────────────────────────────────────────────────────────────────
def test_lotes_respeitam_o_tamanho_e_mantem_os_ids_da_tela():
    cabos = [L("P50 A 2 m")] + [L("")] + [L("CAA 2 ABC 3 m")]      # a linha vazia ocupa o índice 1 e não é enviada
    outros = [L(f"1-P{i}") for i in range(5)]
    lotes, truncado = dividir_em_lotes(cabos, outros, tamanho=3)
    assert not truncado and [len(l["cabos"]) + len(l["outros"]) for l in lotes] == [3, 3, 1]  # 2 cabos + 5 outros
    assert lotes[0]["cabos"] == [("CABOS-0", "I", "P50 A 2 m"), ("CABOS-2", "I", "CAA 2 ABC 3 m")]
    assert lotes[0]["outros"][0][0] == "OUTROS-0"


def test_lotes_truncam_acima_do_limite():
    lotes, truncado = dividir_em_lotes([], [L("1-X")] * 10, tamanho=4, limite=6)
    assert truncado and sum(len(l["outros"]) for l in lotes) == 6


def test_sem_linhas_nao_ha_lotes():
    assert dividir_em_lotes([], [L("  ")]) == ([], False)


# ── prompt ───────────────────────────────────────────────────────────────────
def test_prompt_troca_placeholders_e_filtra_achados_do_lote():
    lote = {"cabos": [("CABOS-0", "I", "P50 A 2 m")], "outros": [("OUTROS-3", "I", "1-CFU")]}
    previos = [
        {"linha_id": "OUTROS-3", "regra_id": "C2-CFU-SUPL", "mensagem": "sem SUPL"},
        {"linha_id": "OUTROS-9", "regra_id": "C1-OP", "mensagem": "de outro lote"},
        {"linha_id": "GERAL", "regra_id": "C2-P50", "mensagem": "total baixo"},
    ]
    texto = montar_prompt(PROMPT, lote, previos)
    assert "{{" not in texto
    assert "[CABOS-0] | I | P50 A 2 m" in texto and "[OUTROS-3] | I | 1-CFU" in texto
    assert "OUTROS-3 | C2-CFU-SUPL | sem SUPL" in texto and "GERAL | C2-P50" in texto
    assert "de outro lote" not in texto and "name:" not in texto  # cabeçalho do prompt não vai ao modelo


# ── interpretação da resposta ────────────────────────────────────────────────
IDS = {"cabos": {"CABOS-0"}, "outros": {"OUTROS-1", "OUTROS-2"}}


def resposta(*itens):
    return json.dumps({"achados": list(itens)})


def item(**kw):
    base = {"linha_id": "OUTROS-1", "severidade": "aviso", "regra": "digitacao-suspeita", "problema": "parece CFU", "sugestao": "1-CFU"}
    return {**base, **kw}


def test_resposta_valida():
    achados, desc = interpretar_resposta(resposta(item()), IDS)
    assert desc == [] and achados == [{
        "linha_id": "OUTROS-1", "tabela": "outros", "regra_id": "IA:digitacao-suspeita", "severidade": "aviso",
        "mensagem": "parece CFU", "origem": "ia", "sugestao": "1-CFU"}]


def test_tolera_cerca_de_codigo():
    assert len(interpretar_resposta("```json\n" + resposta(item()) + "\n```", IDS)[0]) == 1


def test_descarta_linha_inventada_e_item_malformado():
    achados, desc = interpretar_resposta(resposta(
        item(linha_id="OUTROS-99"),            # não estava no lote
        item(problema=""),                     # sem descrição
        "texto solto",                         # não é objeto
        item(linha_id="CABOS-0"),
    ), IDS)
    assert [a["linha_id"] for a in achados] == ["CABOS-0"] and achados[0]["tabela"] == "cabos" and len(desc) == 3
    assert {d["motivo"] for d in desc} == {"linha não enviada à IA", "sem descrição do problema", "item não é um objeto"}


def test_severidade_invalida_vira_info_e_regra_ausente_vira_ia():
    a = interpretar_resposta(resposta(item(severidade="gravissimo", regra=None, sugestao=None)), IDS)[0][0]
    assert a["severidade"] == "info" and a["regra_id"] == "IA:ia" and "sugestao" not in a


@pytest.mark.parametrize("texto", ["não é json", "[]", '{"outra": []}', '{"achados": "x"}', ""])
def test_resposta_ilegivel_levanta(texto):
    with pytest.raises(ValueError):
        interpretar_resposta(texto, IDS)


# ── orquestração ─────────────────────────────────────────────────────────────
def rodar(coro):
    return asyncio.run(coro)


def test_orquestracao_ok_junta_os_lotes():
    chamadas = []

    async def chamar(texto):
        chamadas.append(texto)
        id_ = "OUTROS-0" if len(chamadas) == 1 else "OUTROS-2"
        return resposta(item(linha_id=id_))

    r = rodar(validar_com_ia(chamar, PROMPT, [], [L("1-A")] * 3, []))
    assert r["status"] == "ok" and r["lotes"] == 1 and len(chamadas) == 1  # 3 linhas cabem num lote
    assert [a["linha_id"] for a in r["achados"]] == ["OUTROS-0"]


def test_parcial_quando_um_lote_falha():
    import services.validacao_ia as svc
    original = svc.dividir_em_lotes
    svc.dividir_em_lotes = lambda c, o: original(c, o, tamanho=1)
    try:
        estado = {"n": 0}

        async def chamar(texto):
            estado["n"] += 1
            if estado["n"] == 2:
                raise RuntimeError("cota esgotada")
            return resposta(item(linha_id="OUTROS-0" if estado["n"] == 1 else "OUTROS-2"))

        r = rodar(validar_com_ia(chamar, PROMPT, [], [L("1-A"), L("1-B"), L("1-C")], []))
        assert r["status"] == "parcial" and r["lotes"] == 3 and len(r["achados"]) == 2
        assert "lote 2" in r["mensagem"] and "cota" in r["mensagem"]

        async def sempre_falha(texto):
            raise RuntimeError("fora do ar")

        r = rodar(validar_com_ia(sempre_falha, PROMPT, [], [L("1-A"), L("1-B")], []))
        assert r["status"] == "erro" and r["achados"] == []
    finally:
        svc.dividir_em_lotes = original


def test_json_quebrado_nao_derruba_e_vira_erro():
    async def chamar(texto):
        return "```isso nao e json```"

    r = rodar(validar_com_ia(chamar, PROMPT, [], [L("1-A")], []))
    assert r["status"] == "erro" and "JSON" in r["mensagem"]


def test_sem_linhas_nao_chama_o_modelo():
    async def chamar(texto):
        raise AssertionError("não devia chamar")

    r = rodar(validar_com_ia(chamar, PROMPT, [L("")], [], []))
    assert r["status"] == "ok" and r["lotes"] == 0


# ── rota ─────────────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    database.init_db()
    ai_chat._rate_hits.clear()
    return TestClient(appmod.app)


@pytest.fixture
def modelo(monkeypatch):
    """Substitui a chamada ao Gemini. `modelo.resposta` define o retorno; `modelo.chamadas` guarda o que foi enviado."""
    class Falso:
        resposta = resposta(item(linha_id="OUTROS-0"))
        erro = None
        chamadas = []

    async def falso(api_key, modelos, texto, temperatura=0.0):
        Falso.chamadas.append({"api_key": api_key, "modelos": modelos, "texto": texto, "temperatura": temperatura})
        if Falso.erro:
            raise Falso.erro
        return Falso.resposta

    Falso.chamadas = []
    monkeypatch.setattr(rota, "chamar_gemini", falso)
    return Falso


def cab(role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


CORPO = {"cabos": [], "outros": [L("1-CFUU")], "projeto_codigo": "229"}


def test_ia_exige_autenticacao(client):
    assert client.post("/api/validacao/ia", json=CORPO).status_code == 401


def test_sem_chave_devolve_status_e_nao_chama_o_modelo(client, modelo):
    r = client.post("/api/validacao/ia", json=CORPO, headers=cab())
    assert r.status_code == 200 and r.json()["status"] == "sem_chave" and modelo.chamadas == []


def test_com_chave_padrao_devolve_achados_da_ia(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "CHAVE-PADRAO")
    r = client.post("/api/validacao/ia", json={**CORPO, "achados_previos": [
        {"linha_id": "OUTROS-0", "regra_id": "C1-BASE", "mensagem": "CFUU não encontrado"}]}, headers=cab())
    corpo = r.json()
    assert r.status_code == 200 and corpo["status"] == "ok" and corpo["resumo"]["aviso"] == 1
    assert corpo["achados"][0]["origem"] == "ia" and corpo["achados"][0]["linha_id"] == "OUTROS-0"
    enviado = modelo.chamadas[0]
    assert enviado["api_key"] == "CHAVE-PADRAO" and enviado["temperatura"] == 0.0
    assert enviado["modelos"][0] == "gemini-3.1-flash-lite" and "[OUTROS-0] | I | 1-CFUU" in enviado["texto"]
    assert "C1-BASE | CFUU não encontrado" in enviado["texto"]


def test_chave_do_usuario_prevalece_e_nao_e_salva(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "CHAVE-PADRAO")
    client.post("/api/validacao/ia", json=CORPO, headers={**cab(), "X-Gemini-Key": "CHAVE-DO-USUARIO"})
    assert modelo.chamadas[0]["api_key"] == "CHAVE-DO-USUARIO"
    assert ai_chat._ler_configuracao("gemini_api_key") is None


def test_falha_da_ia_nao_vira_erro_http(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    modelo.erro = RuntimeError("429 quota exhausted")
    r = client.post("/api/validacao/ia", json=CORPO, headers=cab())
    assert r.status_code == 200 and r.json()["status"] == "erro" and "429" in r.json()["mensagem"]


def test_resposta_quebrada_nao_trava(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    modelo.resposta = "isso não é json"
    r = client.post("/api/validacao/ia", json=CORPO, headers=cab())
    assert r.status_code == 200 and r.json()["status"] == "erro" and r.json()["achados"] == []


def test_limite_por_minuto_so_vale_para_a_chave_padrao(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    monkeypatch.setattr(ai_chat, "RATE_LIMIT_POR_MINUTO", 2)
    codigos = [client.post("/api/validacao/ia", json=CORPO, headers=cab()).status_code for _ in range(3)]
    assert codigos == [200, 200, 429]
    própria = [client.post("/api/validacao/ia", json=CORPO, headers={**cab(), "X-Gemini-Key": "MINHA"}).status_code for _ in range(4)]
    assert própria == [200] * 4


def test_usa_o_prompt_editado_do_projeto_e_respeita_ativo_false(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    editado = PROMPT.replace("Sua tarefa é APENAS apontar", "MARCADOR-DO-PROJETO apontar")
    assert client.post("/api/validacao/prompts/validar-planilhas", headers=cab("admin"),
                       json={"projeto_codigo": "229", "conteudo": editado}).status_code == 200
    client.post("/api/validacao/ia", json=CORPO, headers=cab())
    assert "MARCADOR-DO-PROJETO" in modelo.chamadas[-1]["texto"]
    outro = client.post("/api/validacao/ia", json={**CORPO, "projeto_codigo": "027"}, headers=cab())
    assert "MARCADOR-DO-PROJETO" not in modelo.chamadas[-1]["texto"] and outro.json()["status"] == "ok"

    desligado = editado.replace("ativo: true", "ativo: false")
    client.post("/api/validacao/prompts/validar-planilhas", headers=cab("admin"),
                json={"projeto_codigo": "229", "conteudo": desligado})
    antes = len(modelo.chamadas)
    r = client.post("/api/validacao/ia", json=CORPO, headers=cab())
    assert r.json()["status"] == "desativado" and len(modelo.chamadas) == antes


def test_prompt_inexistente(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    r = client.post("/api/validacao/ia", json={**CORPO, "prompt_id": "nao-existe"}, headers=cab())
    assert r.json()["status"] == "indisponivel"


# ── TASK-054: listas sem modelos mortos e fallback ampliado ──────────────────
def test_modelos_reserva_sem_entradas_mortas():
    assert "gemini-1.5-flash" not in MODELOS_RESERVA and "gemini-1.5-pro" not in MODELOS_RESERVA
    assert "gemini-2.5-flash" not in MODELOS_RESERVA  # desligamento anunciado para 16/10/2026
    assert MODELOS_RESERVA[0] == "gemini-3.1-flash-lite"  # modelo padrão


def test_erros_de_fallback_cobrem_multiturn_e_invalid_argument():
    assert "multiturn" in _ERROS_DE_FALLBACK and "invalid_argument" in _ERROS_DE_FALLBACK


# ── TASK-055: chamar_openai / chamar_claude (mesmo contrato de chamar_gemini) ────────────────────
def rodar(coro):
    return asyncio.run(coro)


def test_chamar_openai_usa_modo_json_e_devolve_o_texto(monkeypatch):
    chamadas = []

    class FakeCompletions:
        async def create(self, **kw):
            chamadas.append(kw)
            msg = types.SimpleNamespace(content='{"achados": []}')
            return types.SimpleNamespace(choices=[types.SimpleNamespace(message=msg)])

    class FakeAsyncOpenAI:
        def __init__(self, api_key=None):
            self.api_key = api_key
            self.chat = types.SimpleNamespace(completions=FakeCompletions())

    monkeypatch.setattr(openai, "AsyncOpenAI", FakeAsyncOpenAI)
    resultado = rodar(chamar_openai("CHAVE", ["gpt-4o-mini"], "texto do prompt", 0.0))
    assert resultado == '{"achados": []}'
    assert chamadas[0]["model"] == "gpt-4o-mini" and chamadas[0]["response_format"] == {"type": "json_object"}
    assert chamadas[0]["messages"] == [{"role": "user", "content": "texto do prompt"}]


def test_chamar_openai_tenta_o_proximo_modelo_so_para_erro_de_cota(monkeypatch):
    class FakeCompletions:
        def __init__(self):
            self.n = 0

        async def create(self, **kw):
            self.n += 1
            if self.n == 1:
                raise RuntimeError("429 rate limit exceeded")
            msg = types.SimpleNamespace(content="ok")
            return types.SimpleNamespace(choices=[types.SimpleNamespace(message=msg)])

    fake = FakeCompletions()

    class FakeAsyncOpenAI:
        def __init__(self, api_key=None):
            self.chat = types.SimpleNamespace(completions=fake)

    monkeypatch.setattr(openai, "AsyncOpenAI", FakeAsyncOpenAI)
    resultado = rodar(chamar_openai("CHAVE", ["gpt-4o-mini", "gpt-4o"], "texto", 0.0))
    assert resultado == "ok" and fake.n == 2


def test_chamar_openai_erro_sem_gatilho_nao_tenta_outro_modelo(monkeypatch):
    class FakeCompletions:
        async def create(self, **kw):
            raise RuntimeError("erro irrecuperável")

    class FakeAsyncOpenAI:
        def __init__(self, api_key=None):
            self.chat = types.SimpleNamespace(completions=FakeCompletions())

    monkeypatch.setattr(openai, "AsyncOpenAI", FakeAsyncOpenAI)
    with pytest.raises(RuntimeError, match="irrecuperável"):
        rodar(chamar_openai("CHAVE", ["gpt-4o-mini", "gpt-4o"], "texto", 0.0))


def test_chamar_claude_devolve_o_texto_do_bloco(monkeypatch):
    chamadas = []

    class FakeMessages:
        async def create(self, **kw):
            chamadas.append(kw)
            bloco = types.SimpleNamespace(type="text", text='{"achados": []}')
            return types.SimpleNamespace(content=[bloco])

    class FakeAsyncAnthropic:
        def __init__(self, api_key=None):
            self.api_key = api_key
            self.messages = FakeMessages()

    monkeypatch.setattr(anthropic, "AsyncAnthropic", FakeAsyncAnthropic)
    resultado = rodar(chamar_claude("CHAVE", ["claude-haiku-5-5"], "texto do prompt", 0.0))
    assert resultado == '{"achados": []}'
    assert chamadas[0]["model"] == "claude-haiku-5-5" and chamadas[0]["max_tokens"] == 4096
    assert chamadas[0]["messages"] == [{"role": "user", "content": "texto do prompt"}]


def test_chamar_claude_tenta_o_proximo_modelo_so_para_erro_de_cota(monkeypatch):
    class FakeMessages:
        def __init__(self):
            self.n = 0

        async def create(self, **kw):
            self.n += 1
            if self.n == 1:
                raise RuntimeError("overloaded_error: 503 unavailable")
            return types.SimpleNamespace(content=[types.SimpleNamespace(type="text", text="ok")])

    fake = FakeMessages()

    class FakeAsyncAnthropic:
        def __init__(self, api_key=None):
            self.messages = fake

    monkeypatch.setattr(anthropic, "AsyncAnthropic", FakeAsyncAnthropic)
    resultado = rodar(chamar_claude("CHAVE", ["claude-haiku-5-5", "claude-sonnet-5-5"], "texto", 0.0))
    assert resultado == "ok" and fake.n == 2


def test_modelos_reserva_claude_tem_haiku_primeiro():
    assert MODELOS_RESERVA_CLAUDE[0] == "claude-haiku-5-5" and "claude-sonnet-5-5" in MODELOS_RESERVA_CLAUDE


def test_modelos_reserva_openai_nao_vazia():
    assert MODELOS_RESERVA_OPENAI


# ── TASK-055: _preparar_ia escolhe a função certa por provedor salvo ─────────
def _salvar_provider(provider):
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES ('ai_provider', ?)", (provider,))
    conn.commit()
    conn.close()


@pytest.fixture
def modelo_openai(monkeypatch):
    class Falso:
        resposta = json.dumps({"achados": []})
        erro = None
        chamadas = []

    async def falso(api_key, modelos, texto, temperatura=0.0):
        Falso.chamadas.append({"api_key": api_key, "modelos": modelos, "texto": texto})
        if Falso.erro:
            raise Falso.erro
        return Falso.resposta

    Falso.chamadas = []
    monkeypatch.setattr(rota, "chamar_openai", falso)
    return Falso


@pytest.fixture
def modelo_claude(monkeypatch):
    class Falso:
        resposta = json.dumps({"achados": []})
        erro = None
        chamadas = []

    async def falso(api_key, modelos, texto, temperatura=0.0):
        Falso.chamadas.append({"api_key": api_key, "modelos": modelos, "texto": texto})
        if Falso.erro:
            raise Falso.erro
        return Falso.resposta

    Falso.chamadas = []
    monkeypatch.setattr(rota, "chamar_claude", falso)
    return Falso


def test_provider_openai_salvo_usa_chamar_openai_e_nao_chamar_gemini(client, modelo, modelo_openai, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "CHAVE-OPENAI")
    _salvar_provider("openai")
    r = client.post("/api/validacao/ia", json=CORPO, headers=cab())
    assert r.status_code == 200 and r.json()["status"] == "ok"
    assert len(modelo_openai.chamadas) == 1 and modelo.chamadas == []
    assert modelo_openai.chamadas[0]["api_key"] == "CHAVE-OPENAI"
    # O "modelo" do cabeçalho do prompt salvo é um ID Gemini — não entra na lista para OpenAI.
    assert "gemini-3.1-flash-lite" not in modelo_openai.chamadas[0]["modelos"]
    assert modelo_openai.chamadas[0]["modelos"] == MODELOS_RESERVA_OPENAI


def test_provider_claude_salvo_usa_chamar_claude_com_chave_do_header(client, modelo, modelo_claude, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "CHAVE-PADRAO-CLAUDE")
    _salvar_provider("claude")
    r = client.post("/api/validacao/ia", json=CORPO, headers={**cab(), "X-Anthropic-Key": "CHAVE-DO-USUARIO"})
    assert r.status_code == 200 and r.json()["status"] == "ok"
    assert modelo_claude.chamadas[0]["api_key"] == "CHAVE-DO-USUARIO" and modelo.chamadas == []
    assert modelo_claude.chamadas[0]["modelos"][0] == "claude-haiku-5-5"


def test_provider_claude_sem_nenhuma_chave_devolve_sem_chave(client, modelo, modelo_claude):
    _salvar_provider("claude")
    r = client.post("/api/validacao/ia", json=CORPO, headers=cab())
    assert r.json()["status"] == "sem_chave" and modelo_claude.chamadas == [] and modelo.chamadas == []


def test_provider_anthropic_salvo_e_tratado_como_claude(client, modelo, modelo_claude, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "K")
    _salvar_provider("anthropic")
    r = client.post("/api/validacao/ia", json=CORPO, headers=cab())
    assert r.json()["status"] == "ok" and len(modelo_claude.chamadas) == 1


def test_provider_desconhecido_salvo_cai_para_gemini(client, modelo, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    _salvar_provider("deepseek")
    r = client.post("/api/validacao/ia", json=CORPO, headers=cab())
    assert r.json()["status"] == "ok" and len(modelo.chamadas) == 1
