"""tests/test_ai_chat.py — chat geral de IA (`routers/ai_chat.py`).

TASK-054: troca da API de sessão do Gemini (`chats.create`/`send_message_stream`) por uma chamada
única (`generate_content_stream`) com o histórico embutido em `contents`, e ampliação dos gatilhos
de fallback para cobrir o erro "Multiturn chat is not enabled for this model".
TASK-055: branches reais para os provedores OpenAI e Claude (Anthropic), além do Gemini.

Nenhum teste faz chamada de rede real: os SDKs (`google.genai`, `openai`, `anthropic`) são sempre
substituídos por fakes via monkeypatch — mesmo padrão já usado em tests/test_obras_contexto.py.
"""
import types

import anthropic
import openai
import pytest
from fastapi.testclient import TestClient
from google import genai as google_genai

import app as appmod
import database
import routers.ai_chat as ai_chat
from middleware.auth_middleware import create_jwt_token


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    database.init_db()
    ai_chat._rate_hits.clear()
    return TestClient(appmod.app)


def cab(role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


def corpo(prompt, history=None, **kw):
    base = {"prompt": prompt, "history": history or [], "projeto_codigo": "229"}
    base.update(kw)
    return base


# ── fakes do SDK do Gemini (google.genai) ────────────────────────────────────
class _FakeChunk:
    def __init__(self, text):
        self.text = text


class FakeModelsGemini:
    """Substitui `client.aio.models.generate_content_stream`. `respostas[modelo]` define o
    comportamento: uma lista de textos (sucesso, em chunks) ou uma Exception (falha)."""

    def __init__(self, respostas):
        self.respostas = respostas
        self.chamadas = []  # [{"model", "contents", "system_instruction"}]

    async def generate_content_stream(self, *, model, contents, config=None):
        self.chamadas.append({
            "model": model, "contents": contents,
            "system_instruction": getattr(config, "system_instruction", None),
        })
        resultado = self.respostas.get(model)
        if isinstance(resultado, Exception):
            raise resultado

        async def _gerador():
            for texto in (resultado or []):
                yield _FakeChunk(texto)
        return _gerador()


class FakeAio:
    def __init__(self, models):
        self.models = models


class FakeGeminiClient:
    """Substitui `genai.Client(api_key=...)`. `FakeGeminiClient.ULTIMO` guarda a última instância
    criada, para o teste inspecionar `.aio.models.chamadas`."""
    ULTIMO = None

    def __init__(self, api_key=None, respostas=None):
        self.api_key = api_key
        modelos = FakeModelsGemini(respostas or FakeGeminiClient.RESPOSTAS)
        self.aio = FakeAio(modelos)
        FakeGeminiClient.ULTIMO = self

    RESPOSTAS = {}


def _gemini_client_com(monkeypatch, respostas):
    def _factory(api_key=None):
        return FakeGeminiClient(api_key=api_key, respostas=respostas)
    monkeypatch.setattr(google_genai, "Client", _factory)


# ── TASK-054: generate_content_stream sem sessão ─────────────────────────────
def test_gemini_usa_generate_content_stream_sem_sessao_e_aplica_system_instruction(client, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    _gemini_client_com(monkeypatch, {"gemini-3.1-flash-lite": ["Olá, ", "mundo!"]})

    r = client.post("/api/gemini/chat", json=corpo("oi, como vai você?"), headers=cab())
    assert r.status_code == 200 and r.text == "Olá, mundo!"

    chamada = FakeGeminiClient.ULTIMO.aio.models.chamadas[0]
    assert chamada["model"] == "gemini-3.1-flash-lite"
    assert chamada["system_instruction"]  # regras/contexto de obras aplicados (RULES Regra 11: não loga conteúdo)
    # Chamada única com o histórico embutido em `contents` — nada de `chats.create`/sessão.
    assert chamada["contents"][-1] == {"role": "user", "parts": [{"text": "oi, como vai você?"}]}


def test_historico_e_enviado_no_contents_em_ordem(client, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    _gemini_client_com(monkeypatch, {"gemini-3.1-flash-lite": ["resposta com contexto"]})

    historico = [
        {"role": "user", "parts": [{"text": "qual o ativo mais usado?"}]},
        {"role": "model", "parts": [{"text": "CFU é o mais comum."}]},
    ]
    r = client.post(
        "/api/gemini/chat",
        json=corpo("e o segundo mais usado, qual é? me conte mais sobre isso detalhadamente", history=historico),
        headers=cab(),
    )
    assert r.status_code == 200 and r.text == "resposta com contexto"

    contents = FakeGeminiClient.ULTIMO.aio.models.chamadas[0]["contents"]
    assert contents[0] == historico[0] and contents[1] == historico[1]
    assert contents[2]["role"] == "user" and "segundo mais usado" in contents[2]["parts"][0]["text"]


# ── TASK-054: fallback cobre o erro de "multiturn" / invalid_argument ────────
def test_erro_multiturn_aciona_fallback_para_o_proximo_modelo(client, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    erro_multiturn = RuntimeError(
        "400 INVALID_ARGUMENT. {'error': {'code': 400, 'message': 'Multiturn chat is not enabled "
        "for this model', 'status': 'INVALID_ARGUMENT'}}"
    )
    _gemini_client_com(monkeypatch, {
        "gemini-3.1-flash-lite": erro_multiturn,
        "gemini-3.6-flash": ["funcionou no segundo modelo"],
    })

    # Mensagem exata reportada pelo usuário, com histórico prévio (para não cair no "forçar modelo curto").
    historico = [{"role": "user", "parts": [{"text": "oi"}]}, {"role": "model", "parts": [{"text": "oi, tudo bem?"}]}]
    r = client.post(
        "/api/gemini/chat",
        json=corpo("substitua o texto S3 por S3T na tabela outros", history=historico),
        headers=cab(),
    )
    assert r.status_code == 200
    assert "[ERRO]" not in r.text and r.text == "funcionou no segundo modelo"
    modelos_tentados = [c["model"] for c in FakeGeminiClient.ULTIMO.aio.models.chamadas]
    assert modelos_tentados == ["gemini-3.1-flash-lite", "gemini-3.6-flash"]


def test_erro_sem_gatilho_de_fallback_nao_tenta_o_proximo_modelo(client, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    _gemini_client_com(monkeypatch, {
        "gemini-3.1-flash-lite": RuntimeError("erro irrecuperável qualquer"),
        "gemini-3.6-flash": ["não devia chegar aqui"],
    })
    r = client.post("/api/gemini/chat", json=corpo("mensagem curta qualquer"), headers=cab())
    assert r.status_code == 200 and "[ERRO] erro irrecuperável qualquer" in r.text
    assert len(FakeGeminiClient.ULTIMO.aio.models.chamadas) == 1  # não tentou o segundo modelo


def test_todos_os_modelos_falhando_com_gatilho_devolve_erro_sem_derrubar(client, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    _gemini_client_com(monkeypatch, {
        "gemini-3.1-flash-lite": RuntimeError("429 quota exhausted"),
        "gemini-3.6-flash": RuntimeError("429 quota exhausted"),
        "gemini-3.5-flash": RuntimeError("429 quota exhausted"),
    })
    r = client.post("/api/gemini/chat", json=corpo("mensagem curta qualquer"), headers=cab())
    assert r.status_code == 200 and "[ERRO]" in r.text and "quota" in r.text


# ── TASK-054: listas de modelo sem as entradas mortas ────────────────────────
def test_lista_de_preferencias_do_chat_nao_tem_modelos_mortos(client, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "K")
    # Todos os 3 modelos esperados falham com gatilho de fallback: a lista completa tentada fica
    # visível nas chamadas. Se o código tentasse algum modelo fora desta lista, o fake trataria
    # como sucesso vazio e pararia ali — a asserção de igualdade abaixo cobre isso também.
    erro_cota = RuntimeError("429 quota exhausted")
    _gemini_client_com(monkeypatch, {
        "gemini-3.1-flash-lite": erro_cota, "gemini-3.6-flash": erro_cota, "gemini-3.5-flash": erro_cota,
    })

    r = client.post(
        "/api/gemini/chat",
        json=corpo("mensagem bem mais longa para não forçar o modelo curto de propósito aqui"),
        headers=cab(),
    )
    assert r.status_code == 200 and "[ERRO]" in r.text
    modelos_tentados = [c["model"] for c in FakeGeminiClient.ULTIMO.aio.models.chamadas]
    assert "gemini-1.5-flash" not in modelos_tentados and "gemini-1.5-pro" not in modelos_tentados
    assert "gemini-2.5-flash" not in modelos_tentados
    assert modelos_tentados == ["gemini-3.1-flash-lite", "gemini-3.6-flash", "gemini-3.5-flash"]


# ── TASK-055: branch OpenAI ───────────────────────────────────────────────────
class _FakeDelta:
    def __init__(self, content):
        self.content = content


class _FakeChoiceOpenAI:
    def __init__(self, content):
        self.delta = _FakeDelta(content)


class _FakeChunkOpenAI:
    def __init__(self, content):
        self.choices = [_FakeChoiceOpenAI(content)]


class FakeCompletionsOpenAI:
    def __init__(self, respostas):
        self.respostas = respostas
        self.chamadas = []

    async def create(self, *, model, messages, stream=False):
        self.chamadas.append({"model": model, "messages": messages, "stream": stream})
        resultado = self.respostas.get(model)
        if isinstance(resultado, Exception):
            raise resultado

        async def _gerador():
            for texto in (resultado or []):
                yield _FakeChunkOpenAI(texto)
        return _gerador()


class FakeAsyncOpenAI:
    ULTIMO = None

    def __init__(self, api_key=None, base_url=None, respostas=None):
        self.api_key = api_key
        self.base_url = base_url
        completions = FakeCompletionsOpenAI(respostas or {})
        self.chat = types.SimpleNamespace(completions=completions)
        FakeAsyncOpenAI.ULTIMO = self


def _openai_client_com(monkeypatch, respostas):
    def _factory(api_key=None, base_url=None):
        return FakeAsyncOpenAI(api_key=api_key, base_url=base_url, respostas=respostas)
    monkeypatch.setattr(openai, "AsyncOpenAI", _factory)


def test_chat_com_provider_openai_responde_em_streaming(client, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "K")
    _openai_client_com(monkeypatch, {"gpt-4o-mini": ["resposta ", "da OpenAI"]})

    r = client.post("/api/gemini/chat", json=corpo("oi, como vai?", provider="openai"), headers=cab())
    assert r.status_code == 200 and r.text == "resposta da OpenAI"
    chamada = FakeAsyncOpenAI.ULTIMO.chat.completions.chamadas[0]
    assert chamada["model"] == "gpt-4o-mini" and chamada["stream"] is True
    assert chamada["messages"][0]["role"] == "system"
    assert chamada["messages"][-1] == {"role": "user", "content": "oi, como vai?"}


def test_chat_openai_fallback_em_erro_de_cota(client, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "K")
    _openai_client_com(monkeypatch, {
        "gpt-4o-mini": RuntimeError("429 you exceeded your quota"),
        "gpt-4o": ["funcionou no modelo de reserva"],
    })
    r = client.post("/api/gemini/chat", json=corpo("oi", provider="openai"), headers=cab())
    assert r.status_code == 200 and r.text == "funcionou no modelo de reserva"


def test_chat_sem_chave_openai_devolve_401(client):
    r = client.post("/api/gemini/chat", json=corpo("oi", provider="openai"), headers=cab())
    assert r.status_code == 401


# ── TASK-055: branch Claude ───────────────────────────────────────────────────
class FakeTextStream:
    def __init__(self, textos):
        self._textos = textos

    def __aiter__(self):
        return self._gerador()

    async def _gerador(self):
        for t in self._textos:
            yield t


class FakeStreamManager:
    def __init__(self, textos_ou_erro):
        self._item = textos_ou_erro

    async def __aenter__(self):
        if isinstance(self._item, Exception):
            raise self._item
        return types.SimpleNamespace(text_stream=FakeTextStream(self._item))

    async def __aexit__(self, *a):
        return False


class FakeMessagesClaude:
    def __init__(self, respostas):
        self.respostas = respostas
        self.chamadas = []

    def stream(self, *, model, max_tokens, system=None, messages=None):
        self.chamadas.append({"model": model, "max_tokens": max_tokens, "system": system, "messages": messages})
        return FakeStreamManager(self.respostas.get(model))


class FakeAsyncAnthropic:
    ULTIMO = None

    def __init__(self, api_key=None, respostas=None):
        self.api_key = api_key
        self.messages = FakeMessagesClaude(respostas or {})
        FakeAsyncAnthropic.ULTIMO = self


def _claude_client_com(monkeypatch, respostas):
    def _factory(api_key=None):
        return FakeAsyncAnthropic(api_key=api_key, respostas=respostas)
    monkeypatch.setattr(anthropic, "AsyncAnthropic", _factory)


def test_chat_com_provider_claude_usa_haiku_por_padrao(client, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "K")
    _claude_client_com(monkeypatch, {"claude-haiku-5-5": ["resposta ", "do Claude"]})

    r = client.post("/api/gemini/chat", json=corpo("oi, como vai?", provider="claude"), headers=cab())
    assert r.status_code == 200 and r.text == "resposta do Claude"
    chamada = FakeAsyncAnthropic.ULTIMO.messages.chamadas[0]
    assert chamada["model"] == "claude-haiku-5-5"
    assert chamada["system"]
    assert chamada["messages"][-1] == {"role": "user", "content": "oi, como vai?"}


def test_chat_claude_fallback_do_haiku_para_o_sonnet(client, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "K")
    _claude_client_com(monkeypatch, {
        "claude-haiku-5-5": RuntimeError("503 overloaded: service unavailable"),
        "claude-sonnet-5-5": ["funcionou com o Sonnet"],
    })
    r = client.post("/api/gemini/chat", json=corpo("oi", provider="claude"), headers=cab())
    assert r.status_code == 200 and r.text == "funcionou com o Sonnet"
    modelos_tentados = [c["model"] for c in FakeAsyncAnthropic.ULTIMO.messages.chamadas]
    assert modelos_tentados == ["claude-haiku-5-5", "claude-sonnet-5-5"]


def test_chat_sem_chave_claude_devolve_401(client):
    r = client.post("/api/gemini/chat", json=corpo("oi", provider="claude"), headers=cab())
    assert r.status_code == 401


def test_provider_anthropic_e_tratado_como_claude(client, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "K")
    _claude_client_com(monkeypatch, {"claude-haiku-5-5": ["ok"]})
    r = client.post("/api/gemini/chat", json=corpo("oi", provider="anthropic"), headers=cab())
    assert r.status_code == 200 and r.text == "ok"


def test_ai_provider_salvo_no_banco_e_respeitado_quando_request_manda_gemini(client, monkeypatch):
    """Front-end sempre envia provider:"gemini" (o valor salvo é quem decide — ver static/resumo.js)."""
    monkeypatch.setenv("ANTHROPIC_API_KEY", "K")
    _claude_client_com(monkeypatch, {"claude-haiku-5-5": ["veio do Claude salvo"]})
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES ('ai_provider', 'claude')")
    conn.commit()
    conn.close()
    r = client.post("/api/gemini/chat", json=corpo("oi"), headers=cab())  # provider default = "gemini"
    assert r.status_code == 200 and r.text == "veio do Claude salvo"


# ── precedência de chave por provedor (reaproveita o padrão já testado pro Gemini) ───────────────
def test_resolver_credencial_openai_usa_env_padrao(monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "ENV-OPENAI")
    assert ai_chat.resolver_credencial(None, None, "openai") == ("ENV-OPENAI", "padrao")


def test_resolver_credencial_claude_usa_env_padrao(monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "ENV-CLAUDE")
    assert ai_chat.resolver_credencial(None, None, "claude") == ("ENV-CLAUDE", "padrao")


def test_resolver_credencial_chave_do_usuario_prevalece_sobre_env(monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "ENV-OPENAI")
    assert ai_chat.resolver_credencial("CHAVE-DO-USUARIO", "CHAVE-SALVA", "openai") == ("CHAVE-DO-USUARIO", "usuario")
    assert ai_chat.resolver_credencial(None, "CHAVE-SALVA", "openai") == ("CHAVE-SALVA", "salva")


def test_sem_nenhuma_chave_devolve_nenhuma():
    assert ai_chat.resolver_credencial(None, None, "claude") == (None, "nenhuma")


# ── GET /models para Claude (lista fixa) ──────────────────────────────────────
def test_models_claude_devolve_lista_fixa_com_haiku_como_padrao(client):
    r = client.get("/api/gemini/models", headers={**cab(), "X-Provider": "claude", "X-Anthropic-Key": "K"})
    assert r.status_code == 200
    corpo_resp = r.json()
    assert corpo_resp["provider"] == "claude"
    ids = [m["id"] for m in corpo_resp["models"]]
    assert ids[0] == "claude-haiku-5-5" and "★ (Padrão)" in corpo_resp["models"][0]["label"]
    assert "claude-sonnet-5-5" in ids and "claude-opus-5-5" in ids


def test_models_claude_sem_chave_devolve_401(client):
    r = client.get("/api/gemini/models", headers={**cab(), "X-Provider": "claude"})
    assert r.status_code == 401


def test_models_openai_usa_header_proprio_x_openai_key(client, monkeypatch):
    class FakeModelEntry:
        def __init__(self, id_):
            self.id = id_

    class FakeModelsList:
        def list(self):
            return types.SimpleNamespace(data=[FakeModelEntry("gpt-4o"), FakeModelEntry("gpt-4o-mini")])

    class FakeOpenAISync:
        def __init__(self, api_key=None, base_url=None):
            self.api_key = api_key
            self.models = FakeModelsList()

    monkeypatch.setattr(openai, "OpenAI", FakeOpenAISync)
    r = client.get("/api/gemini/models", headers={**cab(), "X-Provider": "openai", "X-OpenAI-Key": "K"})
    assert r.status_code == 200
    ids = [m["id"] for m in r.json()["models"]]
    assert ids == ["gpt-4o", "gpt-4o-mini"]


# ── UI: seletor de provedor e correção do bug "ai_provider fixo em gemini" ────
def test_modal_tem_seletor_de_provedor_para_os_tres():
    html = open("static/resumo.html", encoding="utf-8").read()
    for provider in ("gemini", "openai", "claude"):
        assert f'data-provider="{provider}"' in html
    for secao in ("section-gemini", "section-openai", "section-claude"):
        assert f'id="{secao}"' in html
    assert 'id="input-apikey-openai"' in html and 'id="input-apikey-claude"' in html
    assert 'claude-haiku-5-5' in html  # modelo padrão pré-selecionado do Claude


def test_bug_do_provider_fixo_em_gemini_foi_corrigido():
    js = open("static/resumo.js", encoding="utf-8").read()
    assert "localStorage.setItem('ai_provider', 'gemini')" not in js
    assert "localStorage.setItem('ai_provider', provider)" in js
