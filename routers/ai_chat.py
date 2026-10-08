"""
routers/ai_chat.py — Integração com Gemini, OpenAI e Claude (Anthropic — TASK-055).
"""
import os
import re
import json
import time
import threading
from collections import defaultdict, deque
from fastapi import APIRouter, Request, HTTPException
from fastapi.responses import PlainTextResponse, StreamingResponse
from database import get_connection
from models import ChatRequest
from config import PROMPT_OBRAS_PATH, PROMPT_PATH, logger
from routers.regras import get_regras
from routers.obras import buscar_obra, listar_visiveis
from services.obras_contexto import menciona_obra, montar_indice, obras_citadas, resumir_obra

router = APIRouter(prefix="/api/gemini", tags=["ai"])

# O frontend envia este texto no header quando o usuário não digitou chave. Ele é "verdadeiro"
# em Python, então precisa ser descartado explicitamente — senão `header or env` nunca chega
# na chave padrão do ambiente.
SENTINELA_CHAVE = "SAVED_IN_BACKEND"

# Limite por usuário só se aplica quando a chave usada é a padrão do sistema (cota compartilhada).
# Em memória e por processo: suficiente para conter abuso, não é contabilidade exata.
RATE_LIMIT_POR_MINUTO = int(os.getenv("AI_RATE_LIMIT_POR_MINUTO", "20") or "20")
_rate_lock = threading.Lock()
_rate_hits: dict = defaultdict(deque)

# Erros que acionam o fallback para o próximo modelo da lista (cota/indisponibilidade), em vez de
# devolver erro direto ao usuário (TASK-054). "invalid_argument"/"multiturn" cobrem o erro "Multiturn
# chat is not enabled for this model" (e variações futuras parecidas) como defesa em profundidade —
# a troca para generate_content_stream (sem sessão) já elimina a causa raiz desse erro específico.
_GATILHOS_FALLBACK = ("429", "quota", "exhausted", "not found", "404", "unavailable", "invalid_argument", "multiturn")


def _ler_configuracao(chave: str, padrao=None):
    conn = get_connection()
    try:
        row = conn.execute("SELECT valor FROM configuracoes WHERE chave = ?", (chave,)).fetchone()
    finally:
        conn.close()
    return row[0] if row else padrao


def _chave_padrao(provider: str = "gemini") -> str:
    """Chave padrão do ambiente para o provedor (TASK-055: OpenAI e Claude ganharam a própria)."""
    if provider == "openai":
        return (os.getenv("OPENAI_API_KEY") or "").strip()
    if provider == "claude":
        return (os.getenv("ANTHROPIC_API_KEY") or "").strip()
    return (os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or "").strip()


def resolver_credencial(header_key: str | None, saved_key: str | None, provider: str = "gemini"):
    """Devolve (chave, origem). Precedência: usuário > salva > padrão do ambiente.

    origem: "usuario" | "salva" | "padrao" | "nenhuma". `provider` escolhe a variável de ambiente
    padrão certa (TASK-055) — "gemini" (padrão, compatível com as chamadas existentes), "openai" ou
    "claude".
    """
    chave = (header_key or "").strip()
    if chave and chave != SENTINELA_CHAVE:
        return chave, "usuario"
    if saved_key:
        return saved_key, "salva"
    padrao = _chave_padrao(provider)
    if padrao:
        return padrao, "padrao"
    return None, "nenhuma"


def origem_chave_ativa() -> str:
    """Origem da chave que seria usada sem header (diagnóstico; nunca expõe o valor)."""
    return resolver_credencial(None, _ler_configuracao("gemini_api_key"))[1]


def _checar_rate_limit(request: Request):
    user = getattr(request.state, "user", None) or {}
    ident = user.get("email") or user.get("sub") or (request.client.host if request.client else "anon")
    agora = time.monotonic()
    with _rate_lock:
        janela = _rate_hits[ident]
        while janela and agora - janela[0] > 60:
            janela.popleft()
        if len(janela) >= RATE_LIMIT_POR_MINUTO:
            raise HTTPException(
                status_code=429,
                detail=f"Limite de {RATE_LIMIT_POR_MINUTO} mensagens por minuto atingido com a chave padrão. "
                       "Aguarde um instante ou informe sua própria chave.",
            )
        janela.append(agora)


@router.get("/models")
def get_gemini_models(request: Request):
    provider = request.headers.get("X-Provider", "gemini")
    if provider == "anthropic":
        provider = "claude"
    openai_base_url = request.headers.get("X-OpenAI-Base-URL", "")

    if provider == "openai":
        api_key, _origem = resolver_credencial(
            request.headers.get("X-OpenAI-Key"), _ler_configuracao("openai_api_key"), "openai"
        )
        if not api_key:
            raise HTTPException(status_code=401, detail="API Key não encontrada.")
        try:
            from openai import OpenAI
            base = openai_base_url or "https://api.openai.com/v1"
            client = OpenAI(api_key=api_key if api_key else "ollama", base_url=base)
            models_raw = [m.id for m in client.models.list().data]
            modelos_formatados = [{"id": m, "label": m} for m in sorted(models_raw)]
            return {"models": modelos_formatados, "provider": "openai"}
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))

    if provider == "claude":
        api_key, _origem = resolver_credencial(
            request.headers.get("X-Anthropic-Key"), _ler_configuracao("anthropic_api_key"), "claude"
        )
        if not api_key:
            raise HTTPException(status_code=401, detail="API Key não encontrada.")
        # Anthropic não tem um endpoint de listagem tão aberto quanto Gemini/OpenAI (TASK-055) —
        # lista fixa da família atual é suficiente; o Haiku é o padrão pedido pelo usuário.
        modelos_formatados = [
            {"id": "claude-haiku-5-5", "label": "claude-haiku-5-5 ★ (Padrão)"},
            {"id": "claude-sonnet-5-5", "label": "claude-sonnet-5-5 (Reserva)"},
            {"id": "claude-opus-5-5", "label": "claude-opus-5-5"},
        ]
        return {"models": modelos_formatados, "provider": "claude"}

    # Gemini (padrão)
    api_key, _origem = resolver_credencial(
        request.headers.get("X-Gemini-Key"), _ler_configuracao("gemini_api_key")
    )
    if not api_key:
        raise HTTPException(status_code=401, detail="API Key não encontrada.")
    try:
        from google import genai as _genai
        _client = _genai.Client(api_key=api_key)
        preferencias_gemini = ["gemini-3.1-flash-lite", "gemini-3.6-flash", "gemini-3.5-flash"]
        models_page = _client.models.list()
        modelos_raw = []
        for m in models_page:
            name = m.name.split("/")[-1] if "/" in m.name else m.name
            if name.startswith("gemini"):
                modelos_raw.append(name)
        modelos_formatados = []
        for m in modelos_raw:
            label = m
            if m == "gemini-3.1-flash-lite":
                label = f"{m} ★ (Padrão)"
            elif m in preferencias_gemini:
                label = f"{m} (Reserva)"
            modelos_formatados.append({"id": m, "label": label})
            
        ids = [m["id"] for m in modelos_formatados]
        if "gemini-3.1-flash-lite" not in ids:
            modelos_formatados.insert(0, {"id": "gemini-3.1-flash-lite", "label": "gemini-3.1-flash-lite ★ (Padrão)"})
        return {"models": modelos_formatados, "provider": "gemini"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


def _contexto_obras(req, request):
    """(texto extra do usuário, bloco de instruções) sobre as obras salvas — TASK-028. Vazio quando o pedido não menciona
    obras: nenhuma consulta ao banco e o prompt final fica idêntico ao de antes."""
    if not menciona_obra(req.prompt):
        return "", ""
    try:
        user = getattr(request.state, "user", None)
        obras = listar_visiveis(user, req.projeto_codigo) if user else []   # próprias + públicas de outros (TASK-057)
        extra = montar_indice(obras)
        for o in obras_citadas(req.prompt, obras):
            completa = buscar_obra(user, o["id"], req.projeto_codigo)
            if completa:
                extra += "\n\n" + resumir_obra(completa)
        instrucoes = ""
        if os.path.exists(PROMPT_OBRAS_PATH):
            with open(PROMPT_OBRAS_PATH, "r", encoding="utf-8") as f:
                instrucoes = f.read()
        return extra, instrucoes
    except Exception as e:  # o contexto de obras é um extra: nunca derruba o chat
        logger.warning(f"Erro ao montar contexto de obras: {e}")
        return "", ""


@router.post("/chat")
async def gemini_chat(req: ChatRequest, request: Request):
    provider = req.provider or "gemini"
    if provider == "anthropic":
        provider = "claude"
    openai_base_url = req.openai_base_url or ""

    # Headers e configuração salva são por provedor (TASK-055), mesma precedência de sempre
    # (usuário > salva > padrão do ambiente), com os nomes já usados pro Gemini preservados.
    header_por_provider = {
        "gemini": request.headers.get("X-Gemini-Key"),
        "openai": request.headers.get("X-OpenAI-Key"),
        "claude": request.headers.get("X-Anthropic-Key"),
    }
    custom_model_header_por_provider = {
        "gemini": request.headers.get("X-Gemini-Model"),
        "openai": request.headers.get("X-OpenAI-Model"),
        "claude": request.headers.get("X-Anthropic-Model"),
    }
    chave_config_por_provider = {
        "gemini": "gemini_api_key", "openai": "openai_api_key", "claude": "anthropic_api_key",
    }
    modelo_config_por_provider = {
        "gemini": "gemini_model", "openai": "openai_model", "claude": "anthropic_model",
    }

    conn = get_connection()
    cursor = conn.cursor()
    saved_provider = _ler_configuracao("ai_provider", "gemini")
    saved_base_url = _ler_configuracao("openai_base_url", "")

    if provider == "gemini" and saved_provider != "gemini":
        provider = "claude" if saved_provider == "anthropic" else saved_provider
    if provider not in ("gemini", "openai", "claude"):
        provider = "gemini"
    if not openai_base_url and saved_base_url:
        openai_base_url = saved_base_url

    chave_config = chave_config_por_provider[provider]
    modelo_config = modelo_config_por_provider[provider]
    saved_key = _ler_configuracao(chave_config)
    saved_model = _ler_configuracao(modelo_config)

    api_key, origem_chave = resolver_credencial(header_por_provider[provider], saved_key, provider)
    if not api_key:
        conn.close()
        raise HTTPException(status_code=401, detail="API Key não fornecida, não salva e sem chave padrão do sistema.")
    # Só a chave digitada pelo usuário é persistida; a padrão vem sempre do ambiente.
    if origem_chave == "usuario":
        cursor.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES (?, ?)", (chave_config, api_key))
        conn.commit()

    custom_model = custom_model_header_por_provider[provider]
    if custom_model == SENTINELA_CHAVE:
        custom_model = saved_model
    elif custom_model is not None:
        cursor.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES (?, ?)", (modelo_config, custom_model))
        conn.commit()
    conn.close()

    # Fast-Path
    fast_match = re.match(
        r'^(adicionar|add|adc|coloca|inserir|remover|rem|tira|excluir)\s+([\d.,]+)\s+(.+)$',
        req.prompt.strip(), re.IGNORECASE
    )
    if fast_match and not menciona_obra(req.prompt):   # "adicionar 2 obra X" vai para a IA (TASK-028)
        acao = fast_match.group(1).lower()
        qtd  = fast_match.group(2).replace(",", ".")
        ativo = fast_match.group(3).upper().strip()
        op = "R" if acao.startswith("rem") or acao in ["tira", "excluir"] else "I"
        fake_json = {"outros": [{"ativo": f"{qtd}-{ativo}", "operacao": op}]}
        return PlainTextResponse(f"*(Fast-Path)*\n```json\n{json.dumps(fake_json, indent=2)}\n```")

    if origem_chave == "padrao":
        _checar_rate_limit(request)

    prompt_lower = req.prompt.lower()
    palavras_contexto = [
        'analis', 'alterar', 'editar', 'excluir', 'remover', 'substituir',
        'tudo', 'todas', 'tabela', 'leia', 'leitura', 'completo', 'lista',
        'quantos', 'total', 'verifique', 'cheque', 'corrig'
    ]
    precisa_contexto = any(kw in prompt_lower for kw in palavras_contexto)
    if precisa_contexto and req.table_context and req.table_context.strip():
        prompt_final = req.prompt.strip() + "\n\n" + req.table_context.strip()
    else:
        prompt_final = req.prompt.strip()

    obras_extra, obras_instrucoes = _contexto_obras(req, request)
    if obras_extra:
        prompt_final += "\n\n" + obras_extra

    system_instruction = "Você é um especialista em redes elétricas."
    if os.path.exists(PROMPT_PATH):
        with open(PROMPT_PATH, "r", encoding="utf-8") as f:
            system_instruction = f.read()

    try:
        regras = get_regras(projeto_codigo=req.projeto_codigo)
        if regras:
            regras_txt = "\n".join(f"- {r['conteudo']}" for r in regras)
            system_instruction += f"\n\n🔹 REGRAS APRENDIDAS:\n{regras_txt}"
    except Exception as e:
        logger.warning(f"Erro ao carregar regras: {e}")

    if obras_instrucoes:
        system_instruction += "\n\n" + obras_instrucoes

    if provider == "gemini":
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=api_key)

        # gemini-1.5-flash/pro: desligados desde 2025 (404). gemini-2.5-flash: desligamento
        # anunciado pela Google para 16/10/2026 — removidos da lista (TASK-054).
        preferencias = [
            "gemini-3.1-flash-lite",
            "gemini-3.6-flash",
            "gemini-3.5-flash",
        ]
        modelos = [custom_model] if custom_model else preferencias
        if len(req.prompt.split()) < 15 and not precisa_contexto:
            modelos = ["gemini-3.1-flash-lite"] + [m for m in modelos if m != "gemini-3.1-flash-lite"]

        sdk_history = []
        for msg in req.history:
            parts = msg.get("parts", [])
            if parts:
                texto = parts[0].get("text", "") if isinstance(parts[0], dict) else str(parts[0])
                sdk_history.append({"role": msg.get("role", "user"), "parts": [{"text": texto}]})
        # Chamada única com o histórico embutido em `contents` (TASK-054), em vez da API de sessão
        # (`chats.create`/`send_message_stream`), que alguns modelos recusam com "Multiturn chat is
        # not enabled for this model". `generate_content_stream` não tem essa restrição porque não
        # depende de sessão nenhuma — é a mesma chamada única usada em `chamar_gemini`, em streaming.
        contents = sdk_history + [{"role": "user", "parts": [{"text": prompt_final}]}]

        async def _stream_gemini():
            last_err = None
            for modelo in modelos:
                try:
                    response = await client.aio.models.generate_content_stream(
                        model=modelo,
                        contents=contents,
                        config=types.GenerateContentConfig(system_instruction=system_instruction),
                    )
                    async for chunk in response:
                        if chunk.text:
                            yield chunk.text
                    return
                except Exception as e:
                    err = str(e).lower()
                    if any(k in err for k in _GATILHOS_FALLBACK):
                        last_err = e
                        logger.warning(f"Gemini {modelo} indisponível. Tentando fallback...")
                        continue
                    yield f"\n[ERRO] {str(e)}"
                    return
            msg = str(last_err) if last_err else "Nenhum modelo Gemini disponível."
            yield f"\n[ERRO] {msg}"

        return StreamingResponse(_stream_gemini(), media_type="text/plain")

    if provider == "openai":
        from openai import AsyncOpenAI
        base = openai_base_url or "https://api.openai.com/v1"
        client = AsyncOpenAI(api_key=api_key, base_url=base)

        preferencias = ["gpt-4o-mini", "gpt-4o"]
        modelos = [custom_model] if custom_model else preferencias

        mensagens = [{"role": "system", "content": system_instruction}]
        for msg in req.history:
            parts = msg.get("parts", [])
            texto = (parts[0].get("text", "") if isinstance(parts[0], dict) else str(parts[0])) if parts else ""
            mensagens.append({"role": "assistant" if msg.get("role") == "model" else "user", "content": texto})
        mensagens.append({"role": "user", "content": prompt_final})

        async def _stream_openai():
            last_err = None
            for modelo in modelos:
                try:
                    stream = await client.chat.completions.create(model=modelo, messages=mensagens, stream=True)
                    async for chunk in stream:
                        delta = chunk.choices[0].delta.content if chunk.choices else None
                        if delta:
                            yield delta
                    return
                except Exception as e:
                    err = str(e).lower()
                    if any(k in err for k in _GATILHOS_FALLBACK):
                        last_err = e
                        logger.warning(f"OpenAI {modelo} indisponível. Tentando fallback...")
                        continue
                    yield f"\n[ERRO] {str(e)}"
                    return
            msg = str(last_err) if last_err else "Nenhum modelo OpenAI disponível."
            yield f"\n[ERRO] {msg}"

        return StreamingResponse(_stream_openai(), media_type="text/plain")

    if provider == "claude":
        from anthropic import AsyncAnthropic
        client = AsyncAnthropic(api_key=api_key)

        # Haiku primeiro: modelo mais rápido/barato da linha, padrão pedido pelo usuário (TASK-055).
        # Sonnet como reserva de qualidade, não como padrão.
        preferencias = ["claude-haiku-5-5", "claude-sonnet-5-5"]
        modelos = [custom_model] if custom_model else preferencias

        mensagens = []
        for msg in req.history:
            parts = msg.get("parts", [])
            texto = (parts[0].get("text", "") if isinstance(parts[0], dict) else str(parts[0])) if parts else ""
            mensagens.append({"role": "assistant" if msg.get("role") == "model" else "user", "content": texto})
        mensagens.append({"role": "user", "content": prompt_final})

        async def _stream_claude():
            last_err = None
            for modelo in modelos:
                try:
                    async with client.messages.stream(
                        model=modelo, max_tokens=4096, system=system_instruction, messages=mensagens,
                    ) as stream:
                        async for texto in stream.text_stream:
                            yield texto
                    return
                except Exception as e:
                    err = str(e).lower()
                    if any(k in err for k in _GATILHOS_FALLBACK):
                        last_err = e
                        logger.warning(f"Claude {modelo} indisponível. Tentando fallback...")
                        continue
                    yield f"\n[ERRO] {str(e)}"
                    return
            msg = str(last_err) if last_err else "Nenhum modelo Claude disponível."
            yield f"\n[ERRO] {msg}"

        return StreamingResponse(_stream_claude(), media_type="text/plain")

    return PlainTextResponse(f"[ERRO] Provider '{provider}' não suportado.")
