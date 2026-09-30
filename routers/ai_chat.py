"""
routers/ai_chat.py — Integração com Gemini e OpenAI.
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
from config import PROMPT_PATH, logger
from routers.regras import get_regras

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


def _ler_configuracao(chave: str, padrao=None):
    conn = get_connection()
    try:
        row = conn.execute("SELECT valor FROM configuracoes WHERE chave = ?", (chave,)).fetchone()
    finally:
        conn.close()
    return row[0] if row else padrao


def _chave_padrao() -> str:
    return (os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or "").strip()


def resolver_credencial(header_key: str | None, saved_key: str | None):
    """Devolve (chave, origem). Precedência: usuário > salva > padrão do ambiente.

    origem: "usuario" | "salva" | "padrao" | "nenhuma".
    """
    chave = (header_key or "").strip()
    if chave and chave != SENTINELA_CHAVE:
        return chave, "usuario"
    if saved_key:
        return saved_key, "salva"
    padrao = _chave_padrao()
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
    openai_base_url = request.headers.get("X-OpenAI-Base-URL", "")

    api_key, _origem = resolver_credencial(
        request.headers.get("X-Gemini-Key"), _ler_configuracao("gemini_api_key")
    )
    if not api_key:
        raise HTTPException(status_code=401, detail="API Key não encontrada.")

    if provider == "openai":
        try:
            from openai import OpenAI
            base = openai_base_url or "https://api.openai.com/v1"
            client = OpenAI(api_key=api_key if api_key else "ollama", base_url=base)
            models_raw = [m.id for m in client.models.list().data]
            modelos_formatados = [{"id": m, "label": m} for m in sorted(models_raw)]
            return {"models": modelos_formatados, "provider": "openai"}
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))

    # Gemini (padrão)
    try:
        from google import genai as _genai
        _client = _genai.Client(api_key=api_key)
        preferencias_gemini = ["gemini-3.1-flash-lite", "gemini-3.6-flash", "gemini-2.5-flash", "gemini-1.5-flash"]
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


@router.post("/chat")
async def gemini_chat(req: ChatRequest, request: Request):
    header_key = request.headers.get("X-Gemini-Key")
    custom_model = request.headers.get("X-Gemini-Model")
    provider = req.provider or "gemini"
    openai_base_url = req.openai_base_url or ""

    conn = get_connection()
    cursor = conn.cursor()
    saved_key = _ler_configuracao("gemini_api_key")
    saved_model = _ler_configuracao("gemini_model")
    saved_provider = _ler_configuracao("ai_provider", "gemini")
    saved_base_url = _ler_configuracao("openai_base_url", "")

    api_key, origem_chave = resolver_credencial(header_key, saved_key)
    if not api_key:
        conn.close()
        raise HTTPException(status_code=401, detail="API Key não fornecida, não salva e sem chave padrão do sistema.")
    # Só a chave digitada pelo usuário é persistida; a padrão vem sempre do ambiente.
    if origem_chave == "usuario":
        cursor.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES ('gemini_api_key', ?)", (api_key,))
        conn.commit()

    if custom_model == SENTINELA_CHAVE:
        custom_model = saved_model
    elif custom_model is not None:
        cursor.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES ('gemini_model', ?)", (custom_model,))
        conn.commit()

    if provider == "gemini" and saved_provider != "gemini":
        provider = saved_provider
    if not openai_base_url and saved_base_url:
        openai_base_url = saved_base_url
    conn.close()

    # Fast-Path
    fast_match = re.match(
        r'^(adicionar|add|adc|coloca|inserir|remover|rem|tira|excluir)\s+([\d.,]+)\s+(.+)$',
        req.prompt.strip(), re.IGNORECASE
    )
    if fast_match:
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

    if provider == "gemini":
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=api_key)

        preferencias = [
            "gemini-3.1-flash-lite",
            "gemini-2.5-flash",
            "gemini-3.6-flash",
            "gemini-1.5-flash",
            "gemini-1.5-pro",
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

        async def _stream_gemini():
            last_err = None
            for modelo in modelos:
                try:
                    chat = client.aio.chats.create(
                        model=modelo,
                        config=types.GenerateContentConfig(system_instruction=system_instruction),
                        history=sdk_history
                    )
                    response = await chat.send_message_stream(prompt_final)
                    async for chunk in response:
                        if chunk.text:
                            yield chunk.text
                    return
                except Exception as e:
                    err = str(e).lower()
                    if any(k in err for k in ["429", "quota", "exhausted", "not found", "404", "unavailable"]):
                        last_err = e
                        logger.warning(f"Gemini {modelo} indisponível. Tentando fallback...")
                        continue
                    yield f"\n[ERRO] {str(e)}"
                    return
            msg = str(last_err) if last_err else "Nenhum modelo Gemini disponível."
            yield f"\n[ERRO] {msg}"

        return StreamingResponse(_stream_gemini(), media_type="text/plain")

    return PlainTextResponse("[ERRO] Provider não suportado neste momento.")
