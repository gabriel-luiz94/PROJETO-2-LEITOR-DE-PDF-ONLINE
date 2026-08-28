"""
routers/ai_chat.py — Integração com Gemini e OpenAI.
"""
import os
import re
import json
from fastapi import APIRouter, Request, HTTPException
from fastapi.responses import PlainTextResponse, StreamingResponse
from database import get_connection
from models import ChatRequest
from config import PROMPT_PATH, logger

router = APIRouter(prefix="/api/gemini", tags=["ai"])


@router.get("/models")
def get_gemini_models(request: Request):
    api_key = request.headers.get("X-Gemini-Key") or os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    provider = request.headers.get("X-Provider", "gemini")
    openai_base_url = request.headers.get("X-OpenAI-Base-URL", "")

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT valor FROM configuracoes WHERE chave = 'gemini_api_key'")
    row = cursor.fetchone()
    saved_key = row[0] if row else None
    conn.close()

    if not api_key or api_key == "SAVED_IN_BACKEND":
        if saved_key:
            api_key = saved_key
        else:
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
    api_key = request.headers.get("X-Gemini-Key") or os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    custom_model = request.headers.get("X-Gemini-Model")
    provider = req.provider or "gemini"
    openai_base_url = req.openai_base_url or ""

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT valor FROM configuracoes WHERE chave = 'gemini_api_key'")
    row = cursor.fetchone()
    saved_key = row[0] if row else None
    cursor.execute("SELECT valor FROM configuracoes WHERE chave = 'gemini_model'")
    row_model = cursor.fetchone()
    saved_model = row_model[0] if row_model else None
    cursor.execute("SELECT valor FROM configuracoes WHERE chave = 'ai_provider'")
    row_prov = cursor.fetchone()
    saved_provider = row_prov[0] if row_prov else "gemini"
    cursor.execute("SELECT valor FROM configuracoes WHERE chave = 'openai_base_url'")
    row_base = cursor.fetchone()
    saved_base_url = row_base[0] if row_base else ""

    # Só salva se não estiver em branco E não for "SAVED_IN_BACKEND"
    # No futuro, se removermos api keys locais (migrando pro Supabase), adaptaremos aqui
    if api_key and api_key != "SAVED_IN_BACKEND":
        cursor.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES ('gemini_api_key', ?)", (api_key,))
        conn.commit()
    elif saved_key:
        api_key = saved_key
    else:
        conn.close()
        raise HTTPException(status_code=401, detail="API Key não fornecida ou não salva internamente.")

    if custom_model == "SAVED_IN_BACKEND":
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
        conn2 = get_connection()
        cursor2 = conn2.cursor()
        cursor2.execute("SELECT conteudo FROM regras")
        regras = cursor2.fetchall()
        conn2.close()
        if regras:
            regras_txt = "\n".join(f"- {r[0]}" for r in regras)
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
