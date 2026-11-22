"""Endereços permitidos para clientes compatíveis com a API da OpenAI."""
import os
from urllib.parse import urlsplit

from config import APP_MODE

OPENAI_BASE = "https://api.openai.com/v1"
DEEPSEEK_BASE = "https://api.deepseek.com"
DEEPSEEK_BASES = frozenset({DEEPSEEK_BASE, DEEPSEEK_BASE + "/v1"})
BASES_PUBLICAS = DEEPSEEK_BASES | {
    OPENAI_BASE,
    "https://dashscope-intl.aliyuncs.com/compatible-mode/v1",
    "https://openrouter.ai/api/v1",
}
BASES_LOCAIS = frozenset({"http://localhost:11434/v1", "http://127.0.0.1:11434/v1"})


def _normalizar_base(valor: str) -> str:
    base = valor.strip().rstrip("/")
    if not base or len(base) > 300 or any(c.isspace() or ord(c) < 32 for c in base):
        raise ValueError("Informe uma URL base válida com http:// ou https://.")
    partes = urlsplit(base)
    if (partes.scheme not in ("http", "https") or not partes.hostname
            or partes.username is not None or partes.query or partes.fragment
            or any(c in base for c in ("\\", "?", "#"))):
        raise ValueError("Informe uma URL base válida com http:// ou https://.")
    # A leitura de port também valida portas fora da faixa e valores malformados.
    porta = partes.port
    host = partes.hostname
    if ":" in host:
        host = f"[{host}]"
    padrao = 443 if partes.scheme == "https" else 80
    porta_url = f":{porta}" if porta is not None and porta != padrao else ""
    return f"{partes.scheme}://{host}{porta_url}{partes.path}"


def validar_base_openai(valor: str) -> str:
    """Valida o destino antes que uma rota crie o cliente ou leia a chave."""
    base = _normalizar_base(valor) if valor.strip() else OPENAI_BASE
    if base in BASES_PUBLICAS or (APP_MODE == "desktop" and os.getenv("HEADLESS") != "1" and base in BASES_LOCAIS):
        return base
    permitidas = os.getenv("AI_OPENAI_ALLOWED_BASE_URLS", "").split(",")
    if any(base == _normalizar_base(url) for url in permitidas if url.strip()):
        return base
    raise ValueError("URL base de IA não autorizada neste OrçaCore.")


def modelo_openai_compativel(base: str, modelo: str | None) -> str | None:
    # Usa o padrão atual do DeepSeek e atualiza o atalho legado.
    if base in DEEPSEEK_BASES and modelo in (None, "", "deepseek-chat"):
        return "deepseek-flash"
    return modelo


def criar_cliente_openai(api_key: str, base_url: str = OPENAI_BASE):
    from openai import AsyncOpenAI, DefaultAsyncHttpxClient

    return AsyncOpenAI(api_key=api_key, base_url=base_url,
                       http_client=DefaultAsyncHttpxClient(follow_redirects=False))
