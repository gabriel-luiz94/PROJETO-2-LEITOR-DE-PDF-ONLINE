"""
services/supabase_client.py — Cliente centralizado para acesso ao Supabase.

Singleton thread-safe. Lê credenciais EXCLUSIVAMENTE de variáveis de ambiente (.env).
Nunca armazena chaves no código-fonte.

TASK-039: antes, o cliente (e o httpx.Client/pool HTTP por baixo) era criado uma única vez por
processo e nunca recriado — se o Supabase/PostgREST fechasse uma conexão ociosa (comportamento
normal depois de alguns minutos sem uso), o processo nunca mais reconectava sozinho até reiniciar.
Duas defesas, combinadas:
  1. TTL proativo: get_supabase() recria o cliente sozinho depois de _TTL_SEGUNDOS, mesmo que
     nenhuma chamada tenha falhado ainda (evita reusar uma conexão que pode já estar morta do lado
     do servidor).
  2. Reset reativo: registrar_falha() (chamada pelos routers depois de uma falha real) reseta o
     cliente imediatamente quando a falha parece ser de rede/conexão — não quando é erro de
     autenticação/permissão/dados, que resetar não resolveria.
"""
import os
import threading
import time
from supabase import create_client, Client
from config import logger

_TTL_SEGUNDOS = 180  # tempo máximo de vida do cliente antes de ser recriado por precaução


_supabase_client: Client | None = None
_client_lock = threading.Lock()
_client_initialized = False
_client_criado_em = 0.0


def _expirado() -> bool:
    return _client_initialized and _supabase_client is not None and (time.monotonic() - _client_criado_em) > _TTL_SEGUNDOS


def get_supabase() -> Client | None:
    """
    Retorna o cliente Supabase singleton.
    Lê SUPABASE_URL e SUPABASE_KEY exclusivamente de variáveis de ambiente.
    Retorna None se as credenciais não estiverem configuradas.
    """
    global _supabase_client, _client_initialized, _client_criado_em

    if _client_initialized and not _expirado():
        return _supabase_client

    with _client_lock:
        # Double-check após adquirir o lock
        if _client_initialized and not _expirado():
            return _supabase_client
        if _client_initialized and _expirado():
            logger.info(f"Cliente Supabase expirou (TTL de {_TTL_SEGUNDOS}s); recriando por precaução.")
            _supabase_client = None
            _client_initialized = False

        url = os.environ.get("SUPABASE_URL", "").strip()
        key = os.environ.get("SUPABASE_KEY", "").strip()

        if not url or not key:
            logger.warning(
                "Supabase não configurado: defina SUPABASE_URL e SUPABASE_KEY no arquivo .env"
            )
            _client_initialized = True
            _client_criado_em = time.monotonic()
            return None

        try:
            _supabase_client = create_client(url, key)
            logger.info("Cliente Supabase inicializado com sucesso.")
        except Exception as e:
            logger.error(f"Erro ao inicializar cliente Supabase: {e}")
            _supabase_client = None

        _client_initialized = True
        _client_criado_em = time.monotonic()
        return _supabase_client


def reset_supabase_client():
    """Reseta o singleton (reconexão após mudança de credenciais, ou após falha de rede detectada)."""
    global _supabase_client, _client_initialized
    with _client_lock:
        _supabase_client = None
        _client_initialized = False


def _eh_erro_de_rede(exc: BaseException) -> bool:
    """True quando a falha parece ser de conexão/timeout — não de autenticação, permissão (RLS) ou
    dados, que resetar o cliente não resolveria e só esconderia o erro real."""
    try:
        import httpx
        if isinstance(exc, (httpx.TransportError, httpx.TimeoutException)):
            return True
    except ImportError:
        pass
    return isinstance(exc, (ConnectionError, OSError, TimeoutError))


def registrar_falha(contexto: str, exc: BaseException) -> None:
    """Chame no `except` de toda chamada ao Supabase (TASK-039/041): loga a falha — nunca mais
    silenciosa — e, se parecer erro de rede, reseta o cliente para a PRÓXIMA chamada ter uma chance
    real de reconectar, em vez de repetir a mesma falha para sempre."""
    logger.warning(f"Falha ao acessar o Supabase ({contexto}): {exc}")
    if _eh_erro_de_rede(exc):
        reset_supabase_client()
