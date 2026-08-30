"""
services/supabase_client.py — Cliente centralizado para acesso ao Supabase.

Singleton thread-safe. Lê credenciais EXCLUSIVAMENTE de variáveis de ambiente (.env).
Nunca armazena chaves no código-fonte.
"""
import os
import threading
from supabase import create_client, Client
from config import logger


_supabase_client: Client | None = None
_client_lock = threading.Lock()
_client_initialized = False


def get_supabase() -> Client | None:
    """
    Retorna o cliente Supabase singleton.
    Lê SUPABASE_URL e SUPABASE_KEY exclusivamente de variáveis de ambiente.
    Retorna None se as credenciais não estiverem configuradas.
    """
    global _supabase_client, _client_initialized

    if _client_initialized:
        return _supabase_client

    with _client_lock:
        # Double-check após adquirir o lock
        if _client_initialized:
            return _supabase_client

        url = os.environ.get("SUPABASE_URL", "").strip()
        key = os.environ.get("SUPABASE_KEY", "").strip()

        if not url or not key:
            logger.warning(
                "Supabase não configurado: defina SUPABASE_URL e SUPABASE_KEY no arquivo .env"
            )
            _client_initialized = True
            return None

        try:
            _supabase_client = create_client(url, key)
            logger.info("Cliente Supabase inicializado com sucesso.")
        except Exception as e:
            logger.error(f"Erro ao inicializar cliente Supabase: {e}")
            _supabase_client = None

        _client_initialized = True
        return _supabase_client


def reset_supabase_client():
    """Reseta o singleton (útil para reconexão após mudança de credenciais)."""
    global _supabase_client, _client_initialized
    with _client_lock:
        _supabase_client = None
        _client_initialized = False
