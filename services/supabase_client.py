"""
services/supabase_client.py — Cliente centralizado para acesso ao Supabase.
"""
import os
from supabase import create_client, Client
from config import logger
from database import get_connection

def get_supabase() -> Client | None:
    """
    Inicializa e retorna o cliente Supabase.
    Tenta pegar as chaves do .env (variáveis de ambiente).
    Se não encontrar, tenta buscar na tabela configuracoes local.
    Retorna None se as chaves não estiverem disponíveis.
    """
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_KEY")

    if not url or not key:
        try:
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("SELECT valor FROM configuracoes WHERE chave = 'supabase_url'")
            row_url = cur.fetchone()
            if row_url: url = row_url[0]
            
            cur.execute("SELECT valor FROM configuracoes WHERE chave = 'supabase_key'")
            row_key = cur.fetchone()
            if row_key: key = row_key[0]
            conn.close()
        except Exception as e:
            logger.warning(f"Erro ao buscar credenciais do Supabase no banco local: {e}")

    if url and key:
        try:
            return create_client(url, key)
        except Exception as e:
            logger.error(f"Erro ao inicializar cliente Supabase: {e}")
            return None
    return None
