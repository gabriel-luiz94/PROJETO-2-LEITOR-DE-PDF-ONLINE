"""
services/sync_service.py — Lógica de Sincronização Offline-First (Supabase <-> SQLite)
"""
from services.supabase_client import get_supabase
from services.offline_queue import process_queue
from database import get_connection, get_row_connection
from config import logger


def sync_tabela_master():
    """
    Baixa a tabela mestre (Admin) do Supabase e atualiza o SQLite local.
    Suporta paginação para baixar mais de 1000 registros do Supabase.
    """
    supabase = get_supabase()
    if not supabase:
        logger.info("Sync Master cancelado: Cliente Supabase não configurado/offline.")
        return False

    try:
        process_queue()

        # Busca todas as páginas do Supabase (1000 por lote)
        master_rows = []
        page_size = 1000
        start = 0
        while True:
            res = supabase.table("tabela_orcamento_master").select("*").range(start, start + page_size - 1).execute()
            data = res.data or []
            master_rows.extend(data)
            if len(data) < page_size:
                break
            start += page_size

        conn = get_connection()
        cursor = conn.cursor()
        
        cursor.execute("BEGIN TRANSACTION")
        cursor.execute("DELETE FROM tabela_orcamento_master")
        
        for row in master_rows:
            cursor.execute('''
                INSERT INTO tabela_orcamento_master 
                (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                row.get("ativo", ""),
                row.get("desc_ativo", ""),
                row.get("componente", ""),
                row.get("projeto", ""),
                row.get("mdo", ""),
                row.get("codigo", ""),
                row.get("desc_codigo", ""),
                row.get("fator_i", 0.0),
                row.get("fator_r", 0.0),
                row.get("filtro", "")
            ))
            
        cursor.execute("COMMIT")
        conn.close()
        logger.info(f"Sync Master concluído. {len(master_rows)} registros atualizados localmente.")
        return True

    except Exception as e:
        logger.error(f"Erro ao sincronizar tabela master: {e}")
        try:
            conn = get_connection()
            conn.cursor().execute("ROLLBACK")
            conn.close()
        except:
            pass
        return False


def get_merged_orcamento(user_id: str = None) -> list[dict]:
    """
    Retorna a tabela de orçamento do Supabase (tabela_orcamento_master)
    com fallback para a tabela local se o Supabase ainda não tiver sido sincronizado.
    """
    conn = get_row_connection()
    cursor = conn.cursor()
    
    # 1. Tabela Master oficial sincronizada do Supabase
    cursor.execute("SELECT * FROM tabela_orcamento_master")
    master_rows = [dict(r) for r in cursor.fetchall()]
    
    if master_rows and len(master_rows) > 0:
        conn.close()
        return master_rows
    
    # 2. Fallback para tabela local se a master estiver vazia
    if user_id:
        cursor.execute("SELECT * FROM tabela_orcamento WHERE user_id = ? OR user_id IS NULL", (user_id,))
    else:
        cursor.execute("SELECT * FROM tabela_orcamento")
    
    user_rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return user_rows
