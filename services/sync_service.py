"""
services/sync_service.py — Lógica de Sincronização Offline-First (Supabase <-> SQLite)
"""
from services.supabase_client import get_supabase
from database import get_connection, get_row_connection
from config import logger


def sync_tabela_master():
    """
    Baixa a tabela mestre (Admin) do Supabase e atualiza o SQLite local.
    """
    supabase = get_supabase()
    if not supabase:
        logger.info("Sync Master cancelado: Cliente Supabase não configurado/offline.")
        return False

    try:
        # Busca todas as linhas da tabela master no Supabase
        # Limitado a 5000 inicialmente; se o banco for muito grande, precisa paginar
        response = supabase.table("tabela_orcamento_master").select("*").limit(5000).execute()
        master_rows = response.data

        if not master_rows:
            return True # Vazia, nada a fazer

        conn = get_connection()
        cursor = conn.cursor()
        
        # Inicia uma transação local para garantir consistência
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
        # Tenta fazer rollback caso a transação tenha falhado
        try:
            conn = get_connection()
            conn.cursor().execute("ROLLBACK")
            conn.close()
        except:
            pass
        return False


def get_merged_orcamento(user_id: str = None) -> list[dict]:
    """
    Retorna a tabela de orçamento fundida:
    - Prioridade 1: tabela_orcamento_master (O que o admin subir sempre sobrepõe).
    - Prioridade 2: tabela_orcamento (Alterações locais do usuário, se não houver código correspondente no Master).
    """
    conn = get_row_connection()
    cursor = conn.cursor()
    
    # Busca tabela Mestre
    cursor.execute("SELECT * FROM tabela_orcamento_master")
    master_rows = [dict(r) for r in cursor.fetchall()]
    
    # Busca tabela do Usuário (se tiver user_id, busca só dele; senão busca tudo que estiver local, para compatibilidade offline/single-user)
    if user_id:
        cursor.execute("SELECT * FROM tabela_orcamento WHERE user_id = ? OR user_id IS NULL", (user_id,))
    else:
        cursor.execute("SELECT * FROM tabela_orcamento")
    
    user_rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    
    # A fusão ("Merge"):
    # Chave de sobreposição ideal é o "código". Mas alguns ativos não têm código (são identificados por 'ativo' e 'componente').
    # Vamos usar uma tupla (codigo, ativo) como chave de sobreposição.
    
    merged_dict = {}
    
    # 1. Adiciona os dados do usuário (Prioridade Baixa)
    for row in user_rows:
        cod = (row.get("codigo") or "").strip()
        atv = (row.get("ativo") or "").strip()
        key = f"{cod}##{atv}"
        merged_dict[key] = row
        
    # 2. Adiciona os dados Mestre (Prioridade Alta - Sobrescreve)
    for row in master_rows:
        cod = (row.get("codigo") or "").strip()
        atv = (row.get("ativo") or "").strip()
        key = f"{cod}##{atv}"
        # Sobrescreve o que o usuário tinha com o do Master
        merged_dict[key] = row
        
    return list(merged_dict.values())
