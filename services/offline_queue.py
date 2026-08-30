"""
services/offline_queue.py — Fila persistente para operações offline.

Registra operações CRUD feitas quando o sistema está sem internet ou quando o Supabase
falha. Tenta re-executar as operações automaticamente (Replay) em background.
"""
import json
import threading
import time
from database import get_connection, get_row_connection
from services.supabase_client import get_supabase
from config import logger


def enqueue_operation(tabela: str, operacao: str, registro_id: str, dados: dict):
    """
    Adiciona uma operação à fila offline.
    operacao: 'INSERT', 'UPDATE', 'DELETE'
    """
    conn = get_connection()
    cur = conn.cursor()
    try:
        dados_json = json.dumps(dados) if dados else None
        cur.execute('''
            INSERT INTO sync_log (tabela, operacao, registro_id, dados_json)
            VALUES (?, ?, ?, ?)
        ''', (tabela, operacao, str(registro_id), dados_json))
        conn.commit()
        logger.info(f"Operação {operacao} em {tabela} enfileirada para sync (ID: {registro_id})")
    except Exception as e:
        logger.error(f"Erro ao enfileirar operação offline: {e}")
    finally:
        conn.close()


def process_queue():
    """
    Tenta processar todas as operações pendentes na fila.
    """
    supabase = get_supabase()
    if not supabase:
        return False

    conn = get_row_connection()
    cur = conn.cursor()
    cur.execute('''
        SELECT * FROM sync_log 
        WHERE sincronizado = 0 
        ORDER BY id ASC 
        LIMIT 50
    ''')
    pendentes = cur.fetchall()
    conn.close()

    if not pendentes:
        return True

    logger.info(f"Processando {len(pendentes)} operações pendentes na fila offline...")
    sucesso_total = True

    for p in pendentes:
        log_id = p["id"]
        tabela = p["tabela"]
        operacao = p["operacao"]
        registro_id = p["registro_id"]
        dados = json.loads(p["dados_json"]) if p["dados_json"] else {}
        tentativas = p["tentativas"] + 1

        sucesso_item = False
        erro_msg = None

        try:
            if operacao == "INSERT":
                supabase.table(tabela).insert(dados).execute()
                sucesso_item = True
            elif operacao == "UPDATE":
                supabase.table(tabela).update(dados).eq("id", registro_id).execute()
                sucesso_item = True
            elif operacao == "DELETE":
                supabase.table(tabela).delete().eq("id", registro_id).execute()
                sucesso_item = True
        except Exception as e:
            erro_msg = str(e)
            logger.error(f"Erro ao processar sync_log ID {log_id}: {erro_msg}")
            
            # Se o erro for de conflito (ex: PK já existe), podemos considerar sucesso no INSERT
            # ou implementar lógica de merge. Por simplicidade, vamos marcar como falha.
            if "duplicate key" in erro_msg.lower():
                # Já existe, vamos apenas atualizar em vez de falhar
                try:
                    if operacao == "INSERT":
                        supabase.table(tabela).update(dados).eq("id", registro_id).execute()
                        sucesso_item = True
                except:
                    pass

        # Atualiza o status no banco
        conn_w = get_connection()
        cur_w = conn_w.cursor()
        if sucesso_item:
            cur_w.execute("UPDATE sync_log SET sincronizado = 1, erro = NULL WHERE id = ?", (log_id,))
        else:
            cur_w.execute(
                "UPDATE sync_log SET tentativas = ?, erro = ? WHERE id = ?", 
                (tentativas, erro_msg, log_id)
            )
            sucesso_total = False
        conn_w.commit()
        conn_w.close()

    return sucesso_total


def _queue_worker():
    """Thread em background para processar a fila periodicamente."""
    while True:
        try:
            process_queue()
        except Exception as e:
            pass
        time.sleep(30)  # Tenta a cada 30 segundos


def start_offline_queue_worker():
    """Inicia a thread de processamento da fila."""
    thread = threading.Thread(target=_queue_worker, daemon=True, name="OfflineQueueWorker")
    thread.start()
    logger.info("Worker da fila offline iniciado.")
