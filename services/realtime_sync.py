"""
services/realtime_sync.py — Listener do Supabase Realtime.

Fica escutando mudanças na nuvem e atualiza o banco local automaticamente.
Emite eventos via WebSocket para atualizar a UI.
"""
import threading
import json
from config import logger, APP_MODE
from services.supabase_client import get_supabase
from database import get_connection
from websocket_manager import manager


def _realtime_worker():
    """Worker que se conecta ao Supabase Realtime e escuta eventos."""
    supabase = get_supabase()
    if not supabase:
        logger.warning("Realtime sync desativado: Supabase não configurado.")
        return

    logger.info("Conectando ao Supabase Realtime...")

    try:
        def on_change(response):
            logger.info(f"Mudança detectada na nuvem: {response}")
            payload = response.get("payload", {})
            event_type = payload.get("type")
            record = payload.get("record", {})
            old_record = payload.get("old_record", {})

            conn = get_connection()
            cur = conn.cursor()

            try:
                if event_type == "INSERT" or event_type == "UPDATE":
                    cur.execute('''
                        INSERT OR REPLACE INTO tabela_orcamento_master 
                        (id, ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        record.get("id"),
                        record.get("ativo", ""),
                        record.get("desc_ativo", ""),
                        record.get("componente", ""),
                        record.get("projeto", ""),
                        record.get("mdo", ""),
                        record.get("codigo", ""),
                        record.get("desc_codigo", ""),
                        record.get("fator_i", 0.0),
                        record.get("fator_r", 0.0),
                        record.get("filtro", ""),
                        record.get("updated_at")
                    ))
                elif event_type == "DELETE":
                    cur.execute("DELETE FROM tabela_orcamento_master WHERE id = ?", (old_record.get("id"),))

                conn.commit()
                logger.info(f"Banco local atualizado ({event_type}).")

                # Notifica o frontend via WebSocket
                import asyncio
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                loop.run_until_complete(manager.broadcast(json.dumps({
                    "type": "sync_master",
                    "action": event_type,
                    "record": record
                })))
                loop.close()

            except Exception as e:
                logger.error(f"Erro ao processar evento realtime: {e}")
            finally:
                conn.close()

        # Inscreve no canal da tabela master
        supabase.table("tabela_orcamento_master").on("*", on_change).subscribe()
        
        # Mantém a thread viva
        import time
        while True:
            time.sleep(10)

    except Exception as e:
        logger.error(f"Erro fatal no Supabase Realtime: {e}")


def start_realtime_sync():
    """Inicia o listener realtime em uma thread separada (Apenas modo desktop)."""
    if APP_MODE == "server":
        logger.info("Modo server: Supabase Realtime listener desativado (usando DB direto).")
        return
        
    thread = threading.Thread(target=_realtime_worker, daemon=True, name="RealtimeSyncThread")
    thread.start()
    logger.info("Thread de sync realtime iniciada.")
