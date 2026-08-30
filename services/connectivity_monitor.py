"""
services/connectivity_monitor.py — Monitor de conectividade com a nuvem.

Verifica periodicamente se o Supabase está acessível e emite eventos de
status via WebSocket para atualizar o indicador no frontend.
"""
import threading
import time
import json
from services.supabase_client import get_supabase
from websocket_manager import manager
from config import logger, APP_MODE


class ConnectivityStatus:
    is_online = False
    last_checked = 0


def _monitor_worker():
    """Worker que testa a conexão a cada 15 segundos."""
    while True:
        supabase = get_supabase()
        is_online = False
        
        if supabase:
            try:
                # Testa conectividade com uma query levíssima
                supabase.table("configuracoes").select("chave").limit(1).execute()
                is_online = True
            except Exception:
                pass
                
        # Se houve mudança de status, notifica
        if ConnectivityStatus.is_online != is_online:
            ConnectivityStatus.is_online = is_online
            logger.info(f"Status de conectividade alterado: {'ONLINE' if is_online else 'OFFLINE'}")
            
            import asyncio
            try:
                loop = asyncio.get_event_loop()
            except RuntimeError:
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                
            if loop.is_running():
                asyncio.create_task(manager.broadcast(json.dumps({
                    "type": "connectivity",
                    "status": "online" if is_online else "offline"
                })))
            else:
                loop.run_until_complete(manager.broadcast(json.dumps({
                    "type": "connectivity",
                    "status": "online" if is_online else "offline"
                })))
                loop.close()
                
        time.sleep(15)


def start_connectivity_monitor():
    """Inicia a thread de monitoramento (Apenas desktop)."""
    if APP_MODE == "server":
        return
        
    # Presumir online inicialmente para não travar a UI enquanto testa
    ConnectivityStatus.is_online = True
    
    thread = threading.Thread(target=_monitor_worker, daemon=True, name="ConnectivityMonitor")
    thread.start()
    logger.info("Monitor de conectividade iniciado.")
