"""
services/connectivity_monitor.py — Monitor de conectividade com a nuvem (TASK-040).

Verifica periodicamente se o Supabase está acessível e emite eventos de
status via WebSocket para atualizar o indicador no frontend. Ao detectar a
volta da conexão, força a recriação do cliente Supabase (`reset_supabase_client`)
em vez de só reportar o status — complementa o TTL proativo e o reset reativo
de `services/supabase_client.py` (TASK-039).
"""
import asyncio
import json
import threading
import time

from config import logger
from services.supabase_client import get_supabase, reset_supabase_client
from websocket_manager import manager

INTERVALO_S = 15


class ConnectivityStatus:
    is_online = False
    last_checked = 0


def _rede_ativa() -> bool:
    try:
        from services import rede_adaptativa
        return rede_adaptativa.ativa()
    except Exception:  # noqa: BLE001
        return False


def _mensagem(online: bool) -> str:
    """Evento do WebSocket. No desktop leva também o caminho de rede e a data da cópia local (o indicador da tela mostra "modo local — cópia de …")."""
    msg = {"type": "connectivity", "status": "online" if online else "offline"}
    if _rede_ativa():
        try:
            from services import rede_adaptativa, copia_local
            msg["caminho"] = rede_adaptativa.estado()["caminho"]
            msg["copia"] = copia_local.estado()["ultima"]
        except Exception:  # noqa: BLE001
            pass
    return json.dumps(msg)


def _testar_conexao() -> bool:
    if _rede_ativa():                                      # TASK-121 (desktop): quem testa a rede é a vigia de rede_adaptativa; aqui só se espelha o estado dela
        from services import rede_adaptativa
        return rede_adaptativa.estado()["online"]
    supabase = get_supabase()
    if not supabase:
        return False
    try:
        supabase.table("configuracoes").select("chave").limit(1).execute()
        return True
    except Exception:
        return False


def _monitor_worker(loop: asyncio.AbstractEventLoop):
    """Worker que testa a conexão a cada INTERVALO_S segundos, numa thread separada."""
    while True:
        is_online = _testar_conexao()

        if ConnectivityStatus.is_online != is_online:
            if is_online:
                # volta de uma queda: força um cliente novo em vez de só reportar "online" com um
                # cliente que pode ter ficado numa conexão ruim durante a janela offline.
                reset_supabase_client()
            ConnectivityStatus.is_online = is_online
            logger.info(f"Status de conectividade com o Supabase alterado: {'ONLINE' if is_online else 'OFFLINE'}")
            mensagem = _mensagem(is_online)
            try:
                asyncio.run_coroutine_threadsafe(manager.broadcast(mensagem), loop)
            except Exception as e:  # noqa: BLE001 — a vigia nunca morre por falha ao notificar
                logger.warning(f"Falha ao transmitir status de conectividade: {e}")

        ConnectivityStatus.last_checked = time.time()
        time.sleep(3 if _rede_ativa() else INTERVALO_S)       # espelhar o estado da vigia de rede não custa rede: pode ser frequente


def start_connectivity_monitor():
    """Inicia a thread de monitoramento. Precisa ser chamada de dentro de um handler async (ou de
    qualquer lugar com um loop asyncio já rodando na thread principal), para capturar o loop certo
    e poder agendar o broadcast do WebSocket a partir da thread do monitor com segurança."""
    loop = asyncio.get_event_loop()

    # Presumir online inicialmente para não travar a UI enquanto testa
    ConnectivityStatus.is_online = True

    thread = threading.Thread(target=_monitor_worker, args=(loop,), daemon=True, name="ConnectivityMonitor")
    thread.start()
    logger.info("Monitor de conectividade iniciado.")
