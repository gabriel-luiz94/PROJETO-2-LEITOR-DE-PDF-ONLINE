"""
websocket_manager.py — Gerenciador de conexões WebSocket e endpoint /ws.
"""
import json
from fastapi import WebSocket, WebSocketDisconnect
from typing import Dict, List
from config import logger


class ConnectionManager:
    def __init__(self):
        # Mapeia user_id -> lista de WebSockets
        self.active_connections: Dict[str, List[WebSocket]] = {}
        # Conexões anônimas ou globais
        self.anonymous_connections: List[WebSocket] = []

    async def connect(self, websocket: WebSocket, user_id: str = None):
        await websocket.accept()
        if user_id:
            if user_id not in self.active_connections:
                self.active_connections[user_id] = []
            self.active_connections[user_id].append(websocket)
        else:
            self.anonymous_connections.append(websocket)

    def disconnect(self, websocket: WebSocket, user_id: str = None):
        if user_id and user_id in self.active_connections:
            if websocket in self.active_connections[user_id]:
                self.active_connections[user_id].remove(websocket)
            if not self.active_connections[user_id]:
                del self.active_connections[user_id]
        else:
            if websocket in self.anonymous_connections:
                self.anonymous_connections.remove(websocket)

    async def send_personal_message(self, message: str, user_id: str):
        if user_id in self.active_connections:
            for connection in self.active_connections[user_id]:
                try:
                    await connection.send_text(message)
                except Exception as e:
                    logger.warning(f"Erro ao enviar WebSocket para user {user_id}: {e}")

    async def broadcast(self, message: str):
        """Envia para todos os usuários conectados (logados ou não)."""
        all_connections = self.anonymous_connections.copy()
        for conns in self.active_connections.values():
            all_connections.extend(conns)

        for c in all_connections:
            try:
                await c.send_text(message)
            except Exception as e:
                pass


# Instância global
manager = ConnectionManager()
