"""
websocket_manager.py — Gerenciador de conexões WebSocket e endpoint /ws.
"""
import json
from fastapi import WebSocket, WebSocketDisconnect
from config import logger


class ConnectionManager:
    def __init__(self):
        self.active_connections: list[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        if websocket in self.active_connections:
            self.active_connections.remove(websocket)

    async def broadcast(self, message: str):
        for c in self.active_connections:
            try:
                await c.send_text(message)
            except Exception as e:
                logger.warning(f"Erro ao enviar WebSocket: {e}")


# Instância global
manager = ConnectionManager()
