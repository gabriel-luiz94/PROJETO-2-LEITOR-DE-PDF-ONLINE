"""
routers/obras.py — Rotas para gerenciamento de obras.
"""
from fastapi import APIRouter
from database import get_connection
from models import ObraModel

router = APIRouter(prefix="/api/obras", tags=["obras"])


@router.get("")
def get_obras():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nome, data, dados_json FROM obras ORDER BY data DESC")
    rows = cursor.fetchall()
    conn.close()
    return [{"id": r[0], "nome": r[1], "data": r[2], "dados_json": r[3]} for r in rows]


@router.post("")
def save_obra(obra: ObraModel):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO obras (id, nome, data, dados_json) VALUES (?, ?, ?, ?)",
                   (obra.id, obra.nome, obra.data, obra.dados_json))
    conn.commit()
    conn.close()
    return {"status": "success"}


@router.delete("/{obra_id}")
def delete_obra(obra_id: str):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM obras WHERE id = ?", (obra_id,))
    conn.commit()
    conn.close()
    return {"status": "success"}
