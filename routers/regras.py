"""
routers/regras.py — Rotas para gerenciamento de regras de aprendizado da IA.
"""
from fastapi import APIRouter
from database import get_connection
from models import RegraModel

router = APIRouter(prefix="/api/regras", tags=["regras"])


@router.get("")
def get_regras():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, conteudo FROM regras ORDER BY id ASC")
    rows = cursor.fetchall()
    conn.close()
    return [{"id": r[0], "conteudo": r[1]} for r in rows]


@router.post("")
def save_regra(regra: RegraModel):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO regras (conteudo, embedding) VALUES (?, NULL)", (regra.conteudo,))
    conn.commit()
    conn.close()
    return {"status": "success"}
