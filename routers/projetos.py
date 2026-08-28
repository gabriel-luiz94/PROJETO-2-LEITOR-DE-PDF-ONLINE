"""
routers/projetos.py — Rotas para gerenciamento de projetos base.
"""
from fastapi import APIRouter
from fastapi.responses import JSONResponse
from database import get_row_connection, get_connection
from models import ProjetoRequest

router = APIRouter(prefix="/api/projetos", tags=["projetos"])


@router.get("")
def get_projetos():
    conn = get_row_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM projetos ORDER BY nome")
    rows = cursor.fetchall()
    conn.close()
    return {"projetos": [dict(r) for r in rows]}


@router.post("")
def add_projeto(req: ProjetoRequest):
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES (?, ?)",
                       (req.nome.strip().upper(), req.codigo.strip()))
        conn.commit()
        conn.close()
        return {"status": "ok"}
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})
