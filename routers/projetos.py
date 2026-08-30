"""
routers/projetos.py — Rotas para gerenciamento de projetos base.
"""
from fastapi import APIRouter
from fastapi.responses import JSONResponse
from fastapi import Request
from database import get_row_connection, get_connection
from models import ProjetoRequest
from middleware.auth_middleware import get_current_user_from_state

router = APIRouter(prefix="/api/projetos", tags=["projetos"])


from services.supabase_client import get_supabase

@router.get("")
def get_projetos(request: Request):
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("projetos").select("*").order("nome").execute()
            if res.data and len(res.data) > 0:
                conn = get_connection()
                cur = conn.cursor()
                for p in res.data:
                    cur.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES (?, ?)", (p["nome"], p["codigo"]))
                conn.commit()
                conn.close()
                return {"projetos": res.data}
        except Exception as e:
            pass

    conn = get_row_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM projetos ORDER BY nome")
    rows = cursor.fetchall()
    conn.close()
    return {"projetos": [dict(r) for r in rows]}


from fastapi import HTTPException

@router.post("")
def add_projeto(req: ProjetoRequest, request: Request):
    user = get_current_user_from_state(request)
    if user.get("role") != "admin":
        raise HTTPException(status_code=403, detail="Apenas administradores podem cadastrar novos projetos.")

    nome = req.nome.strip().upper()
    codigo = req.codigo.strip()

    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("projetos").upsert({"nome": nome, "codigo": codigo}).execute()
        except Exception:
            pass

    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES (?, ?)", (nome, codigo))
        conn.commit()
        conn.close()
        return {"status": "ok"}
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})
