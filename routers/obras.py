"""
routers/obras.py — Rotas para gerenciamento de obras.
"""
from fastapi import APIRouter, Request
from database import get_connection
from models import ObraModel
from middleware.auth_middleware import get_current_user_from_state

router = APIRouter(prefix="/api/obras", tags=["obras"])


from services.supabase_client import get_supabase

@router.get("")
def get_obras(request: Request):
    user = get_current_user_from_state(request)
    user_id = user["user_id"]
    
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("obras").select("*").eq("user_id", user_id).order("data", desc=True).execute()
            if res.data is not None:
                return res.data
        except Exception:
            pass

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nome, data, dados_json FROM obras WHERE user_id = ? OR user_id IS NULL ORDER BY data DESC", (user_id,))
    rows = cursor.fetchall()
    conn.close()
    return [{"id": r[0], "nome": r[1], "data": r[2], "dados_json": r[3]} for r in rows]


@router.post("")
def save_obra(obra: ObraModel, request: Request):
    user = get_current_user_from_state(request)
    user_id = user["user_id"]

    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("obras").upsert({
                "id": obra.id,
                "nome": obra.nome,
                "data": obra.data,
                "dados_json": obra.dados_json,
                "user_id": user_id
            }).execute()
        except Exception:
            pass

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO obras (id, nome, data, dados_json, user_id) VALUES (?, ?, ?, ?, ?)",
                   (obra.id, obra.nome, obra.data, obra.dados_json, user_id))
    conn.commit()
    conn.close()
    return {"status": "success"}


@router.delete("/{obra_id}")
def delete_obra(obra_id: str, request: Request):
    user = get_current_user_from_state(request)
    user_id = user["user_id"]

    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("obras").delete().eq("id", obra_id).eq("user_id", user_id).execute()
        except Exception:
            pass

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM obras WHERE id = ? AND user_id = ?", (obra_id, user_id))
    conn.commit()
    conn.close()
    return {"status": "success"}
