"""
routers/obras.py — Rotas para gerenciamento de obras.
"""
from fastapi import APIRouter, HTTPException, Request
from database import get_connection
from models import ObraModel
from middleware.auth_middleware import get_current_user_from_state

router = APIRouter(prefix="/api/obras", tags=["obras"])


from services.supabase_client import get_supabase, registrar_falha

@router.get("")
def get_obras(request: Request, projeto: str = None):
    user = get_current_user_from_state(request)
    user_id = user["user_id"]

    supabase = get_supabase()
    if supabase:
        try:
            query = supabase.table("obras").select("*").eq("user_id", user_id)
            if projeto:
                query = query.eq("projeto", projeto)
            res = query.order("data", desc=True).execute()
            if res.data is not None:
                return res.data
        except Exception as e:
            registrar_falha("obras.get_obras", e)

    conn = get_connection()
    cursor = conn.cursor()
    if projeto:
        cursor.execute(
            "SELECT id, nome, data, dados_json, projeto FROM obras WHERE (user_id = ? OR user_id IS NULL) AND projeto = ? ORDER BY data DESC",
            (user_id, projeto)
        )
    else:
        cursor.execute(
            "SELECT id, nome, data, dados_json, projeto FROM obras WHERE user_id = ? OR user_id IS NULL ORDER BY data DESC",
            (user_id,)
        )
    rows = cursor.fetchall()
    conn.close()
    return [{"id": r[0], "nome": r[1], "data": r[2], "dados_json": r[3], "projeto": r[4]} for r in rows]


def listar_leves(user_id, projeto):
    """Obras do usuário no projeto SEM `dados_json` (id, nome, data, projeto), mais recentes primeiro. Barato: serve ao índice da IA."""
    supabase = get_supabase()
    if supabase:
        try:
            res = (supabase.table("obras").select("id, nome, data, projeto").eq("user_id", user_id).eq("projeto", projeto)
                   .order("data", desc=True).execute())
            if res.data is not None:
                return res.data
        except Exception as e:
            registrar_falha("obras.listar_leves", e)
    conn = get_connection()
    rows = conn.execute(
        "SELECT id, nome, data, projeto FROM obras WHERE (user_id = ? OR user_id IS NULL) AND projeto = ? ORDER BY data DESC",
        (user_id, projeto)).fetchall()
    conn.close()
    return [{"id": r[0], "nome": r[1], "data": r[2], "projeto": r[3]} for r in rows]


def buscar_obra(user_id, obra_id):
    """Uma obra completa do usuário (com `dados_json`) ou None."""
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("obras").select("*").eq("id", obra_id).eq("user_id", user_id).execute()
            if res.data:
                return res.data[0]
        except Exception as e:
            registrar_falha("obras.buscar_obra", e)
    conn = get_connection()
    r = conn.execute("SELECT id, nome, data, dados_json, projeto FROM obras WHERE id = ? AND (user_id = ? OR user_id IS NULL)",
                     (obra_id, user_id)).fetchone()
    conn.close()
    return {"id": r[0], "nome": r[1], "data": r[2], "dados_json": r[3], "projeto": r[4]} if r else None


@router.get("/indice")
def indice_obras(request: Request, projeto: str):
    """Índice leve (sem dados_json) das obras do usuário no projeto — TASK-028."""
    return listar_leves(get_current_user_from_state(request)["user_id"], projeto)


@router.get("/{obra_id}")
def get_obra(obra_id: str, request: Request):
    """Uma obra completa do usuário (o frontend valida o id citado pela IA contra isto) — TASK-028."""
    obra = buscar_obra(get_current_user_from_state(request)["user_id"], obra_id)
    if obra is None:
        raise HTTPException(status_code=404, detail="Obra não encontrada.")
    return obra


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
                "user_id": user_id,
                "projeto": obra.projeto
            }).execute()
        except Exception as e:
            registrar_falha("obras.save_obra", e)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO obras (id, nome, data, dados_json, user_id, projeto) VALUES (?, ?, ?, ?, ?, ?)",
                   (obra.id, obra.nome, obra.data, obra.dados_json, user_id, obra.projeto))
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
        except Exception as e:
            registrar_falha("obras.delete_obra", e)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM obras WHERE id = ? AND user_id = ?", (obra_id, user_id))
    conn.commit()
    conn.close()
    return {"status": "success"}
