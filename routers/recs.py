"""
routers/recs.py — Rotas para histórico de RECs.
Consolida os antigos /api/recs e /api/rec.
"""
import json
from datetime import datetime
from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import JSONResponse
from database import get_connection, get_row_connection
from models import RecModel, RecSaveRequest
from middleware.auth_middleware import get_current_user_from_state

router = APIRouter(tags=["recs"])


from services.supabase_client import get_supabase

@router.get("/api/recs")
def get_recs(request: Request):
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("historico_rec").select("numero_obra, data_criacao, user_id").order("data_criacao", desc=True).execute()
            if res.data is not None:
                return res.data
        except Exception:
            pass

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT numero_obra, data_criacao FROM historico_rec ORDER BY data_criacao DESC")
    rows = cursor.fetchall()
    conn.close()
    return [{"numero_obra": r[0], "data_criacao": r[1]} for r in rows]


@router.get("/api/recs/{numero_obra}")
def get_rec(numero_obra: str, request: Request):
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("historico_rec").select("dados_json").eq("numero_obra", numero_obra).execute()
            if res.data and len(res.data) > 0:
                return {"dados_json": res.data[0]["dados_json"]}
        except Exception:
            pass

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT dados_json FROM historico_rec WHERE numero_obra = ?", (numero_obra,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {"dados_json": row[0]}
    raise HTTPException(status_code=404, detail="REC não encontrado")


@router.post("/api/recs")
def save_rec(rec: RecModel, request: Request):
    user = get_current_user_from_state(request)
    user_id = str(user["user_id"])
    user_role = user.get("role", "operador")
    user_email = user.get("email", "")
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    num_obra_alvo = rec.numero_obra.strip()
    is_copy = False
    message = f"Obra Nº {num_obra_alvo} salva no banco com sucesso!"

    # 1. Verifica se já existe um REC com esse número
    conn = get_row_connection()
    cur = conn.cursor()
    cur.execute("SELECT user_id FROM historico_rec WHERE numero_obra = ?", (num_obra_alvo,))
    existing = cur.fetchone()
    conn.close()

    existing_owner_id = str(existing["user_id"]) if existing and existing["user_id"] else None

    supabase = get_supabase()
    if not existing_owner_id and supabase:
        try:
            res_check = supabase.table("historico_rec").select("user_id").eq("numero_obra", num_obra_alvo).execute()
            if res_check.data and len(res_check.data) > 0:
                existing_owner_id = str(res_check.data[0]["user_id"]) if res_check.data[0].get("user_id") else None
        except Exception:
            pass

    # 2. Se o REC já existe, pertence a outro usuário e o atual NÃO é admin:
    # Cria uma cópia pessoal para preservar o REC original
    if existing_owner_id and existing_owner_id != user_id and user_role != "admin":
        user_alias = user_email.split("@")[0] if user_email else "copia"
        num_obra_alvo = f"{num_obra_alvo} (Cópia - {user_alias})"
        is_copy = True
        message = f"O REC original foi preservado. Sua versão modificada foi salva como: '{num_obra_alvo}'."

    # 3. Salva no Supabase
    if supabase:
        try:
            supabase.table("historico_rec").upsert({
                "numero_obra": num_obra_alvo,
                "dados_json": rec.dados_json,
                "data_criacao": agora,
                "user_id": user_id
            }).execute()
        except Exception:
            pass

    # 4. Salva no SQLite local
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO historico_rec (numero_obra, dados_json, data_criacao, user_id) VALUES (?, ?, ?, ?)",
                   (num_obra_alvo, rec.dados_json, agora, user_id))
    conn.commit()
    conn.close()
    return {
        "status": "success",
        "numero_obra": num_obra_alvo,
        "is_copy": is_copy,
        "message": message
    }


@router.delete("/api/recs/{numero_obra}")
def delete_rec(numero_obra: str, request: Request):
    user = get_current_user_from_state(request)
    user_id = str(user["user_id"])
    user_role = user.get("role", "operador")

    conn = get_row_connection()
    cur = conn.cursor()
    cur.execute("SELECT user_id FROM historico_rec WHERE numero_obra = ?", (numero_obra,))
    existing = cur.fetchone()
    conn.close()

    existing_owner_id = str(existing["user_id"]) if existing and existing["user_id"] else None

    # Bloqueia se um operador tentar excluir REC de outro usuário
    if user_role != "admin" and existing_owner_id and existing_owner_id != user_id:
        raise HTTPException(status_code=403, detail="Você só pode excluir os RECs criados por você.")

    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("historico_rec").delete().eq("numero_obra", numero_obra).execute()
        except Exception:
            pass

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM historico_rec WHERE numero_obra = ?", (numero_obra,))
    conn.commit()
    conn.close()
    return {"status": "success", "msg": f"REC {numero_obra} excluído com sucesso."}


# ── Rotas com prefixo /api/rec (Alternativas do Frontend) ───────────────

@router.post("/api/rec/salvar")
def salvar_rec_alt(req: RecSaveRequest, request: Request):
    user = get_current_user_from_state(request)
    try:
        conn = get_connection()
        cursor = conn.cursor()
        data_agora = datetime.now().isoformat()
        dados_str = json.dumps(req.dados)

        cursor.execute("INSERT OR REPLACE INTO historico_rec (numero_obra, dados_json, data_criacao, user_id) VALUES (?, ?, ?, ?)",
                       (req.numero_obra.strip(), dados_str, data_agora, user["user_id"]))
        conn.commit()
        conn.close()
        return {"status": "ok"}
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})


@router.get("/api/rec/{numero_obra}")
def recuperar_rec_alt(numero_obra: str, request: Request):
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("historico_rec").select("dados_json").eq("numero_obra", numero_obra.strip()).execute()
            if res.data and len(res.data) > 0:
                return {"status": "ok", "dados": json.loads(res.data[0]["dados_json"])}
        except Exception:
            pass

    try:
        conn = get_row_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT dados_json FROM historico_rec WHERE numero_obra = ?", (numero_obra.strip(),))
        row = cursor.fetchone()
        conn.close()

        if row:
            return {"status": "ok", "dados": json.loads(row["dados_json"])}
        else:
            return JSONResponse(status_code=404, content={"error": "Obra não encontrada"})
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})
