"""
routers/recs.py — Rotas para histórico de RECs.
Consolida os antigos /api/recs e /api/rec.
"""
import json
from datetime import datetime
from fastapi import APIRouter, HTTPException
from fastapi.responses import JSONResponse
from database import get_connection, get_row_connection
from models import RecModel, RecSaveRequest

router = APIRouter(tags=["recs"])


@router.get("/api/recs")
def get_recs():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT numero_obra, data_criacao FROM historico_rec ORDER BY data_criacao DESC")
    rows = cursor.fetchall()
    conn.close()
    return [{"numero_obra": r[0], "data_criacao": r[1]} for r in rows]


@router.get("/api/recs/{numero_obra}")
def get_rec(numero_obra: str):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT dados_json FROM historico_rec WHERE numero_obra = ?", (numero_obra,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {"dados_json": row[0]}
    raise HTTPException(status_code=404, detail="REC não encontrado")


@router.post("/api/recs")
def save_rec(rec: RecModel):
    conn = get_connection()
    cursor = conn.cursor()
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("INSERT OR REPLACE INTO historico_rec (numero_obra, dados_json, data_criacao) VALUES (?, ?, ?)",
                   (rec.numero_obra, rec.dados_json, agora))
    conn.commit()
    conn.close()
    return {"status": "success"}


@router.delete("/api/recs/{numero_obra}")
def delete_rec(numero_obra: str):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM historico_rec WHERE numero_obra = ?", (numero_obra,))
    conn.commit()
    conn.close()
    return {"status": "success"}


# ── Rotas com prefixo /api/rec (Alternativas do Frontend) ───────────────

@router.post("/api/rec/salvar")
def salvar_rec_alt(req: RecSaveRequest):
    try:
        conn = get_connection()
        cursor = conn.cursor()
        data_agora = datetime.now().isoformat()
        dados_str = json.dumps(req.dados)

        cursor.execute("INSERT OR REPLACE INTO historico_rec (numero_obra, dados_json, data_criacao) VALUES (?, ?, ?)",
                       (req.numero_obra.strip(), dados_str, data_agora))
        conn.commit()
        conn.close()
        return {"status": "ok"}
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})


@router.get("/api/rec/{numero_obra}")
def recuperar_rec_alt(numero_obra: str):
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
