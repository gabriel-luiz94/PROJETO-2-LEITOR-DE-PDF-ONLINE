"""
routers/regras.py — Rotas para gerenciamento de regras de aprendizado da IA
               e regras de conversão de orçamento.
"""
import json
from fastapi import APIRouter
from pydantic import BaseModel
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


# ── Regras de Conversão de Orçamento ──────────────────────────────────────────

class RegrasConversaoModel(BaseModel):
    projeto_codigo: str
    regras: list


@router.get("/conversao")
def get_regras_conversao(projeto_codigo: str = "DEFAULT"):
    chave = f"regras_conversao_{projeto_codigo}"
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT valor FROM configuracoes WHERE chave = ?", (chave,))
    row = cursor.fetchone()
    conn.close()
    if row:
        try:
            return {"regras": json.loads(row[0])}
        except Exception:
            return {"regras": []}
    return {"regras": []}


@router.post("/conversao")
def save_regras_conversao(payload: RegrasConversaoModel):
    chave = f"regras_conversao_{payload.projeto_codigo}"
    valor = json.dumps(payload.regras, ensure_ascii=False)
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES (?, ?)",
        (chave, valor)
    )
    conn.commit()
    conn.close()
    return {"status": "success"}
