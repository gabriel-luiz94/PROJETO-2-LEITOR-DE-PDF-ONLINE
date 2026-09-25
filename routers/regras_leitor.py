"""
routers/regras_leitor.py — Rotas para as regras do leitor (TASK-006): Processamento e
Classificação, duas tabelas por projeto que alimentam o motor de regras
(static/regras_leitor_engine.js). Mesmo padrão de sincronização de routers/regras.py
(get_regras_conversao/save_regras_conversao, TASK-003): nuvem primeiro, fallback local, erro de
sincronização visível.
"""
import json
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from database import get_connection
from services.supabase_client import get_supabase
from config import logger

router = APIRouter(prefix="/api/regras-leitor", tags=["regras-leitor"])


class RegrasLeitorPayload(BaseModel):
    projeto_codigo: str
    regras: list


def _get_regras(tabela: str, projeto_codigo: str):
    chave_local = f"regras_leitor_{tabela}"

    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table(chave_local).select("regras_json").eq("projeto_codigo", projeto_codigo).execute()
            if res.data:
                return json.loads(res.data[0]["regras_json"])
        except Exception as e:
            logger.warning(f"Falha ao buscar regras_leitor_{tabela} no Supabase: {e}")

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(f"SELECT regras_json FROM {chave_local} WHERE projeto_codigo = ?", (projeto_codigo,))
    row = cursor.fetchone()
    conn.close()
    if row:
        try:
            return json.loads(row[0])
        except Exception:
            return []
    return []


def _save_regras(tabela: str, payload: RegrasLeitorPayload):
    chave_local = f"regras_leitor_{tabela}"
    valor = json.dumps(payload.regras, ensure_ascii=False)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        f"INSERT OR REPLACE INTO {chave_local} (projeto_codigo, regras_json) VALUES (?, ?)",
        (payload.projeto_codigo, valor)
    )
    conn.commit()
    conn.close()

    supabase = get_supabase()
    if supabase:
        try:
            supabase.table(chave_local).upsert({
                "projeto_codigo": payload.projeto_codigo,
                "regras_json": valor
            }).execute()
        except Exception as e:
            logger.warning(f"Erro ao salvar regras_leitor_{tabela} no Supabase: {e}")
            raise HTTPException(status_code=500, detail=f"Regras salvas localmente, mas falharam ao sincronizar com a nuvem (Supabase): {e}")

    return {"status": "success"}


@router.get("/processamento")
def get_processamento(projeto_codigo: str = "DEFAULT"):
    return {"regras": _get_regras("processamento", projeto_codigo)}


@router.post("/processamento")
def save_processamento(payload: RegrasLeitorPayload):
    return _save_regras("processamento", payload)


@router.get("/classificacao")
def get_classificacao(projeto_codigo: str = "DEFAULT"):
    return {"regras": _get_regras("classificacao", projeto_codigo)}


@router.post("/classificacao")
def save_classificacao(payload: RegrasLeitorPayload):
    return _save_regras("classificacao", payload)
