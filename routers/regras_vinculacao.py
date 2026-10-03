"""
routers/regras_vinculacao.py — Regras de vinculação estrutura<->cabo (TASK-032): quantidade de
cabos exigida e compatibilidade entre eles, por tipo de estrutura, editável por projeto. Mesmo
padrão de sincronização de routers/regras_leitor.py (TASK-006/007): nuvem primeiro, fallback
local, erro de sincronização visível. Sem histórico/admin-gate — mesmo nível de abertura de
routers/regras.py (regras_conversao).
"""
import json
import re
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from database import get_connection
from services.supabase_client import get_supabase
from config import logger

router = APIRouter(prefix="/api/regras-vinculacao", tags=["regras-vinculacao"])

COMPATIBILIDADES_VALIDAS = {"LIVRE", "MESMO_TIPO_FASE_OPERACAO"}


class RegrasVinculacaoPayload(BaseModel):
    projeto_codigo: str
    regras: list


def _validar_regras(regras) -> list:
    """Espelha os campos que o algoritmo de vinculação (static/resumo.js) de fato interpreta."""
    if not isinstance(regras, list):
        return ["O payload de regras precisa ser uma lista."]
    erros = []
    for idx, r in enumerate(regras):
        prefixo = f"Regra #{idx + 1}"
        if not isinstance(r, dict):
            erros.append(f"{prefixo}: precisa ser um objeto.")
            continue
        tipo = r.get("tipo_estrutura")
        if not isinstance(tipo, str) or not re.fullmatch(r"\d", tipo or ""):
            erros.append(f"{prefixo}: 'tipo_estrutura' precisa ser um único dígito (string), recebido: {tipo!r}.")
        qtd = r.get("qtd_cabos")
        if not isinstance(qtd, int) or isinstance(qtd, bool) or qtd < 1:
            erros.append(f"{prefixo}: 'qtd_cabos' precisa ser um inteiro >= 1.")
        compat = r.get("compatibilidade")
        if compat not in COMPATIBILIDADES_VALIDAS:
            erros.append(f"{prefixo}: 'compatibilidade' precisa ser um de {sorted(COMPATIBILIDADES_VALIDAS)}.")
    return erros


def _get_regras(projeto_codigo: str):
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("regras_vinculacao").select("regras_json").eq("projeto_codigo", projeto_codigo).execute()
            if res.data:
                return json.loads(res.data[0]["regras_json"])
        except Exception as e:
            logger.warning(f"Falha ao buscar regras_vinculacao no Supabase: {e}")

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT regras_json FROM regras_vinculacao WHERE projeto_codigo = ?", (projeto_codigo,))
    row = cursor.fetchone()
    conn.close()
    if row:
        try:
            return json.loads(row[0])
        except Exception:
            return []
    return []


def _save_regras(payload: RegrasVinculacaoPayload):
    erros = _validar_regras(payload.regras)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})

    valor = json.dumps(payload.regras, ensure_ascii=False)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT OR REPLACE INTO regras_vinculacao (projeto_codigo, regras_json) VALUES (?, ?)",
        (payload.projeto_codigo, valor)
    )
    conn.commit()
    conn.close()

    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("regras_vinculacao").upsert({
                "projeto_codigo": payload.projeto_codigo,
                "regras_json": valor
            }).execute()
        except Exception as e:
            logger.warning(f"Erro ao salvar regras_vinculacao no Supabase: {e}")
            raise HTTPException(status_code=500, detail=f"Regras salvas localmente, mas falharam ao sincronizar com a nuvem (Supabase): {e}")

    return {"status": "success"}


@router.get("")
def get_regras_vinculacao(projeto_codigo: str = "DEFAULT"):
    return {"regras": _get_regras(projeto_codigo)}


@router.post("")
def save_regras_vinculacao(payload: RegrasVinculacaoPayload):
    return _save_regras(payload)
