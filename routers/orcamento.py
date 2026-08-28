"""
routers/orcamento.py — Rotas para gerenciamento e cálculo do orçamento.
"""
import io
import csv
import sqlite3
from fastapi import APIRouter, UploadFile, File, Request
from fastapi.responses import JSONResponse
from database import get_connection, get_row_connection
from models import SalvarOrcamentoRequest, OrcamentoRequest, DetalhesRequest
from services.orcamento_calc import processar_calculo
from services.sync_service import get_merged_orcamento

router = APIRouter(prefix="/api/orcamento", tags=["orcamento"])


@router.post("/upload")
async def upload_orcamento(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        # Decode and parse CSV
        text = contents.decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(text), delimiter=";")

        # fallback to comma if no columns found
        if not reader.fieldnames or len(reader.fieldnames) < 2:
            reader = csv.DictReader(io.StringIO(text), delimiter=",")

        conn = get_connection()
        cursor = conn.cursor()

        # Limpar tabela atual
        cursor.execute("DELETE FROM tabela_orcamento")

        for row in reader:
            # Map by ignoring case
            row_upper = {k.strip().upper(): v for k, v in row.items() if k}
            ativo = row_upper.get("ATIVO", "")
            codigo = row_upper.get("CODIGO", "")
            if not ativo and not codigo:
                continue

            try:
                fator_i = float(row_upper.get("FATOR I", "0").replace(",", "."))
            except ValueError:
                fator_i = 0.0

            try:
                fator_r = float(row_upper.get("FATOR R", "0").replace(",", "."))
            except ValueError:
                fator_r = 0.0

            cursor.execute('''
                INSERT INTO tabela_orcamento (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                ativo,
                row_upper.get("DESC ATIVO", ""),
                row_upper.get("COMPONENTE", ""),
                row_upper.get("PROJETO", ""),
                row_upper.get("MDO", ""),
                row_upper.get("CODIGO", ""),
                row_upper.get("DESC CODIGO", ""),
                fator_i,
                fator_r,
                row_upper.get("FILTRO", "")
            ))

        conn.commit()
        conn.close()
        return {"status": "ok", "message": "Tabela carregada com sucesso."}
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})


@router.get("/dados")
def get_orcamento_dados(request: Request):
    # TODO: Extrair user_id do token/sessão local (request.state.user_id) na Fase 3
    user_id = None
    merged_rows = get_merged_orcamento(user_id)
    return {"dados": merged_rows}


@router.post("/salvar")
def salvar_orcamento(req: SalvarOrcamentoRequest, request: Request):
    try:
        # TODO: Extrair user_id do token/sessão na Fase 3
        user_id = None
        conn = get_connection()
        cursor = conn.cursor()
        if user_id:
            cursor.execute("DELETE FROM tabela_orcamento WHERE user_id = ?", (user_id,))
        else:
            cursor.execute("DELETE FROM tabela_orcamento WHERE user_id IS NULL")
        for row in req.dados:
            ativo = row.get("ativo", "").strip()
            codigo = row.get("codigo", "").strip()
            if not ativo and not codigo:
                continue

            try:
                fator_i = float(str(row.get("fator_i", "0")).replace(",", "."))
            except ValueError:
                fator_i = 0.0

            try:
                fator_r = float(str(row.get("fator_r", "0")).replace(",", "."))
            except ValueError:
                fator_r = 0.0

            cursor.execute('''
                INSERT INTO tabela_orcamento (user_id, ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                user_id,
                ativo,
                row.get("desc_ativo", "").strip(),
                row.get("componente", "").strip(),
                row.get("projeto", "").strip(),
                row.get("mdo", "").strip(),
                codigo,
                row.get("desc_codigo", "").strip(),
                fator_i,
                fator_r,
                row.get("filtro", "").strip()
            ))

        conn.commit()
        conn.close()
        return {"status": "ok", "message": "Tabela atualizada com sucesso."}
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})


@router.get("/search")
def search_orcamento(q: str = ""):
    if not q.strip():
        return {"resultados": []}

    termo = f"%{q.strip().upper()}%"
    conn = get_row_connection()
    cursor = conn.cursor()
    query = """
        SELECT * FROM tabela_orcamento 
        WHERE upper(ativo) LIKE ? 
           OR upper(desc_ativo) LIKE ? 
           OR upper(codigo) LIKE ? 
           OR upper(desc_codigo) LIKE ?
           OR upper(filtro) LIKE ?
        LIMIT 50
    """
    cursor.execute(query, (termo, termo, termo, termo, termo))
    rows = cursor.fetchall()
    conn.close()

    return {"resultados": [dict(r) for r in rows]}


@router.post("/detalhes")
def get_detalhes_codigos(req: DetalhesRequest):
    if not req.codigos:
        return {}

    conn = get_row_connection()
    cursor = conn.cursor()

    placeholders = ",".join(["?"] * len(req.codigos))
    cursor.execute(f"SELECT codigo, desc_codigo, mdo FROM tabela_orcamento WHERE codigo IN ({placeholders})", tuple(req.codigos))
    rows = cursor.fetchall()
    conn.close()

    resultado = {}
    for r in rows:
        resultado[r["codigo"]] = {
            "desc_codigo": r["desc_codigo"],
            "mdo": r["mdo"]
        }
    return resultado


@router.post("/calcular")
def calcular_orcamento(req: OrcamentoRequest, request: Request):
    # TODO: Extrair user_id do token na Fase 3
    user_id = None
    merged_rows = get_merged_orcamento(user_id)
    return processar_calculo(req.cabos, req.outros, req.projeto, merged_rows)
