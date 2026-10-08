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
from services.sync_service import (
    get_merged_orcamento, sync_tabela_master, filtrar_linhas_por_categoria,
    validar_linhas_do_projeto,
)
from middleware.auth_middleware import get_current_user_from_state

router = APIRouter(prefix="/api/orcamento", tags=["orcamento"])


from fastapi import APIRouter, UploadFile, File, Request, HTTPException

@router.post("/upload")
async def upload_orcamento(request: Request, file: UploadFile = File(...)):
    user = get_current_user_from_state(request)
    if user.get("role") != "admin":
        raise HTTPException(status_code=403, detail="Apenas administradores podem atualizar a base de orçamento.")
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
def get_orcamento_dados(request: Request, projeto: str = None, genericas: bool = False,
                         nao_reconhecido: bool = False):
    """Sem nenhum dos parâmetros abaixo, devolve a base inteira (comportamento histórico, usado
    pelo cálculo do orçamento — RN-06 já filtra por projeto internamente, com a base completa).

    TASK-052: `projeto` (igualdade exata), `genericas` (projeto vazio) ou `nao_reconhecido`
    (projeto preenchido mas não cadastrado em `projetos`) recortam a tela `/orcamento` por
    categoria — no máximo um por chamada.
    """
    user = getattr(request.state, "user", None)
    user_id = user["user_id"] if user else None
    sync_tabela_master()
    merged_rows = get_merged_orcamento(user_id)

    if projeto or genericas or nao_reconhecido:
        projetos_validos = None
        if nao_reconhecido:
            conn = get_row_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT nome FROM projetos")
            projetos_validos = [r["nome"] for r in cursor.fetchall()]
            conn.close()
        merged_rows = filtrar_linhas_por_categoria(
            merged_rows, projeto=projeto, genericas=genericas, nao_reconhecido=nao_reconhecido,
            projetos_validos=projetos_validos,
        )

    return {"dados": merged_rows}


@router.post("/salvar")
def salvar_orcamento(req: SalvarOrcamentoRequest, request: Request):
    user = get_current_user_from_state(request)
    if user.get("role") != "admin":
        raise HTTPException(status_code=403, detail="Apenas administradores podem atualizar a base de orçamento.")
    user_id = user["user_id"]
    dados = req.dados
    if req.projeto:
        # TASK-052: sem isso, salvar a visão filtrada de um projeto apagaria TODOS os projetos
        # da cópia pessoal do usuário (o DELETE abaixo não tinha filtro de projeto, só de dono).
        # Fora do `try` de baixo de propósito: o `except Exception` genérico ali mascararia o
        # HTTPException como um erro 400 "{error: ...}" em vez do formato padrão de validação.
        try:
            dados = validar_linhas_do_projeto(dados, req.projeto)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))

    try:
        conn = get_connection()
        cursor = conn.cursor()
        if req.projeto:
            projeto_norm = req.projeto.strip().upper()
            if user_id:
                cursor.execute(
                    "DELETE FROM tabela_orcamento WHERE user_id = ? AND UPPER(TRIM(projeto)) = ?",
                    (user_id, projeto_norm))
            else:
                cursor.execute(
                    "DELETE FROM tabela_orcamento WHERE user_id IS NULL AND UPPER(TRIM(projeto)) = ?",
                    (projeto_norm,))
        elif user_id:
            cursor.execute("DELETE FROM tabela_orcamento WHERE user_id = ?", (user_id,))
        else:
            cursor.execute("DELETE FROM tabela_orcamento WHERE user_id IS NULL")
        for row in dados:
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
def search_orcamento(q: str = "", col: str = None):
    if not q.strip():
        return {"resultados": []}

    termo = f"%{q.strip().upper()}%"
    conn = get_row_connection()
    cursor = conn.cursor()
    
    if col == "ativo":
        query = "SELECT * FROM tabela_orcamento WHERE upper(ativo) LIKE ? LIMIT 50"
        cursor.execute(query, (termo,))
    else:
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
    user = getattr(request.state, "user", None)
    user_id = user["user_id"] if user else None
    merged_rows = get_merged_orcamento(user_id)
    return processar_calculo(req.cabos, req.outros, req.projeto, merged_rows)
