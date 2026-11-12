"""
routers/orcamento.py — Rotas para gerenciamento e cálculo do orçamento.
"""
import io
import csv
import sqlite3
from fastapi import APIRouter, UploadFile, File, Request
from fastapi.responses import JSONResponse, Response
from database import get_connection, get_row_connection
from models import SalvarOrcamentoRequest, OrcamentoRequest, DetalhesRequest
from services.orcamento_calc import processar_calculo
from services.orcamento_detalhes import detalhes_por_codigo
from services.orcamento_lv import normalizar_lv
from services.orcamento_arred import normalizar_arred
from services.fator_formula import normalizar_fator
from services.ativo_parametros import preencher_componentes, validar_base, resumo_erros
from services.exportar_xlsx import ErroExportacao, gerar_xlsx
from services.exportar_sap_pdf import ErroPdf, gerar_pdf
from services.rascunho_ativo import avaliar as avaliar_rascunho
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
        try:
            text = contents.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = contents.decode("cp1252", errors="replace")          # TASK-137: CSV salvo como "CSV (separado por vírgulas)" do Excel vem em ANSI (Windows-1252), não UTF-8
        reader = csv.DictReader(io.StringIO(text), delimiter=";")

        # fallback to comma if no columns found
        if not reader.fieldnames or len(reader.fieldnames) < 2:
            reader = csv.DictReader(io.StringIO(text), delimiter=",")

        # Lê e valida TUDO antes de apagar a tabela atual (TASK-078: LV inválido recusa o arquivo sem tocar na base)
        linhas = []
        for n, row in enumerate(reader, start=2):
            # Map by ignoring case
            row_upper = {k.strip().upper(): v for k, v in row.items() if k}
            ativo = row_upper.get("ATIVO", "")
            codigo = row_upper.get("CODIGO", "")
            if not ativo and not codigo:
                continue

            try:
                fator_i, formula_i = normalizar_fator(row_upper.get("FATOR I"), "FATOR I")          # TASK-109: número ou fórmula (=NUM(fios)/8)
                fator_r, formula_r = normalizar_fator(row_upper.get("FATOR R"), "FATOR R")
            except ValueError as e:
                return JSONResponse(status_code=400, content={"error": f"Linha {n} do arquivo: {e} Nada foi alterado."})

            try:
                lv = normalizar_lv(row_upper.get("LV"))
            except ValueError as e:
                return JSONResponse(status_code=400, content={"error": f"Linha {n} do arquivo: {e} Nada foi alterado."})

            try:
                arred = normalizar_arred(row_upper.get("ARRED"))                 # TASK-107: coluna opcional; vazia = NORMAL
            except ValueError as e:
                return JSONResponse(status_code=400, content={"error": f"Linha {n} do arquivo: {e} Nada foi alterado."})

            linhas.append((ativo, row_upper.get("DESC ATIVO", ""), row_upper.get("COMPONENTE", ""), row_upper.get("PROJETO", ""),
                           row_upper.get("MDO", ""), row_upper.get("CODIGO", ""), row_upper.get("DESC CODIGO", ""),
                           fator_i, fator_r, row_upper.get("FILTRO", ""), lv, arred, formula_i, formula_r))

        # TASK-079: ativos com parâmetros (N1=N1(cabo=XLPE), N1(cabo=ALUM)…) — recusa cadastro inconsistente antes de apagar
        reg, n_comp = preencher_componentes([{"ativo": t[0], "componente": t[2], "projeto": t[3], "codigo": t[5], "lv": t[10],
                                       "formula_i": t[12], "formula_r": t[13]} for t in linhas])   # TASK-080; TASK-109: fórmulas do fator para a validação
        linhas = [t[:2] + (r["componente"],) + t[3:] for t, r in zip(linhas, reg)]
        val = validar_base(reg)
        if val["erros"]:
            return JSONResponse(status_code=400, content={"error": f"{resumo_erros(val['erros'])} Nada foi alterado."})

        conn = get_connection()
        cursor = conn.cursor()

        # Limpar tabela atual
        cursor.execute("DELETE FROM tabela_orcamento")
        for t in linhas:
            cursor.execute('''
                INSERT INTO tabela_orcamento (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, lv, arred, formula_i, formula_r)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', t)

        conn.commit()
        conn.close()
        resposta = {"status": "ok", "message": "Tabela carregada com sucesso."}
        if n_comp:
            resposta["componentes_gerados"] = n_comp
        if val["avisos"]:
            resposta["avisos"] = val["avisos"]
        return resposta
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

    # TASK-078: LV inválido recusa o salvamento inteiro (fora do `try` genérico, como a validação de projeto acima)
    for i, row in enumerate(dados, start=1):
        try:
            row["lv"] = normalizar_lv(row.get("lv"))
            row["arred"] = normalizar_arred(row.get("arred"))        # TASK-107
            row["fator_i"], row["formula_i"] = normalizar_fator(row.get("fator_i"), "FATOR I")      # TASK-109: número ou fórmula
            row["fator_r"], row["formula_r"] = normalizar_fator(row.get("fator_r"), "FATOR R")
        except ValueError as e:
            raise HTTPException(status_code=400, detail=f"Linha {i}: {e}")

    dados, n_comp = preencher_componentes(dados)       # TASK-080: COMPONENTE em branco dos grupos
    val = validar_base(dados)                         # TASK-079: ativos com parâmetros
    if val["erros"]:
        raise HTTPException(status_code=400, detail=resumo_erros(val["erros"]))

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

            fator_i, fator_r = row["fator_i"], row["fator_r"]                 # já normalizados acima (TASK-109)

            cursor.execute('''
                INSERT INTO tabela_orcamento (user_id, ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, lv, arred, formula_i, formula_r)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                row.get("filtro", "").strip(),
                row["lv"],
                row["arred"],
                row["formula_i"],
                row["formula_r"]
            ))

        conn.commit()
        conn.close()
        resposta = {"status": "ok", "message": "Tabela atualizada com sucesso."}
        if n_comp:
            resposta["componentes_gerados"] = n_comp
        if val["avisos"]:
            resposta["avisos"] = val["avisos"]
        return resposta
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
    """TASK-129: descrição, filtro e MDO de cada código a partir da Base de Orçamento (a mesma do cálculo: master/nuvem). `projeto` (opcional) prefere as linhas dele."""
    if not req.codigos:
        return {}
    return detalhes_por_codigo(get_merged_orcamento(None), req.codigos, req.projeto)


@router.post("/calcular")
def calcular_orcamento(req: OrcamentoRequest, request: Request):
    user = getattr(request.state, "user", None)
    user_id = user["user_id"] if user else None
    merged_rows = get_merged_orcamento(user_id)
    return processar_calculo(req.cabos, req.outros, req.projeto, merged_rows)


@router.post("/exportar-xlsx")
async def exportar_xlsx(request: Request):
    """TASK-070: planilhas ODI e ODD do Resultado do Orçamento (visão SAP). Corpo: {odi: [...], odd: [...]}, cada linha {mdo, codigo, desc, qtd}."""
    try:
        corpo = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail="Corpo inválido.")
    if not isinstance(corpo, dict):
        raise HTTPException(status_code=400, detail="Corpo inválido.")
    try:
        conteudo = gerar_xlsx(corpo.get("odi", []), corpo.get("odd", []))
    except ErroExportacao as e:
        raise HTTPException(status_code=400, detail=str(e))
    return Response(content=conteudo, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": 'attachment; filename="SAP.xlsx"'})


@router.post("/exportar-sap-pdf")
async def exportar_sap_pdf(request: Request):
    """TASK-140: relatório ODI ou ODD em PDF (formato da nota do SAP). Corpo: {tipo: "ODI"|"ODD", cabecalho: {...}, linhas: [{mdo, codigo, desc, qtd}]}."""
    try:
        corpo = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail="Corpo inválido.")
    if not isinstance(corpo, dict):
        raise HTTPException(status_code=400, detail="Corpo inválido.")
    try:
        conteudo = gerar_pdf(corpo.get("cabecalho") or {}, corpo.get("tipo", ""), corpo.get("linhas", []))
    except ErroPdf as e:
        raise HTTPException(status_code=400, detail=str(e))
    return Response(content=conteudo, media_type="application/pdf", headers={"Content-Disposition": 'attachment; filename="SAP.pdf"'})


@router.post("/rascunho")
async def avaliar_rascunho_ativo(request: Request):
    """TASK-147: avalia um rascunho de ativo (cadastro rápido) SEM gravar. Corpo: {linhas: [...], projeto, simular?: [{ativo}]}.
    Devolve os erros/avisos NOVOS (os que a base já tinha não entram), as linhas com COMPONENTE preenchido e a simulação pedida."""
    user = get_current_user_from_state(request)
    try:
        corpo = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail="Corpo inválido.")
    if not isinstance(corpo, dict):
        raise HTTPException(status_code=400, detail="Corpo inválido.")
    base = get_merged_orcamento(user.get("user_id"))
    return avaliar_rascunho(base, corpo.get("linhas") or [], corpo.get("projeto") or "", corpo.get("simular"))
