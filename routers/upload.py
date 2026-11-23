"""
routers/upload.py — Rotas para upload e extração (PDF/DXF/DWG), trigger de arquivo.

DWG (TASK-043): convertido para DXF via CloudConvert (services/cloudconvert_service.py) antes de
seguir pelo mesmo caminho do .dxf nativo — único formato aqui que depende de internet/chave.
"""
import os
import tempfile
import re
from typing import List
from fastapi import APIRouter, HTTPException, Request, UploadFile, File
from starlette.concurrency import run_in_threadpool
from database import get_connection
from services.pdf_service import extract_pdf_content
from services import importar_sap_pdf
from services.cloudconvert_service import ErroConversaoDwg, converter_dwg_para_dxf
from websocket_manager import manager
import json
from config import IS_FROZEN, logger

router = APIRouter(tags=["upload"])


def _chave_cloudconvert_padrao() -> str:
    return (os.getenv("CLOUDCONVERT_API_KEY") or "").strip()


def _chave_cloudconvert_salva() -> str | None:
    conn = get_connection()
    try:
        row = conn.execute("SELECT valor FROM configuracoes WHERE chave = 'cloudconvert_api_key'").fetchone()
    finally:
        conn.close()
    return row[0] if row else None


def resolver_chave_cloudconvert(header_key: str | None):
    """(chave, origem). Precedência: usuário (header) > salva > padrão do ambiente. Mesmo padrão
    de routers/ai_chat.py:resolver_credencial."""
    chave = (header_key or "").strip()
    if chave:
        return chave, "usuario"
    salva = _chave_cloudconvert_salva()
    if salva:
        return salva, "salva"
    padrao = _chave_cloudconvert_padrao()
    if padrao:
        return padrao, "padrao"
    return None, "nenhuma"


def origem_chave_cloudconvert_ativa() -> str:
    """Origem da chave que seria usada sem header (diagnóstico; nunca expõe o valor)."""
    return resolver_chave_cloudconvert(None)[1]


def _salvar_chave_cloudconvert_se_do_usuario(chave: str, origem: str):
    if origem != "usuario":
        return
    conn = get_connection()
    try:
        conn.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES ('cloudconvert_api_key', ?)", (chave,))
        conn.commit()
    finally:
        conn.close()


def _dxf_de_dwg(conteudo_dwg: bytes, nome_arquivo: str, header_key: str | None) -> bytes:
    chave, origem = resolver_chave_cloudconvert(header_key)
    dxf_bytes = converter_dwg_para_dxf(conteudo_dwg, nome_arquivo, chave)
    _salvar_chave_cloudconvert_se_do_usuario(chave, origem)  # só persiste depois de confirmar que a chave funciona
    return dxf_bytes


def _leitura_do_projeto(projeto):
    """TASK-175: configuração de leitura do desenho (layers/cores de rede) guardada para o projeto; None = o padrão de Pernambuco."""
    codigo = str(projeto or "").strip()
    if not codigo:
        return None
    try:
        from routers.plugin_cad import leitura_salva
        return leitura_salva(codigo)
    except Exception as e:  # noqa: BLE001
        logger.warning(f"Configuração de leitura do projeto indisponível: {e}")
        return None


def _extrair_dxf(caminho: str, leitura):
    import ezdxf
    from services.dxf_service import extract_dxf_content

    return extract_dxf_content(ezdxf.readfile(caminho), com_geometria=True, leitura=leitura)


def _extrair_de_bytes_dxf(conteudo: bytes, leitura=None):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
        tmp.write(conteudo)
        tmp_path = tmp.name
    try:
        return _extrair_dxf(tmp_path, leitura)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@router.post("/upload")
async def upload_file(request: Request, file: UploadFile = File(...), projeto: str = ""):
    leitura = _leitura_do_projeto(projeto)
    try:
        contents = await file.read()
        nome = file.filename.lower()
        if nome.endswith(".dwg"):
            try:
                contents = _dxf_de_dwg(contents, file.filename, request.headers.get("X-CloudConvert-Key"))
            except ErroConversaoDwg as e:
                return {"error": str(e), "data": []}
            linhas, geometria = _extrair_de_bytes_dxf(contents, leitura)
            return {"data": linhas, "geometria": geometria}
        if nome.endswith(".dxf"):
            linhas, geometria = _extrair_de_bytes_dxf(contents, leitura)
            return {"data": linhas, "geometria": geometria}

        import pymupdf
        from services.pdf_geometria import extrair_pdf_com_geometria

        with pymupdf.open(stream=contents, filetype="pdf") as doc:
            itens, geometria, escala = extrair_pdf_com_geometria(doc, leitura)             # TASK-236: rede pelos vetores + escala 1:500/1:1000
            return {"data": itens, "geometria": geometria, "escala": escala}
    except Exception as e:
        logger.error(f"Erro no upload: {e}")
        return {"error": str(e), "data": []}


@router.get("/extract-local")
async def extract_local(request: Request, path: str, projeto: str = ""):
    leitura = _leitura_do_projeto(projeto)
    # Proteção de segurança:
    # A extração de arquivo local por caminho só é permitida em ambiente frozen (.exe) local.
    if not IS_FROZEN:
        return {"error": "Acesso negado. Execução remota não permite ler arquivos locais via path."}

    if not os.path.exists(path):
        return {"error": "Arquivo não encontrado"}
    try:
        caminho_min = path.lower()
        if caminho_min.endswith(".dwg"):
            with open(path, "rb") as f:
                conteudo = f.read()
            try:
                dxf_bytes = _dxf_de_dwg(conteudo, os.path.basename(path), request.headers.get("X-CloudConvert-Key"))
            except ErroConversaoDwg as e:
                return {"error": str(e), "data": []}
            linhas, geometria = _extrair_de_bytes_dxf(dxf_bytes, leitura)
            return {"data": linhas, "geometria": geometria, "filename": os.path.basename(path)}
        if caminho_min.endswith(".dxf"):
            linhas, geometria = _extrair_dxf(path, leitura)
            return {"data": linhas, "geometria": geometria, "filename": os.path.basename(path)}
        import pymupdf
        from services.pdf_geometria import extrair_pdf_com_geometria

        with pymupdf.open(path) as doc:
            itens, geometria, escala = extrair_pdf_com_geometria(doc, leitura)             # TASK-236
            return {"data": itens, "geometria": geometria, "escala": escala, "filename": os.path.basename(path)}
    except Exception as e:
        logger.error(f"Erro em extract-local: {e}")
        return {"error": str(e), "data": []}


@router.post("/api/importar-rec-pdf")
async def importar_rec_pdf(orcamento: UploadFile = File(...), lista: UploadFile = File(...)):
    try:
        import pymupdf

        orcamento_contents = await orcamento.read()
        lista_contents = await lista.read()
        
        extracted_items = []
        
        # Processa ORÇAMENTO
        doc_orc = pymupdf.open(stream=orcamento_contents, filetype="pdf")
        current_mdo = "MAO-DE-OBRA"
        
        for page in doc_orc:
            blocks = page.get_text("dict", flags=11)["blocks"]
            texts = []
            for b in blocks:
                if "lines" not in b: continue
                for l in b["lines"]:
                    if "spans" not in l: continue
                    for s in l["spans"]:
                        text = s["text"].strip()
                        if text:
                            texts.append(text)
            
            for i, text in enumerate(texts):
                if text == "MAO-DE-OBRA":
                    current_mdo = "MAO-DE-OBRA"
                elif text == "MATERIAL":
                    current_mdo = "MATERIAL"
                
                if re.match(r'^\d{6}$', text):
                    codigo = text.lstrip('0') or '0'
                    desc = texts[i-1] if i > 0 else ""
                    
                    qtd = 1.0
                    for j in range(i-1, max(-1, i-5), -1):
                        if re.match(r'^-?\d{1,3}(\.\d{3})*(,\d+)?$', texts[j]) or re.match(r'^-?\d+(,\d+)?$', texts[j]) or re.match(r'^-?\d+(\.\d+)?$', texts[j]):
                            try:
                                clean_val = texts[j].replace('.', '') if ',' in texts[j] and texts[j].count('.') >= 1 else texts[j]
                                clean_val = clean_val.replace(',', '.')
                                qtd = float(clean_val)
                                break
                            except ValueError:
                                pass
                    
                    op = texts[i+1] if i + 1 < len(texts) else "I"
                    if op not in ['I', 'R']:
                        op = 'I'
                    
                    extracted_items.append({
                        "operacao": op,
                        "mdo": current_mdo,
                        "codigo": codigo,
                        "desc_codigo": desc,
                        "total": qtd
                    })
        doc_orc.close()

        # Processa LISTA
        doc_lista = pymupdf.open(stream=lista_contents, filetype="pdf")
        for page in doc_lista:
            blocks = page.get_text("dict", flags=11)["blocks"]
            texts = []
            for b in blocks:
                if "lines" not in b: continue
                for l in b["lines"]:
                    if "spans" not in l: continue
                    for s in l["spans"]:
                        text = s["text"].strip()
                        if text:
                            texts.append(text)
            
            for i, text in enumerate(texts):
                if re.match(r'^\d{6}$', text):
                    codigo = text.lstrip('0') or '0'
                    requisitar = None
                    devolver = None
                    desc = ""
                    
                    for j in range(i+1, min(i+10, len(texts))):
                        if re.match(r'^-?\d{1,3}(\.\d{3})*(,\d+)?$', texts[j]) or re.match(r'^-?\d+(,\d+)?$', texts[j]) or re.match(r'^-?\d+(\.\d+)?$', texts[j]):
                            try:
                                clean_val = texts[j].replace('.', '') if ',' in texts[j] and texts[j].count('.') >= 1 else texts[j]
                                clean_val = clean_val.replace(',', '.')
                                val = float(clean_val)
                                if requisitar is None:
                                    requisitar = val
                                elif devolver is None:
                                    devolver = val
                                    if j + 1 < len(texts):
                                        desc = texts[j+1]
                                    break
                            except ValueError:
                                pass
                    
                    if devolver and devolver > 0:
                        extracted_items.append({
                            "operacao": "R",
                            "mdo": "MATERIAL",
                            "codigo": codigo,
                            "desc_codigo": desc,
                            "total": devolver
                        })
        doc_lista.close()
        
        return {"resultado": extracted_items}
    except Exception as e:
        logger.error(f"Erro em importar-rec: {e}")
        return {"error": str(e)}


@router.get("/trigger-file")
async def trigger_file(path: str):
    await manager.broadcast(json.dumps({"type": "load_file", "path": path}))
    return {"status": "success"}


LIMITE_PDF_SAP = 20 * 1024 * 1024


@router.post("/api/importar-sap-pdf")
async def importar_sap_pdf_rota(arquivos: List[UploadFile] = File(...)):
    """TASK-157: lê o PDF SAP (ODI e/ou ODD, o gerado pelo programa ou o feito por fora no mesmo formato) e devolve as linhas no formato do
    Resultado do orçamento (ODI → I, ODD → R). Usado pelo Importar PDFs e pelo Comparar."""
    if not 1 <= len(arquivos) <= 2:
        raise HTTPException(status_code=400, detail="Escolha 1 ou 2 PDFs SAP (um ODI e/ou um ODD).")
    lidos, resumo = [], []
    for arq in arquivos:
        nome = os.path.basename(arq.filename or "arquivo.pdf")
        conteudo = await arq.read(LIMITE_PDF_SAP + 1)
        if len(conteudo) > LIMITE_PDF_SAP:
            raise HTTPException(status_code=413, detail=f"{nome}: arquivo grande demais (máx. {LIMITE_PDF_SAP // (1024 * 1024)} MB).")
        try:
            lido = await run_in_threadpool(importar_sap_pdf.ler_pdf, conteudo)
        except importar_sap_pdf.ErroSapPdf as e:
            raise HTTPException(status_code=400, detail=f"{nome}: {e}")
        lidos.append(lido)
        resumo += [{"nome": nome, "tipo": sec["tipo"], "nota": lido["nota"], "itens": len(sec["linhas"])} for sec in lido["secoes"]]       # um PDF pode ter ODI e ODD
    try:
        resultado = importar_sap_pdf.para_resultado(lidos)
    except importar_sap_pdf.ErroSapPdf as e:
        raise HTTPException(status_code=400, detail=str(e))
    return {"resultado": resultado, "arquivos": resumo}
