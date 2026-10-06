"""
routers/upload.py — Rotas para upload e extração (PDF/DXF/DWG), trigger de arquivo.

DWG (TASK-043): convertido para DXF via CloudConvert (services/cloudconvert_service.py) antes de
seguir pelo mesmo caminho do .dxf nativo — único formato aqui que depende de internet/chave.
"""
import os
import tempfile
import pymupdf
import ezdxf
import re
from fastapi import APIRouter, Request, UploadFile, File
from database import get_connection
from services.pdf_service import extract_pdf_content
from services.dxf_service import extract_dxf_content
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


def _extrair_de_bytes_dxf(conteudo: bytes) -> list:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
        tmp.write(conteudo)
        tmp_path = tmp.name
    try:
        return extract_dxf_content(ezdxf.readfile(tmp_path))
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@router.post("/upload")
async def upload_file(request: Request, file: UploadFile = File(...)):
    try:
        contents = await file.read()
        nome = file.filename.lower()
        if nome.endswith(".dwg"):
            try:
                contents = _dxf_de_dwg(contents, file.filename, request.headers.get("X-CloudConvert-Key"))
            except ErroConversaoDwg as e:
                return {"error": str(e), "data": []}
            return {"data": _extrair_de_bytes_dxf(contents)}
        if nome.endswith(".dxf"):
            return {"data": _extrair_de_bytes_dxf(contents)}

        # PDF
        doc = pymupdf.open(stream=contents, filetype="pdf")
        return {"data": extract_pdf_content(doc)}
    except Exception as e:
        logger.error(f"Erro no upload: {e}")
        return {"error": str(e), "data": []}


@router.get("/extract-local")
async def extract_local(request: Request, path: str):
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
            return {"data": _extrair_de_bytes_dxf(dxf_bytes), "filename": os.path.basename(path)}
        if caminho_min.endswith(".dxf"):
            doc = ezdxf.readfile(path)
            return {"data": extract_dxf_content(doc), "filename": os.path.basename(path)}
        # PDF
        doc = pymupdf.open(path)
        return {"data": extract_pdf_content(doc), "filename": os.path.basename(path)}
    except Exception as e:
        logger.error(f"Erro em extract-local: {e}")
        return {"error": str(e), "data": []}


@router.post("/api/importar-rec-pdf")
async def importar_rec_pdf(orcamento: UploadFile = File(...), lista: UploadFile = File(...)):
    try:
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
