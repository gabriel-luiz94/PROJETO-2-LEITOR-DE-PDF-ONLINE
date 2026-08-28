"""
routers/upload.py — Rotas para upload e extração (PDF/DXF), trigger de arquivo.
"""
import os
import tempfile
import pymupdf
import ezdxf
import re
from fastapi import APIRouter, UploadFile, File
from services.pdf_service import extract_pdf_content
from services.dxf_service import extract_dxf_content
from websocket_manager import manager
import json
from config import IS_FROZEN, logger

router = APIRouter(tags=["upload"])


@router.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        if file.filename.lower().endswith(".dxf"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
                tmp.write(contents)
                tmp_path = tmp.name
            try:
                doc = ezdxf.readfile(tmp_path)
                return {"data": extract_dxf_content(doc)}
            finally:
                if os.path.exists(tmp_path): 
                    os.remove(tmp_path)

        # PDF
        doc = pymupdf.open(stream=contents, filetype="pdf")
        return {"data": extract_pdf_content(doc)}
    except Exception as e:
        logger.error(f"Erro no upload: {e}")
        return {"error": str(e), "data": []}


@router.get("/extract-local")
async def extract_local(path: str):
    # Proteção de segurança:
    # A extração de arquivo local por caminho só é permitida em ambiente frozen (.exe) local.
    if not IS_FROZEN:
        return {"error": "Acesso negado. Execução remota não permite ler arquivos locais via path."}
        
    if not os.path.exists(path): 
        return {"error": "Arquivo não encontrado"}
    try:
        if path.lower().endswith(".dxf"):
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
        current_mdo = "MÃO-DE-OBRA"
        
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
                    current_mdo = "MÃO-DE-OBRA"
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
