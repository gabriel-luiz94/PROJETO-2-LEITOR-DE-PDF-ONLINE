from fastapi import FastAPI, UploadFile, File, WebSocket, WebSocketDisconnect
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import pymupdf
import ezdxf
from ezdxf.colors import int2rgb, aci2rgb
from ezdxf.tools.text import plain_mtext
import tempfile
import os
import sys
import re
import json
import threading
import webbrowser

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =====================================================
# WEBSOCKET
# =====================================================

class ConnectionManager:
    def __init__(self):
        self.active_connections = []

    async def connect(self, websocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket):
        if websocket in self.active_connections:
            self.active_connections.remove(websocket)

    async def broadcast(self, message):
        for c in self.active_connections:
            try:
                await c.send_text(message)
            except:
                pass

manager = ConnectionManager()

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    try:
        while True:
            await websocket.receive_text()
    except WebSocketDisconnect:
        manager.disconnect(websocket)

@app.get("/trigger-file")
async def trigger_file(path: str):
    await manager.broadcast(json.dumps({"type": "load_file", "path": path}))
    return {"status": "success"}

# =====================================================
# PDF HELPERS
# =====================================================

def color_to_hex(c):
    if c is None: return "#000000"
    try:
        if isinstance(c, (list, tuple)):
            if len(c) >= 3:
                rgb = [max(0, min(255, int(x * 255))) for x in c[:3]]
                return "#%02x%02x%02x" % tuple(rgb)
        val = int(c)
        return "#%06x" % (val & 0xFFFFFF)
    except: return "#000000"

def flags_decomposer(flags):
    l = []
    if flags & 1: l.append("superscript")
    if flags & 2: l.append("italic")
    if flags & 4: l.append("serifed")
    else: l.append("sans")
    if flags & 8: l.append("monospaced")
    else: l.append("proportional")
    if flags & 16: l.append("bold")
    return ", ".join(l)

# =====================================================
# DXF HELPERS (CORRIGIDO)
# =====================================================

def _layer_name_to_hex(layer_name: str) -> str | None:
    """
    Extrai cor RGB do padrão de nomenclatura de layer usado neste DXF.
    Exemplo: 'TXT_FF0000' → '#FF0000', 'TXT_0000FF' → '#0000FF'.
    Retorna None se o layer não seguir o padrão.
    """
    m = re.match(r'.*_([0-9A-Fa-f]{6})$', layer_name)
    return f"#{m.group(1).upper()}" if m else None


def _aci_to_hex(aci: int) -> str:
    """
    Converte índice ACI para hex RGB.
    ACI 7 = branco (fundo escuro) → mapeado para preto no papel branco.
    ACI 0 = BYBLOCK → preto.
    """
    if aci in (0, 7):
        return "#000000"
    try:
        rgb = aci2rgb(aci)
        return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
    except:
        return "#000000"


# Remove códigos de formatação de parágrafo MTEXT: \pxql; \pxqc; \pxqr; etc.
# IMPORTANTE:
#   - SEM re.IGNORECASE: \p (minúsculo) = formatação de parágrafo
#                        \P (maiúsculo) = quebra de parágrafo — NÃO deve ser removido
#   - Ponto-e-vírgula OBRIGATÓRIO (sem ?): garante que \P seguido de texto
#     (ex: \PCAZ 9,5) não seja confundido com código de formatação.
_RE_PARA_FMT = re.compile(r'\\p[^;\\\\]*;')


def _get_entity_base_color(entity, doc) -> str:
    """
    Resolve a cor da entidade a partir de atributos de nivel de entidade/layer.
    NAO verifica codigos inline (\\C, \\c) -- esses sao resolvidos por
    paragrafo em _resolve_part_color para permitir heranca correta.
    Hierarquia:
      1. true_color da entidade
      2. ACI da entidade (nao-BYLAYER, nao-BYBLOCK)
      3. RGB no nome do layer (TXT_RRGGBB)
      4. true_color do layer
      5. ACI do layer
      6. Fallback preto
    """
    if entity.dxf.hasattr("true_color"):
        try:
            rgb = int2rgb(entity.dxf.true_color)
            return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        except:
            pass
    aci = entity.dxf.color
    if aci not in (0, 256):
        return _aci_to_hex(aci)
    try:
        layer_name = entity.dxf.layer
        h = _layer_name_to_hex(layer_name)
        if h:
            return h
        layer = doc.layers.get(layer_name)
        if layer.dxf.hasattr("true_color"):
            rgb = int2rgb(layer.dxf.true_color)
            return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        return _aci_to_hex(layer.color)
    except:
        pass
    return "#000000"


def _resolve_part_color(part_raw: str, inherited: str) -> str:
    """
    Resolve a cor de um parágrafo individual.
    Inline \\c<decimal> (true-color) e \\C<n> (ACI) dentro do part
    substituem a cor herdada. O ÚLTIMO código encontrado no parágrafo 
    é o que define a cor da linha e o que será herdado pelo próximo.
    """
    # Procura todos os códigos de cor no parágrafo
    # \c é true-color (RGB em decimal), \C é ACI
    matches = re.findall(r'\\([cC])(\d+)', part_raw)
    if not matches:
        return inherited
    
    # Pega o último código (é o que "vence" no final do parágrafo)
    code_type, value = matches[-1]
    
    if code_type == 'c':
        try:
            rgb = int2rgb(int(value))
            return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        except:
            pass
    else:
        aci = int(value)
        if aci == 0 or aci == 256:
            return inherited
        return _aci_to_hex(aci)
        
    return inherited


def get_dxf_color(entity, doc) -> str:
    """
    Resolve a cor real de uma entidade TEXT/MTEXT.
    Para MTEXT com multiplos paragrafos, use _get_entity_base_color +
    _resolve_part_color por paragrafo (feito em _process_entity).
    """
    raw = entity.text if entity.dxftype() == "MTEXT" else None
    if raw:
        m = re.search(r'\\c(\d+)', raw)
        if m:
            try:
                rgb = int2rgb(int(m.group(1)))
                return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
            except:
                pass
        m = re.search(r'\\C(\d+)', raw)
        if m:
            return _aci_to_hex(int(m.group(1)))
    return _get_entity_base_color(entity, doc)


def clean_dxf_text(raw: str) -> str:
    """
    Limpa uma fatia de texto MTEXT ja separada por paragrafo.
    Aplica plain_mtext apos remover \\pxq...; e depois limpa residuos.
    """
    if not raw:
        return ""
    text = _RE_PARA_FMT.sub('', raw)                        # remove \pxql; etc.
    text = plain_mtext(text)                                 # limpa \f \C \c {}
    text = re.sub(r'\\[a-zA-Z][^\\;\s]*;?', '', text)      # residuos
    text = re.sub(r'\\{2,}', '', text)                      # barras soltas
    text = re.sub(r'[ \t]+', ' ', text)
    return text.strip()


def _process_entity(entity, doc, base_color=None, block_name=None) -> list[dict]:
    """
    Transforma uma entidade DXF (TEXT, MTEXT, ATTRIB) em uma lista de dicionários formatados.
    Suporta quebra de parágrafos MTEXT e herança de cor.
    """
    tp = entity.dxftype()
    # TEXT e ATTRIB usam .dxf.text, MTEXT usa .text
    raw = entity.dxf.text if tp in ("TEXT", "ATTRIB") else entity.text

    # Cor base da entidade (sem codigos inline — esses sao por paragrafo)
    if base_color and entity.dxf.color == 0:
        entity_base = base_color
    else:
        entity_base = _get_entity_base_color(entity, doc)

    if tp == "TEXT" or tp == "ATTRIB":
        height = getattr(entity.dxf, "height", 0)
    else:
        height = getattr(entity.dxf, "char_height", 0)
    
    style = getattr(entity.dxf, "style", "Standard")
    pos = entity.dxf.insert

    if tp == "MTEXT":
        # Passo 1: remove \pxql; \pxqc; \pxqr; etc.
        pre = _RE_PARA_FMT.sub('', raw)
        # Passo 2: divide por \P, quebras de linha ou mudanças de cor internas
        # Lookahead (?=...) divide ANTES do código, Lookbehind (?<=\}) divide APÓS o bloco.
        # (?<!\{) garante que não dividimos entre o { e o código de cor (\C1;).
        parts = re.split(r'\\P|[\r\n]+|\^M|\^J|(?=\{\\[cC]\d+;)|(?<!\{)(?=\\[cC]\d+;)|(?<=\})', pre)
    else:
        parts = [raw]

    rows = []
    current_color = entity_base  # cor herdada entre paragrafos (\P)

    for part in parts:
        if not part or not part.strip():
            continue
        
        # Resolve a cor desta parte. 
        # Se estiver entre chaves { }, a cor é local.
        p_stripped = part.strip()
        is_block = p_stripped.startswith('{') and p_stripped.endswith('}')
        item_color = _resolve_part_color(part, current_color)
        
        if tp == "MTEXT":
            # Passo 4: limpa o texto do paragrafo/bloco.
            txt = clean_dxf_text(part)
        else:
            txt = re.sub(r'[ \t]+', ' ', part).strip()

        if txt:
            flags = f"DXF_{tp}"
            if block_name:
                flags += f" | BLOCO:{block_name}"

            rows.append({
                "pagina": 1,
                "texto": txt,
                "fonte": style,
                "tamanho": round(height, 2),
                "cor": item_color,
                "flags": flags,
                "layer": getattr(entity.dxf, "layer", ""),
                "_y": float(pos.y),
                "_x": float(pos.x),
            })
            
            # Se for uma quebra de paragrafo real ou mudança global (sem chaves), 
            # atualiza a cor herdada. Se for bloco { }, não atualiza a herança global.
            if not is_block:
                current_color = item_color
    return rows


def extract_dxf_content(doc) -> list[dict]:
    """
    Extrai todo o conteúdo textual do model space, blocos e atributos,
    incluindo a identificação dos nomes dos blocos.
    """
    msp = doc.modelspace()
    all_rows: list[dict] = []

    # 1. Entidades diretas no model space (Texto e MText)
    for entity in msp.query("TEXT MTEXT"):
        all_rows.extend(_process_entity(entity, doc))

    # 2. Entidades dentro de blocos (INSERT)
    for insert in msp.query("INSERT"):
        bname = insert.dxf.name
        try:
            block = doc.blocks.get(bname)
            bcolor = _get_entity_base_color(insert, doc)
            
            # Extrai o nome do bloco como uma entrada (útil para identificar símbolos)
            all_rows.append({
                "pagina": 1,
                "texto": f"[BLOCO: {bname}]",
                "fonte": "BlockName",
                "tamanho": 0,
                "cor": bcolor,
                "flags": "DXF_INSERT",
                "layer": getattr(insert.dxf, "layer", ""),
                "_y": float(insert.dxf.insert.y),
                "_x": float(insert.dxf.insert.x),
            })

            # Extrai textos, mtexts e atributos de DENTRO da definição do bloco
            for entity in block.query("TEXT MTEXT ATTRIB"):
                all_rows.extend(_process_entity(entity, doc, base_color=bcolor, block_name=bname))
            
            # Extrai atributos específicos desta INSTÂNCIA (insert.attribs)
            for attr in insert.attribs:
                all_rows.extend(_process_entity(attr, doc, base_color=bcolor, block_name=bname))
                
        except Exception as e:
            print(f"Erro ao processar bloco {bname}: {e}")
            pass

    # Ordenação: maior Y primeiro (topo), desempate por X (esquerda)
    all_rows.sort(key=lambda r: (-r["_y"], r["_x"]))
    for r in all_rows:
        r.pop("_y", None)
        r.pop("_x", None)

    return all_rows

# =====================================================
# API ENDPOINTS
# =====================================================

@app.post("/upload")
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
                if os.path.exists(tmp_path): os.remove(tmp_path)

        # PDF
        doc = pymupdf.open(stream=contents, filetype="pdf")
        extracted = []
        for i, page in enumerate(doc):
            blocks = page.get_text("dict", flags=11)["blocks"]
            for b in blocks:
                if "lines" not in b: continue
                for l in b["lines"]:
                    if "spans" not in l: continue
                    for s in l["spans"]:
                        text = s["text"].strip()
                        if text:
                            extracted.append({
                                "pagina": i + 1,
                                "texto": text,
                                "fonte": s["font"],
                                "tamanho": round(s["size"], 2),
                                "cor": color_to_hex(s.get("color", 0)),
                                "flags": flags_decomposer(s["flags"])
                            })
        return {"data": extracted}
    except Exception as e:
        return {"error": str(e), "data": []}

@app.get("/extract-local")
async def extract_local(path: str):
    if not os.path.exists(path): return {"error": "Arquivo não encontrado"}
    try:
        if path.lower().endswith(".dxf"):
            doc = ezdxf.readfile(path)
            return {"data": extract_dxf_content(doc), "filename": os.path.basename(path)}
        # PDF
        doc = pymupdf.open(path)
        extracted = []
        for i, page in enumerate(doc):
            blocks = page.get_text("dict", flags=11)["blocks"]
            for b in blocks:
                if "lines" not in b: continue
                for l in b["lines"]:
                    if "spans" not in l: continue
                    for s in l["spans"]:
                        text = s["text"].strip()
                        if text:
                            extracted.append({
                                "pagina": i + 1,
                                "texto": text,
                                "fonte": s["font"],
                                "tamanho": round(s["size"], 2),
                                "cor": color_to_hex(s.get("color", 0)),
                                "flags": flags_decomposer(s["flags"])
                            })
        return {"data": extracted, "filename": os.path.basename(path)}
    except Exception as e:
        return {"error": str(e), "data": []}

# =====================================================
# STATIC FILES
# =====================================================

if getattr(sys, "frozen", False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

static_dir = os.path.join(base_dir, "static")
if not os.path.exists(static_dir): os.makedirs(static_dir)

app.mount("/static", StaticFiles(directory=static_dir), name="static")

@app.get("/")
async def serve_index():
    index_path = os.path.join(static_dir, "index.html")
    if not os.path.exists(index_path): return HTMLResponse("index.html não encontrado na pasta static")
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read())

def open_browser(url):
    webbrowser.open(url)

if __name__ == "__main__":
    import uvicorn
    port = 8000
    if "PORT" not in os.environ:
        threading.Timer(1.5, open_browser, args=(f"http://127.0.0.1:{port}",)).start()
    uvicorn.run(app, host="0.0.0.0", port=port)