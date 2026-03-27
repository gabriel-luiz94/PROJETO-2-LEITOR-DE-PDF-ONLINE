from fastapi import FastAPI, UploadFile, File, WebSocket, WebSocketDisconnect
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import pymupdf
import io
import os
import sys
import threading
import webbrowser
import multiprocessing
import json

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Manage WebSocket connections
class ConnectionManager:
    def __init__(self):
        self.active_connections: list[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        if websocket in self.active_connections:
            self.active_connections.remove(websocket)

    async def broadcast(self, message: str):
        for connection in self.active_connections:
            try:
                await connection.send_text(message)
            except:
                pass

manager = ConnectionManager()

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    try:
        while True:
            await websocket.receive_text() # Just keep connection alive
    except WebSocketDisconnect:
        manager.disconnect(websocket)

@app.get("/trigger-file")
async def trigger_file(path: str):
    await manager.broadcast(json.dumps({"type": "load_file", "path": path}))
    return {"status": "success"}

def color_to_hex(c):
    """Robust conversion of PyMuPDF color values to hex string."""
    if c is None: return "#000000"
    try:
        # If it's a tuple/list (e.g. (r,g,b) in floats)
        if isinstance(c, (list, tuple)):
            if len(c) >= 3:
                rgb = [max(0, min(255, int(x * 255))) for x in c[:3]]
                return "#%02x%02x%02x" % tuple(rgb)
            elif len(c) == 1:
                g = max(0, min(255, int(c[0] * 255)))
                return "#%02x%02x%02x" % (g, g, g)
        
        # If it's an integer (standard RGB packed or signed 32-bit)
        val = int(c)
        # Mask with 0xFFFFFF to handle signed ints correctly and get RRGGBB
        return "#%06x" % (val & 0xFFFFFF)
    except:
        return "#000000"

def flags_decomposer(flags):
    """Make font flags human readable."""
    l = []
    if flags & 2 ** 0:
        l.append("superscript")
    if flags & 2 ** 1:
        l.append("italic")
    if flags & 2 ** 2:
        l.append("serifed")
    else:
        l.append("sans")
    if flags & 2 ** 3:
        l.append("monospaced")
    else:
        l.append("proportional")
    if flags & 2 ** 4:
        l.append("bold")
    return ", ".join(l)

@app.post("/upload")
async def upload_pdf(file: UploadFile = File(...)):
    contents = await file.read()
    # Read the PDF from binary stream
    doc = pymupdf.open(stream=contents, filetype="pdf")
    
    extracted_data = []
    pagina = 1
    for page in doc:
        blocks = page.get_text("dict", flags=11)["blocks"]
        for b in blocks:
            if "lines" not in b:
                continue
            for l in b["lines"]:
                if "spans" not in l:
                    continue
                for s in l["spans"]:
                    text = s["text"].strip()
                    if not text:
                        continue
                    color_hex = color_to_hex(s.get("color", 0))
                    extracted_data.append({
                        "pagina": pagina,
                        "texto": text,
                        "fonte": s["font"],
                        "tamanho": round(s["size"], 2),
                        "cor": color_hex,
                        "flags": flags_decomposer(s["flags"])
                    })
        pagina += 1
    return {"data": extracted_data}

@app.get("/extract-local")
async def extract_local(path: str):
    if not os.path.exists(path):
        return {"error": "Arquivo não encontrado"}
    
    # Read the PDF from local path
    doc = pymupdf.open(path)
    
    extracted_data = []
    pagina = 1
    for page in doc:
        blocks = page.get_text("dict", flags=11)["blocks"]
        for b in blocks:
            if "lines" not in b:
                continue
            for l in b["lines"]:
                if "spans" not in l:
                    continue
                for s in l["spans"]:
                    text = s["text"].strip()
                    if not text:
                        continue
                    color_hex = color_to_hex(s.get("color", 0))
                    extracted_data.append({
                        "pagina": pagina,
                        "texto": text,
                        "fonte": s["font"],
                        "tamanho": round(s["size"], 2),
                        "cor": color_hex,
                        "flags": flags_decomposer(s["flags"])
                    })
        pagina += 1
    return {"data": extracted_data, "filename": os.path.basename(path)}

# Mount static files at /static/ so WebSocket routes are not blocked
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

static_dir = os.path.join(base_dir, 'static')
app.mount("/static", StaticFiles(directory=static_dir), name="static")

@app.get("/")
async def serve_index():
    index_path = os.path.join(static_dir, "index.html")
    with open(index_path, "r", encoding="utf-8") as f:
        content = f.read()
    return HTMLResponse(content=content)

def open_browser(url="http://127.0.0.1:8000"):
    webbrowser.open(url)

if __name__ == "__main__":
    import uvicorn
    # Use environment port if available (for Render/Railway hosting), otherwise default to 8000
    port = int(os.environ.get("PORT", 8000))
    # In local development (no PORT set), open browser
    if "PORT" not in os.environ:
        threading.Timer(1.5, open_browser, args=(f"http://127.0.0.1:{port}",)).start()
    
    uvicorn.run(app, host="0.0.0.0", port=port)
