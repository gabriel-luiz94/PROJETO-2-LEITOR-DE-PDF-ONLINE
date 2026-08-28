"""
app.py — Ponto de entrada da aplicação FastAPI (refatorado).
"""
import os
import sys
import threading
import webbrowser
from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

# Configurações e Banco
from config import STATIC_DIR, NO_CACHE_HEADERS, logger
from database import init_db

# WebSocket
from websocket_manager import manager

# Routers
from routers import obras, regras, recs, projetos, orcamento, ai_chat, upload, health, auth, admin


app = FastAPI(title="Leitor de Projetos Online Pro")

# ── CORS ────────────────────────────────────────────────────────────────────
# Em modo frozen (.exe), mantemos aberto para uso local (localhost, file://)
# Quando formos para a nuvem/login, restringiremos isso.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Inicialização do Banco ──────────────────────────────────────────────────
init_db()

# ── Rotas Estáticas (Frontend) ──────────────────────────────────────────────
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/")
async def serve_index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    if not os.path.exists(index_path): 
        return HTMLResponse("index.html não encontrado na pasta static")
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers=NO_CACHE_HEADERS)

@app.get("/login")
async def serve_login():
    login_path = os.path.join(STATIC_DIR, "login.html")
    if not os.path.exists(login_path): 
        return HTMLResponse("login.html não encontrado")
    with open(login_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers=NO_CACHE_HEADERS)


@app.get("/resumo")
async def serve_resumo():
    resumo_path = os.path.join(STATIC_DIR, "resumo.html")
    if not os.path.exists(resumo_path): 
        return HTMLResponse("resumo.html não encontrado na pasta static")
    with open(resumo_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers=NO_CACHE_HEADERS)

@app.get("/static/resumo.js")
async def serve_resumo_js():
    js_path = os.path.join(STATIC_DIR, "resumo.js")
    if not os.path.exists(js_path): 
        return HTMLResponse("resumo.js não encontrado")
    with open(js_path, "r", encoding="utf-8") as f:
        from fastapi.responses import Response
        return Response(content=f.read(), media_type="application/javascript", headers=NO_CACHE_HEADERS)

@app.get("/admin")
async def serve_admin():
    admin_path = os.path.join(STATIC_DIR, "admin.html")
    if not os.path.exists(admin_path): 
        return HTMLResponse("admin.html não encontrado na pasta static")
    with open(admin_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers=NO_CACHE_HEADERS)

# ── WebSocket ───────────────────────────────────────────────────────────────
@app.websocket("/ws")
async def websocket_endpoint(websocket):
    await manager.connect(websocket)
    try:
        while True:
            await websocket.receive_text()
    except Exception:
        manager.disconnect(websocket)

# ── Montagem dos Routers ────────────────────────────────────────────────────
app.include_router(auth.router)
app.include_router(obras.router)
app.include_router(regras.router)
app.include_router(recs.router)
app.include_router(projetos.router)
app.include_router(orcamento.router)
app.include_router(ai_chat.router)
app.include_router(upload.router)
app.include_router(health.router)
app.include_router(admin.router)


def open_browser(url):
    webbrowser.open(url)


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    if "PORT" not in os.environ:
        threading.Timer(1.5, open_browser, args=(f"http://127.0.0.1:{port}",)).start()
    
    logger.info("Iniciando Leitor de Projetos Pro...")
    uvicorn.run(app, host="0.0.0.0", port=port)