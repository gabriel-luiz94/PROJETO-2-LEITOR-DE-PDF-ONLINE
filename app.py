"""
app.py — Ponto de entrada da aplicação FastAPI (v2.0 - Profissionalizado).

Suporta dois modos de operação:
  - "desktop": app local (.exe ou dev), CORS aberto, permite /extract-local
  - "server":  deploy na nuvem, CORS restrito, HTTPS, sem acesso local
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
from config import STATIC_DIR, NO_CACHE_HEADERS, APP_VERSION, APP_MODE, logger
from database import init_db

# Middleware de Autenticação
from middleware.auth_middleware import AuthMiddleware

# Middleware de No-Cache para HTML estático
from middleware.nocache_middleware import NoCacheHtmlMiddleware

# WebSocket
from websocket_manager import manager

# Routers
from routers import obras, regras, recs, projetos, orcamento, ai_chat, upload, health, auth, admin, update


app = FastAPI(
    title="Leitor de Projetos Online Pro",
    version=APP_VERSION,
    description="Sistema profissional de leitura e análise de projetos PDF/DXF",
)

# ── CORS ────────────────────────────────────────────────────────────────────
if APP_MODE == "server":
    # Modo servidor: CORS restrito a origens conhecidas
    allowed_origins = os.environ.get("CORS_ORIGINS", "").split(",")
    allowed_origins = [o.strip() for o in allowed_origins if o.strip()]
    if not allowed_origins:
        allowed_origins = ["*"]  # Fallback temporário
else:
    # Modo desktop: CORS aberto para uso local (localhost, file://)
    allowed_origins = ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Middleware de Autenticação JWT ───────────────────────────────────────────
app.add_middleware(AuthMiddleware)

# ── Middleware de No-Cache para HTML estático ──────────────────────────────
# Garante que /static/*.html nunca seja servido com cache pelo navegador.
app.add_middleware(NoCacheHtmlMiddleware)

# ── Inicialização do Banco ──────────────────────────────────────────────────
init_db()

# ── Rotas JS sem cache (devem vir ANTES do app.mount para ter prioridade) ──────
# O StaticFiles mount captura /static/* por prefixo; rotas explícitas registradas
# antes do mount têm precedência no Starlette e permitem servir com NO_CACHE_HEADERS.
from fastapi.responses import Response as _JSResponse

def _serve_js(filename: str):
    """Abre e serve um arquivo JS com cabeçalhos de no-cache."""
    js_path = os.path.join(STATIC_DIR, filename)
    if not os.path.exists(js_path):
        return HTMLResponse(f"{filename} não encontrado", status_code=404)
    with open(js_path, "r", encoding="utf-8") as f:
        return _JSResponse(content=f.read(), media_type="application/javascript", headers=NO_CACHE_HEADERS)

@app.get("/static/resumo.js")
async def serve_resumo_js():
    return _serve_js("resumo.js")

@app.get("/static/script.js")
async def serve_script_js():
    return _serve_js("script.js")

@app.get("/static/login.js")
async def serve_login_js():
    return _serve_js("login.js")

@app.get("/static/admin.js")
async def serve_admin_js():
    return _serve_js("admin.js")

@app.get("/static/auth_fetch.js")
async def serve_auth_fetch_js():
    return _serve_js("auth_fetch.js")

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

@app.get("/resultado_orcamento")
async def serve_resultado_orcamento():
    resultado_path = os.path.join(STATIC_DIR, "resultado_orcamento.html")
    if not os.path.exists(resultado_path): 
        return HTMLResponse("resultado_orcamento.html não encontrado na pasta static")
    with open(resultado_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers=NO_CACHE_HEADERS)

@app.get("/orcamento")
async def serve_orcamento():
    orc_path = os.path.join(STATIC_DIR, "orcamento.html")
    if not os.path.exists(orc_path): 
        return HTMLResponse("orcamento.html não encontrado na pasta static")
    with open(orc_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers=NO_CACHE_HEADERS)

@app.get("/admin")
async def serve_admin():
    admin_path = os.path.join(STATIC_DIR, "admin.html")
    if not os.path.exists(admin_path): 
        return HTMLResponse("admin.html não encontrado na pasta static")
    with open(admin_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers=NO_CACHE_HEADERS)

# ── Endpoint de Versão ──────────────────────────────────────────────────────
@app.get("/api/version")
async def get_version():
    return {"version": APP_VERSION, "mode": APP_MODE}

# ── Endpoint de Encerramento (apenas modo desktop) ───────────────────────────
@app.get("/api/shutdown")
async def shutdown():
    """Encerra o processo do servidor. Bloqueado em modo server (segurança)."""
    if APP_MODE == "server":
        from fastapi import HTTPException
        raise HTTPException(status_code=403, detail="Encerramento remoto não permitido no modo servidor.")
    import signal
    logger.info("Encerramento solicitado pelo usuário via /api/shutdown.")
    # Agenda o encerramento em thread separada para a resposta HTTP chegar antes
    def _encerrar():
        import time
        time.sleep(0.5)
        os.kill(os.getpid(), signal.SIGTERM)
    threading.Thread(target=_encerrar, daemon=True).start()
    return {"status": "encerrando"}

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
app.include_router(update.router)


# ── Inicialização Desktop / Servidor ─────────────────────────────────────────

def run_uvicorn(host: str, port: int):
    """Executa o servidor Uvicorn em uma thread separada."""
    import uvicorn
    config = uvicorn.Config(app, host=host, port=port, log_level="warning")
    server = uvicorn.Server(config)
    server.run()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    is_headless = os.environ.get("HEADLESS", "0") == "1" or APP_MODE == "server"

    logger.info(f"Iniciando Leitor de Projetos Pro v{APP_VERSION} [modo: {APP_MODE}]...")

    if is_headless:
        # Modo Servidor / Cloud (Deploy sem janela gráfica)
        import uvicorn
        uvicorn.run(app, host="0.0.0.0", port=port)
    else:
        # Modo Desktop Nativo (Janela de aplicativo independente)
        server_thread = threading.Thread(target=run_uvicorn, args=("127.0.0.1", port), daemon=True)
        server_thread.start()

        # Abre sempre na tela de login com ?new_session=1 para limpar sessão anterior
        app_url = f"http://127.0.0.1:{port}"
        login_url = f"{app_url}/login?new_session=1"

        try:
            import webview
            window = webview.create_window(
                title=f"Leitor de Projetos Pro v{APP_VERSION}",
                url=login_url,
                width=1320,
                height=860,
                min_size=(960, 640),
                background_color="#0d1117",
                text_select=True,
                zoomable=True
            )
            webview.start(private_mode=False)
        except Exception as e:
            logger.warning(f"Janela nativa indisponível ({e}). Abrindo no navegador...")
            webbrowser.open(login_url)
            server_thread.join()