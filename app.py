"""
app.py — Ponto de entrada da aplicação FastAPI (v2.0 - Profissionalizado).

Suporta dois modos de operação:
  - "desktop": app local (.exe ou dev), CORS aberto, permite /extract-local
  - "server":  deploy na nuvem, CORS restrito, HTTPS, sem acesso local
"""

import sys as _sys
if "--trabalhador-leitor-js" in _sys.argv:   # processo filho do modo autônomo (no .exe, sys.executable é o próprio programa)
    from services.autonomo.trabalhador_js import main as _trabalhador_main
    _trabalhador_main(_sys.argv[1:])
    _sys.exit(0)

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
from routers import manual, obras, regras, regras_leitor, regras_vinculacao, vinculacao_config, atalhos, modais_montaveis, barra_modais, rede, recs, projetos, orcamento, ai_chat, upload, health, auth, admin, update, validacao, validacao_prompts, validacao_regras, validacao_ajustes, autonomo, aprendizado, plugin_cad


app = FastAPI(
    title="GridReader",
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

@app.get("/static/resumo_layout.js")
async def serve_resumo_layout_js():
    return _serve_js("resumo_layout.js")

@app.get("/static/ativo_sugestao.js")
async def serve_ativo_sugestao_js():
    return _serve_js("ativo_sugestao.js")       # TASK-136


@app.get("/static/abrir_na_app.js")
async def serve_abrir_na_app_js():
    return _serve_js("abrir_na_app.js")         # TASK-128


@app.get("/static/campos_segredo.js")
async def serve_campos_segredo_js():
    return _serve_js("campos_segredo.js")       # TASK-128


@app.get("/static/marca.js")
async def serve_marca_js():
    return _serve_js("marca.js")                # TASK-132


@app.get("/static/plugin_nanocad.js")
async def serve_plugin_nanocad_js():
    return _serve_js("plugin_nanocad.js")       # TASK-131


@app.get("/static/barra_modais.js")
async def serve_barra_modais_js():
    return _serve_js("barra_modais.js")         # TASK-117

@app.get("/static/modal_montavel_ui.js")
async def serve_modal_montavel_ui_js():
    return _serve_js("modal_montavel_ui.js")    # TASK-117

@app.get("/static/gerador_modal.js")
async def serve_gerador_modal_js():
    return _serve_js("gerador_modal.js")        # TASK-117

@app.get("/static/base_grade.js")
async def serve_base_grade_js():
    return _serve_js("base_grade.js")

@app.get("/static/script.js")
async def serve_script_js():
    return _serve_js("script.js")

@app.get("/static/login.js")
async def serve_login_js():
    return _serve_js("login.js")

@app.get("/static/admin.js")
async def serve_admin_js():
    return _serve_js("admin.js")

@app.get("/static/autonomo.js")
async def serve_autonomo_js():
    return _serve_js("autonomo.js")

@app.get("/static/voz.js")
async def serve_voz_js():
    return _serve_js("voz.js")

@app.get("/static/painel_regras.js")
async def serve_painel_regras_js():
    return _serve_js("painel_regras.js")

@app.get("/static/regras_arquivo.js")
async def serve_regras_arquivo_js():
    return _serve_js("regras_arquivo.js")

@app.get("/static/painel_ajustes.js")
async def serve_painel_ajustes_js():
    return _serve_js("painel_ajustes.js")

@app.get("/static/painel_prompts.js")
async def serve_painel_prompts_js():
    return _serve_js("painel_prompts.js")

@app.get("/static/regras_editor.js")
async def serve_regras_editor_js():
    return _serve_js("regras_editor.js")

@app.get("/static/ler_projeto.js")
async def serve_ler_projeto_js():
    return _serve_js("ler_projeto.js")

@app.get("/static/ativo_parametros.js")
async def serve_ativo_parametros_js():
    return _serve_js("ativo_parametros.js")


@app.get("/static/ativo_avaliar.js")
async def serve_ativo_avaliar_js():
    return _serve_js("ativo_avaliar.js")       # TASK-123


@app.get("/static/ativo_dica.js")
async def serve_ativo_dica_js():
    return _serve_js("ativo_dica.js")          # TASK-122


@app.get("/static/regras_leitor_engine.js")
async def serve_regras_leitor_engine_js():
    # TASK-082: o motor e as regras andam juntos (regras novas com motor antigo em cache quebram a leitura), então o motor nunca fica em cache
    return _serve_js("regras_leitor_engine.js")


@app.get("/static/vinculacao_regras.js")
async def serve_vinculacao_regras_js():
    return _serve_js("vinculacao_regras.js")


@app.get("/static/vinculacao_ui.js")
async def serve_vinculacao_ui_js():
    return _serve_js("vinculacao_ui.js")


@app.get("/static/anotacao_cabos.js")
async def serve_anotacao_cabos_js():
    return _serve_js("anotacao_cabos.js")


@app.get("/static/anotacao_outros.js")
async def serve_anotacao_outros_js():
    return _serve_js("anotacao_outros.js")

@app.get("/static/grade_tabela.js")
async def serve_grade_tabela_js():
    return _serve_js("grade_tabela.js")

@app.get("/static/atalhos_modal.js")
async def serve_atalhos_modal_js():
    return _serve_js("atalhos_modal.js")

@app.get("/static/gerador_modal.js")
async def serve_gerador_modal_js():
    return _serve_js("gerador_modal.js")

@app.get("/static/modal_montavel_engine.js")
async def serve_modal_montavel_engine_js():
    return _serve_js("modal_montavel_engine.js")

@app.get("/static/modal_montavel_editor.js")
async def serve_modal_montavel_editor_js():
    return _serve_js("modal_montavel_editor.js")

@app.get("/static/modal_montavel_arquivo.js")
async def serve_modal_montavel_arquivo_js():
    return _serve_js("modal_montavel_arquivo.js")

@app.get("/static/modal_montavel_ui.js")
async def serve_modal_montavel_ui_js():
    return _serve_js("modal_montavel_ui.js")

@app.get("/static/orcamento_sap.js")
async def serve_orcamento_sap_js():
    return _serve_js("orcamento_sap.js")

@app.get("/static/orcamento_comparar.js")
async def serve_orcamento_comparar_js():
    return _serve_js("orcamento_comparar.js")

@app.get("/static/base_ficha_dados.js")
async def serve_base_ficha_dados_js():
    return _serve_js("base_ficha_dados.js")


@app.get("/static/base_ficha_ui.js")
async def serve_base_ficha_ui_js():
    return _serve_js("base_ficha_ui.js")


@app.get("/static/base_cadastro_dados.js")
async def serve_base_cadastro_dados_js():
    return _serve_js("base_cadastro_dados.js")


@app.get("/static/base_cadastro_ui.js")
async def serve_base_cadastro_ui_js():
    return _serve_js("base_cadastro_ui.js")


@app.get("/static/desktop_arquivos.js")
async def serve_desktop_arquivos_js():
    return _serve_js("desktop_arquivos.js")


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

@app.get("/manual")
async def serve_manual():
    caminho = os.path.join(STATIC_DIR, "manual.html")
    if not os.path.exists(caminho):
        return HTMLResponse("manual.html não encontrado na pasta static")
    with open(caminho, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers=NO_CACHE_HEADERS)

@app.get("/autonomo")
async def serve_autonomo():
    caminho = os.path.join(STATIC_DIR, "autonomo.html")
    if not os.path.exists(caminho):
        return HTMLResponse("autonomo.html não encontrado na pasta static")
    with open(caminho, "r", encoding="utf-8") as f:
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
app.include_router(manual.router)
app.include_router(autonomo.router)
app.include_router(aprendizado.router)
app.include_router(auth.router)
app.include_router(obras.router)
app.include_router(regras.router)
app.include_router(regras_leitor.router)
app.include_router(regras_vinculacao.router)
app.include_router(vinculacao_config.router)
app.include_router(plugin_cad.router)          # TASK-104: plugin do nanoCAD (regras, versão, conferência com a base)
app.include_router(atalhos.router)
app.include_router(modais_montaveis.router)
app.include_router(barra_modais.router)           # TASK-117: configuração da barra de modais por projeto
app.include_router(rede.router)                    # TASK-121: estado da rede / cópia local (desktop)
app.include_router(recs.router)
app.include_router(projetos.router)
app.include_router(orcamento.router)
app.include_router(ai_chat.router)
app.include_router(validacao_prompts.router)
app.include_router(validacao_regras.router)
app.include_router(validacao_ajustes.router)
app.include_router(validacao.router)
app.include_router(upload.router)
app.include_router(health.router)
app.include_router(admin.router)
app.include_router(update.router)


# ── Rede adaptativa + cópia local do banco (TASK-121) — SÓ no modo desktop; na versão online (APP_MODE=server) nada disto roda ──
@app.on_event("startup")
def _iniciar_rede_adaptativa():
    try:
        from services import rede_adaptativa, copia_local
        est = rede_adaptativa.iniciar()               # testa proxy × direto ANTES de qualquer conexão e liga a vigia
        if est is None:
            return
        rede_adaptativa.registrar_ouvinte("online", lambda: copia_local.iniciar_em_segundo_plano("a nuvem voltou"))
        if est["online"]:
            copia_local.iniciar_em_segundo_plano("início")      # em segundo plano: o programa abre e usa a nuvem enquanto copia
    except Exception as e:  # noqa: BLE001 — a rede adaptativa nunca impede o programa de abrir
        logger.warning(f"Rede adaptativa não iniciou: {e}")


# ── Atualização pelo GitHub (TASK-127): depois de abrir, apaga as sobras da atualização anterior (só o .exe desktop) ──
@app.on_event("startup")
def _limpar_residuos_da_atualizacao():
    try:
        from services import auto_updater
        auto_updater.limpar_residuos()
    except Exception as e:  # noqa: BLE001
        logger.warning(f"Limpeza da atualização anterior falhou: {e}")


# ── Plugin do nanoCAD (TASK-131): o .exe instala/atualiza a DLL do plugin em %APPDATA%\OrcamentoNanoCAD\plugin (em segundo plano) ──
@app.on_event("startup")
def _instalar_plugin_nanocad():
    try:
        from services import plugin_instalador
        plugin_instalador.instalar_em_segundo_plano()
    except Exception as e:  # noqa: BLE001
        logger.warning(f"Instalação do plugin do nanoCAD não iniciou: {e}")


# ── Conectividade com a nuvem (TASK-040): ping periódico + reconexão proativa + indicador online/offline ──
@app.on_event("startup")
def _iniciar_monitor_conectividade():
    try:
        from services.connectivity_monitor import start_connectivity_monitor
        start_connectivity_monitor()
    except Exception as e:  # noqa: BLE001 — o monitor nunca impede o programa de abrir
        logger.warning(f"Monitor de conectividade não iniciou: {e}")


# ── Modo autônomo (TASK-031): religa a vigia da pasta no início se ela estava LIGADA na configuração ──
@app.on_event("startup")
def _iniciar_modo_autonomo():
    try:
        from services.autonomo import config_autonomo, pasta
        cfg = config_autonomo.carregar()
        if cfg.get("ligado") and cfg.get("user_id"):
            config_autonomo.garantir_pastas(cfg)
            pasta.obter_vigia().iniciar()
            logger.info("Modo autônomo ligado (vigiando %s).", config_autonomo.pastas(cfg)["entrada"])
    except Exception as e:  # noqa: BLE001 — o modo autônomo nunca impede o programa de abrir
        logger.warning(f"Modo autônomo não iniciou: {e}")


@app.on_event("shutdown")
def _parar_modo_autonomo():
    try:
        from services.autonomo import pasta
        pasta.obter_vigia().parar(3)
    except Exception:  # noqa: BLE001
        pass


# ── Inicialização Desktop / Servidor ─────────────────────────────────────────

_servidor_uvicorn = None     # TASK-133: guardado para encerrar quando a janela fecha


def run_uvicorn(host: str, port: int):
    """Executa o servidor Uvicorn em uma thread separada."""
    global _servidor_uvicorn
    import uvicorn
    config = uvicorn.Config(app, host=host, port=port, log_level="warning")
    server = uvicorn.Server(config)
    _servidor_uvicorn = server
    server.run()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    is_headless = os.environ.get("HEADLESS", "0") == "1" or APP_MODE == "server"

    logger.info(f"Iniciando GridReader v{APP_VERSION} [modo: {APP_MODE}]...")

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
        # TASK-143: o plugin do nanoCAD acha o programa (programa.json) e pode abri-lo já com uma obra (--abrir-obra <arquivo>)
        from services import programa_local
        programa_local.registrar(port, APP_VERSION)
        _arquivo_obra = programa_local.ler_argumento(sys.argv[1:])
        if _arquivo_obra:
            _id_obra = programa_local.carregar_arquivo(_arquivo_obra)
            if _id_obra:
                from urllib.parse import quote as _quote
                login_url += "&next=" + _quote(programa_local.url_pendente(_id_obra), safe="")

        from services.inicio_janela import TELA_ERRO, TELA_INICIAL, esperar_servidor

        try:
            import webview
            try:
                webview.settings['ALLOW_DOWNLOADS'] = True          # TASK-142: sem isto o WebView2 cancela qualquer download (a ponte de arquivos abaixo cobre os blobs)
            except Exception:  # noqa: BLE001 — pywebview antigo, sem essa opção
                pass
            from services.arquivos_desktop import ApiArquivos
            api_arquivos = ApiArquivos()
            # TASK-128: a janela abre já com a tela "Iniciando…" e só carrega o login quando o servidor aceita conexões (sem a mensagem "recusou a conectar")
            window = webview.create_window(
                title=f"GridReader v{APP_VERSION}",
                html=TELA_INICIAL,
                width=1320,
                height=860,
                min_size=(960, 640),
                background_color="#0d1117",
                text_select=True,
                zoomable=True,
                confirm_close=True,          # TASK-133: pergunta antes de fechar (texto em janela_desktop.LOCALIZACAO)
                js_api=api_arquivos,         # TASK-142: salvar/abrir arquivos pelas caixas nativas (static/desktop_arquivos.js)
            )
            api_arquivos.vincular(window)
            from services import janela_desktop
            janela_desktop.vincular(window)              # TASK-143: o plugin do nanoCAD traz a janela para a frente com a obra

            def _abrir_quando_pronto(win):
                threading.Thread(target=janela_desktop.ligar_senhas_salvas, args=(win,), daemon=True).start()      # TASK-133/151: login oferece as senhas salvas (espera o WebView2 ficar pronto, sem travar a abertura)
                if esperar_servidor("127.0.0.1", port):
                    win.load_url(login_url)
                else:
                    win.load_html(TELA_ERRO)

            # perfil persistente: senhas salvas e preenchimento automático continuam de uma vez para outra
            webview.start(_abrir_quando_pronto, window, private_mode=False, storage_path=janela_desktop.pasta_perfil(),
                          localization=janela_desktop.LOCALIZACAO)
            # a janela fechou (o usuário confirmou): encerra o servidor local junto
            janela_desktop.encerrar_servidor(_servidor_uvicorn, server_thread)
        except Exception as e:
            logger.warning(f"Janela nativa indisponível ({e}). Abrindo no navegador...")
            esperar_servidor("127.0.0.1", port)
            webbrowser.open(login_url)
            server_thread.join()