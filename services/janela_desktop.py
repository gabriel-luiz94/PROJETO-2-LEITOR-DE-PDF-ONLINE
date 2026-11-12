"""
services/janela_desktop.py — comportamento da janela do programa no computador (TASK-133).

* Fechar a janela pede confirmação (pywebview `confirm_close` com o texto em português) e, confirmado, ENCERRA o servidor local
  (antes o servidor podia continuar rodando em segundo plano e segurar a porta).
* Senhas salvas: a janela usa um perfil persistente (cookies, senhas e preenchimento automático ficam guardados em
  %APPDATA%\\GridR\\janela) e liga no WebView2 o salvamento de senhas e o preenchimento automático — o login oferece as senhas
  já salvas e pergunta se quer salvar uma nova.
"""
import os
import threading
import time

from config import logger

# textos da janela nativa (pywebview): a confirmação de fechar e os botões
LOCALIZACAO = {
    "global.quitConfirmation": "Fechar o GridR? O servidor local também será encerrado.",
    "global.ok": "Fechar",
    "global.cancel": "Cancelar",
    "global.quit": "Sair",
    "global.saveFile": "Salvar arquivo",
    "windows.fileFilter.allFiles": "Todos os arquivos",
    "windows.fileFilter.otherFiles": "Outros arquivos",
}


def pasta_perfil() -> str:
    """Perfil persistente do navegador da janela (senhas salvas, preenchimento automático, cookies)."""
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    p = os.path.join(base, "GridR", "janela")
    for antigo in ("GridReader", "ReadGrid"):                  # TASK-135/152: os nomes anteriores; leva junto as senhas já salvas
        antiga = os.path.join(base, antigo, "janela")
        if not os.path.exists(p) and os.path.isdir(antiga):
            try:
                os.makedirs(os.path.dirname(p), exist_ok=True)
                os.rename(antiga, p)
            except OSError:
                pass
    os.makedirs(p, exist_ok=True)
    return p


def ligar_senhas_salvas(window, tentativas: int = 80, intervalo: float = 0.25) -> bool:
    """WebView2 (Windows): liga 'salvar senhas' e 'preenchimento automático' (vêm desligados no controle). Outros motores: nada a fazer.

    TASK-151: o WebView2 inicializa em segundo plano — logo depois de abrir a janela `CoreWebView2` ainda é None e antes a função desistia em silêncio (e as senhas nunca eram
    salvas nem preenchidas). Agora espera o controle ficar pronto (até ~20 s) e só então aplica; devolve True quando aplicou. Rode em uma thread (não bloqueia a abertura)."""
    ultimo = None
    for _ in range(max(1, tentativas)):
        try:
            form = window.native
            browser = getattr(form, "browser", None)
            webview2 = getattr(browser, "webview", None)
            if webview2 is None:
                ultimo = "controle do navegador ainda não existe (ou não é WebView2)"
            else:
                from System import Action  # pythonnet (só no Windows)
                resultado = []

                def ligar():
                    core = webview2.CoreWebView2
                    if core is None:
                        resultado.append(False)
                        return
                    core.Settings.IsPasswordAutosaveEnabled = True
                    core.Settings.IsGeneralAutofillEnabled = True
                    resultado.append(True)

                form.Invoke(Action(ligar))
                if resultado and resultado[0]:
                    logger.info("Senhas salvas e preenchimento automático ligados na janela.")
                    return True
                ultimo = "WebView2 ainda inicializando"
        except (AttributeError, ImportError) as e:      # runtime sem essa configuração / não é Windows: tentar de novo não adianta
            logger.info(f"Senhas salvas na janela indisponíveis neste computador: {e}")
            return False
        except Exception as e:  # noqa: BLE001 — a janela ainda pode estar abrindo
            ultimo = str(e)
        time.sleep(intervalo)
    logger.info(f"Senhas salvas na janela não foram ligadas ({ultimo}).")
    return False


def encerrar_servidor(servidor, thread: threading.Thread, espera: float = 5.0):
    """Depois que a janela fecha: pede ao uvicorn para parar (roda os 'shutdown'), espera um pouco e encerra o processo inteiro."""
    try:
        if servidor is not None:
            servidor.should_exit = True
        if thread is not None:
            thread.join(espera)
    except Exception as e:  # noqa: BLE001
        logger.warning(f"Encerramento do servidor: {e}")
    logger.info("Janela fechada: servidor local encerrado.")
    os._exit(0)


# ── a janela atende o plugin do nanoCAD (TASK-143): carrega a obra recebida e vem para a frente ──
_janela = None


def vincular(window) -> None:
    global _janela
    _janela = window


def mostrar(url_relativa: str, porta: int = None) -> bool:
    """Leva a janela do programa para a URL (ex.: /resumo?plugin_pendente=…) e a traz para a frente. False se não há janela."""
    if _janela is None:
        return False
    try:
        base = _janela.get_current_url() or ""
    except Exception:  # noqa: BLE001
        base = ""
    if base.startswith("http"):
        from urllib.parse import urlsplit
        p = urlsplit(base)
        destino = f"{p.scheme}://{p.netloc}{url_relativa}"
    else:
        destino = f"http://127.0.0.1:{porta or int(os.environ.get('PORT', 8000))}{url_relativa}"
    try:
        _janela.load_url(destino)
    except Exception as e:  # noqa: BLE001
        logger.warning(f"Janela não carregou a obra do nanoCAD: {e}")
        return False
    for passo in (lambda: _janela.restore(), lambda: _janela.show()):
        try:
            passo()
        except Exception:  # noqa: BLE001 — pywebview antigo: segue
            pass
    try:                                     # "pisca" no topo para passar na frente do nanoCAD
        _janela.on_top = True
        _janela.on_top = False
    except Exception:  # noqa: BLE001
        pass
    return True
