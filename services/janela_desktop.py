"""
services/janela_desktop.py — comportamento da janela do programa no computador (TASK-133).

* Fechar a janela pede confirmação (pywebview `confirm_close` com o texto em português) e, confirmado, ENCERRA o servidor local
  (antes o servidor podia continuar rodando em segundo plano e segurar a porta).
* Senhas salvas: a janela usa um perfil persistente (cookies, senhas e preenchimento automático ficam guardados em
  %APPDATA%\\GridReader\\janela) e liga no WebView2 o salvamento de senhas e o preenchimento automático — o login oferece as senhas
  já salvas e pergunta se quer salvar uma nova.
"""
import os
import threading

from config import logger

# textos da janela nativa (pywebview): a confirmação de fechar e os botões
LOCALIZACAO = {
    "global.quitConfirmation": "Fechar o GridReader? O servidor local também será encerrado.",
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
    p = os.path.join(base, "GridReader", "janela")
    antiga = os.path.join(base, "ReadGrid", "janela")          # TASK-135: o nome anterior; leva junto as senhas já salvas
    if not os.path.exists(p) and os.path.isdir(antiga):
        try:
            os.makedirs(os.path.dirname(p), exist_ok=True)
            os.rename(antiga, p)
        except OSError:
            pass
    os.makedirs(p, exist_ok=True)
    return p


def ligar_senhas_salvas(window) -> bool:
    """WebView2 (Windows): liga 'salvar senhas' e 'preenchimento automático' (vêm desligados no controle). Outros motores: nada a fazer."""
    try:
        form = window.native
        browser = getattr(form, "browser", None)
        webview2 = getattr(browser, "webview", None)
        if webview2 is None:
            return False
        from System import Action  # pythonnet (só no Windows)

        def ligar():
            core = webview2.CoreWebView2
            if core is None:
                return
            core.Settings.IsPasswordAutosaveEnabled = True
            core.Settings.IsGeneralAutofillEnabled = True

        form.Invoke(Action(ligar))
        return True
    except Exception as e:  # noqa: BLE001 — sem isso o login funciona igual, só não oferece as senhas salvas
        logger.info(f"Senhas salvas na janela indisponíveis neste computador: {e}")
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
