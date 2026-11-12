"""
services/arquivos_desktop.py — Ponte de arquivos da janela desktop (pywebview), TASK-142.

No executável o programa roda numa janela nativa (WebView2), onde baixar um arquivo (`<a download>`) e escolher um arquivo (`<input type=file>`) nem sempre funcionam.
Esta classe é entregue à janela como `js_api` e abre as caixas NATIVAS do Windows: `salvar_arquivo` (Salvar como) e `abrir_arquivos` (Abrir); o navegador entrega/recebe
os bytes em base64 (static/desktop_arquivos.js troca os downloads e os seletores de arquivo por estas chamadas). Só existe na versão desktop; na online nada disto roda.
"""
import base64
import binascii
import os
import re

from config import logger

LIMITE_BYTES = 200 * 1024 * 1024
_MIME_EXT = {"application/json": "json", "text/csv": "csv", "application/pdf": "pdf", "text/plain": "txt", "application/zip": "zip",
             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx"}


def filtros_de_arquivo(aceita: str, rotulo: str = "Arquivos permitidos") -> tuple:
    """`accept` do input (".pdf, .dxf, application/json") → file_types do pywebview ("Arquivos permitidos (*.pdf;*.dxf;*.json)", "Todos os arquivos (*.*)")."""
    exts = []
    for parte in re.split(r"[,\s]+", str(aceita or "")):
        parte = parte.strip().lower()
        if re.fullmatch(r"\.[a-z0-9]{1,8}", parte):
            exts.append(parte[1:])
        elif parte in _MIME_EXT:
            exts.append(_MIME_EXT[parte])
    exts = list(dict.fromkeys(exts))
    if not exts:
        return ("Todos os arquivos (*.*)",)
    return (f"{rotulo} ({';'.join('*.' + e for e in exts)})", "Todos os arquivos (*.*)")


def _escolhido(r):
    """Resultado de create_file_dialog (None, str ou lista/tupla) → lista de caminhos."""
    if not r:
        return []
    return [str(x) for x in ([r] if isinstance(r, str) else list(r)) if x]


class ApiArquivos:
    """js_api da janela: só métodos públicos aparecem para o JavaScript (a janela fica em atributo privado)."""

    def __init__(self):
        self._janela = None

    def vincular(self, janela):
        self._janela = janela

    def _dialogo(self, tipo: str):
        import webview
        fd = getattr(webview, "FileDialog", None)                       # pywebview 5+; antes eram as constantes SAVE_DIALOG/OPEN_DIALOG
        return getattr(fd, tipo) if fd is not None and hasattr(fd, tipo) else getattr(webview, tipo + "_DIALOG")

    def salvar_arquivo(self, nome, dados_base64, aceita=""):
        """Abre o "Salvar como" e grava os bytes. → {ok, caminho} | {cancelado: True} | {ok: False, erro}."""
        try:
            if self._janela is None:
                return {"ok": False, "erro": "janela indisponível"}
            nome = os.path.basename(str(nome or "arquivo")) or "arquivo"
            try:
                dados = base64.b64decode(str(dados_base64 or ""), validate=False)
            except (binascii.Error, ValueError):
                return {"ok": False, "erro": "conteúdo inválido"}
            if len(dados) > LIMITE_BYTES:
                return {"ok": False, "erro": "arquivo grande demais"}
            ext = os.path.splitext(nome)[1]
            filtros = filtros_de_arquivo(aceita or ext, "Arquivo")
            r = _escolhido(self._janela.create_file_dialog(self._dialogo("SAVE"), directory="", save_filename=nome, file_types=filtros))
            if not r:
                return {"cancelado": True}
            caminho = r[0]
            if ext and not os.path.splitext(caminho)[1]:
                caminho += ext
            with open(caminho, "wb") as f:
                f.write(dados)
            return {"ok": True, "caminho": caminho}
        except Exception as e:  # noqa: BLE001 — a ponte nunca derruba a janela
            logger.warning(f"salvar_arquivo falhou: {e}")
            return {"ok": False, "erro": str(e)}

    def abrir_arquivos(self, aceita="", multiplos=False):
        """Abre o "Abrir" e devolve {arquivos: [{nome, tipo, base64}]} | {cancelado: True} | {ok: False, erro}."""
        try:
            if self._janela is None:
                return {"ok": False, "erro": "janela indisponível"}
            r = _escolhido(self._janela.create_file_dialog(self._dialogo("OPEN"), allow_multiple=bool(multiplos), file_types=filtros_de_arquivo(aceita)))
            if not r:
                return {"cancelado": True}
            saida, total = [], 0
            for caminho in r:
                tamanho = os.path.getsize(caminho)
                total += tamanho
                if total > LIMITE_BYTES:
                    return {"ok": False, "erro": "arquivo(s) grande(s) demais"}
                with open(caminho, "rb") as f:
                    saida.append({"nome": os.path.basename(caminho), "base64": base64.b64encode(f.read()).decode("ascii")})
            return {"ok": True, "arquivos": saida}
        except Exception as e:  # noqa: BLE001
            logger.warning(f"abrir_arquivos falhou: {e}")
            return {"ok": False, "erro": str(e)}
