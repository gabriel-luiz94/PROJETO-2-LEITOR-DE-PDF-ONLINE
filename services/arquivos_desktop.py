"""
services/arquivos_desktop.py — Ponte de arquivos da janela desktop (pywebview), TASK-142.

No executável o programa roda numa janela nativa (WebView2), onde baixar um arquivo (`<a download>`) e escolher um arquivo (`<input type=file>`) nem sempre funcionam.
Esta classe é entregue à janela como `js_api` e abre as caixas NATIVAS do Windows: `salvar_arquivo` (Salvar como) e `abrir_arquivos` (Abrir); ao SALVAR o navegador entrega os bytes em base64; ao ABRIR os bytes vêm por HTTP local (TASK-193) (static/desktop_arquivos.js troca os downloads e os seletores de arquivo por estas chamadas). Só existe na versão desktop; na online nada disto roda.
"""
import base64
import binascii
import os
import re
import secrets
import threading
import time

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


# TASK-193: «Abrir» devolve só nome, tamanho e um id; os bytes saem por HTTP local (GET /api/desktop/arquivo/<id>, routers/desktop_arquivos.py), binário e em fluxo.
# Antes o Python lia o arquivo todo, virava base64 (+33%) e atravessava a ponte JSON da janela: em DXF/PDF grande a tela ficava segundos sem reagir.
VALIDADE_ESCOLHIDOS_S = 600
_escolhidos = {}                  # id → (caminho, validade): SÓ arquivos que o usuário escolheu na caixa nativa desta sessão
_trava = threading.Lock()


def registrar_escolhido(caminho: str) -> str:
    """Guarda o caminho escolhido na caixa «Abrir» e devolve o id (aleatório, impossível de adivinhar) que a página usa para buscar os bytes."""
    agora = time.monotonic()
    with _trava:
        for k in [k for k, (_c, ate) in _escolhidos.items() if ate < agora]:
            del _escolhidos[k]
        i = secrets.token_urlsafe(24)
        _escolhidos[i] = (caminho, agora + VALIDADE_ESCOLHIDOS_S)
        return i


def caminho_escolhido(i: str):
    """O caminho registrado para o id (ainda válido e existente) ou None."""
    with _trava:
        c = _escolhidos.get(str(i or ""))
    if not c or c[1] < time.monotonic() or not os.path.isfile(c[0]):
        return None
    return c[0]


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

    def salvar_arquivo(self, nome, dados_base64, aceita="", diretorio=""):
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
            r = _escolhido(self._janela.create_file_dialog(self._dialogo("SAVE"), directory=diretorio if isinstance(diretorio, str) and os.path.isdir(diretorio) else "", save_filename=nome, file_types=filtros))
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
        """Abre o "Abrir" e devolve {arquivos: [{id, nome, tamanho}]} | {cancelado: True} | {ok: False, erro}. Não lê o arquivo: a página busca os bytes em /api/desktop/arquivo/<id>."""
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
                saida.append({"id": registrar_escolhido(caminho), "nome": os.path.basename(caminho), "tamanho": tamanho})
            return {"ok": True, "arquivos": saida}
        except Exception as e:  # noqa: BLE001
            logger.warning(f"abrir_arquivos falhou: {e}")
            return {"ok": False, "erro": str(e)}
