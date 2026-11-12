"""
services/auto_updater.py — Atualização do executável desktop a partir das Releases do GitHub (TASK-127). SÓ modo desktop; na versão online (servidor) nada daqui roda.

Fluxo:
  1. `check_for_updates()` consulta `releases/latest` do repositório (público, sem token), compara a tag (v2.1.0) com `APP_VERSION` e só oferece a
     atualização se a Release tiver o `GridR.exe` E o `GridR.exe.sha256` (gerados pelo workflow `.github/workflows/release-exe.yml`). TASK-152: o programa se chamava
     `Leitor_Projetos.exe`; a Release traz as DUAS cópias (mesmo binário) para quem ainda roda a versão antiga conseguir atualizar, e este código aceita o nome antigo se faltar o novo.
     O resultado fica em cache por 30 min (o limite do GitHub sem token é 60 consultas/hora por IP); offline ou com erro → "sem atualização", em silêncio.
  2. `download_and_install_update()` (1 clique no aviso): baixa o .exe para a pasta do programa, confere o SHA-256 com o da Release, grava um .bat auxiliar
     que espera o programa fechar, guarda o .exe atual como `.antigo`, põe o novo no lugar e reabre. O banco, o .env e as demais pastas ao lado do .exe não são tocados.
  Segurança: a URL de download NUNCA vem do navegador — é sempre a da Release consultada aqui, restrita a https://github.com/<repo>/releases/download/.
"""
import hashlib
import os
import re
import subprocess
import sys
import threading
import time

from config import APP_MODE, APP_VERSION, IS_FROZEN, logger

GITHUB_REPO = os.environ.get("GITHUB_REPO", "gabriel-luiz94/PROJETO-2-LEITOR-DE-PDF-ONLINE")
NOME_EXE = "GridR.exe"
NOME_SHA = NOME_EXE + ".sha256"
NOME_EXE_LEGADO = "Leitor_Projetos.exe"                  # nome até a TASK-152 (versões antigas só procuram por este)
NOME_SHA_LEGADO = NOME_EXE_LEGADO + ".sha256"
CACHE_SEGUNDOS = 30 * 60
TIMEOUT_CONSULTA = 6.0
TIMEOUT_DOWNLOAD = 60.0
URL_RELEASE = "https://api.github.com/repos/{repo}/releases/latest"

_cache = {"t": 0.0, "dados": None}
_aplicando = threading.Lock()


def _url_permitida() -> str:
    return f"https://github.com/{GITHUB_REPO}/releases/download/"


def desligado() -> bool:
    """Só no modo desktop; ATUALIZADOR=0 desliga (ex.: máquina sem internet permitida)."""
    return APP_MODE != "desktop" or os.environ.get("ATUALIZADOR", "1").strip() == "0"


def versao_tupla(texto):
    """'v2.1.0' / '2.1' → (2, 1, 0); None se não for uma versão numérica."""
    m = re.match(r"^\s*[vV]?(\d+)(?:\.(\d+))?(?:\.(\d+))?\s*$", str(texto or ""))
    return tuple(int(x or 0) for x in m.groups()) if m else None


def _get_json(url):
    import httpx
    r = httpx.get(url, headers={"Accept": "application/vnd.github+json", "User-Agent": "LeitorPro-Updater"}, timeout=TIMEOUT_CONSULTA, follow_redirects=True)
    r.raise_for_status()
    return r.json()


def _sem_atualizacao(motivo=""):
    return {"has_update": False, "motivo": motivo} if motivo else {"has_update": False}


def _interpretar(release: dict) -> dict:
    """Release do GitHub → resposta do /check (has_update só com versão MAIOR e com .exe + .sha256)."""
    nova, atual = versao_tupla(release.get("tag_name")), versao_tupla(APP_VERSION)
    if not nova or not atual or nova <= atual or release.get("draft") or release.get("prerelease"):
        return _sem_atualizacao()
    ativos = {a.get("name"): a.get("browser_download_url") for a in release.get("assets") or []}
    url, sha = ativos.get(NOME_EXE), ativos.get(NOME_SHA)
    if not url or not sha:
        url, sha = ativos.get(NOME_EXE_LEGADO), ativos.get(NOME_SHA_LEGADO)
    base = _url_permitida()
    if not url or not sha or not str(url).startswith(base) or not str(sha).startswith(base):
        return _sem_atualizacao("a Release não tem o executável e o SHA-256")
    notas = (release.get("body") or "").strip().replace("\r", "")
    return {
        "has_update": True,
        "latest_version": str(release.get("tag_name")).lstrip("vV"),
        "download_url": url,
        "sha256_url": sha,
        "release_notes": (notas[:240] + "…") if len(notas) > 240 else notas,
        "pode_aplicar": bool(IS_FROZEN),
    }


def check_for_updates(forcar: bool = False) -> dict:
    """Há versão nova? Nunca levanta erro (offline, limite do GitHub… → has_update False)."""
    if desligado():
        return _sem_atualizacao()
    if not forcar and _cache["dados"] is not None and time.monotonic() - _cache["t"] < CACHE_SEGUNDOS:
        return _cache["dados"]
    try:
        from services import rede_adaptativa
        if rede_adaptativa.offline():
            return _sem_atualizacao("sem internet")
    except Exception:  # noqa: BLE001
        pass
    try:
        dados = _interpretar(_get_json(URL_RELEASE.format(repo=GITHUB_REPO)))
    except Exception as e:  # noqa: BLE001 — o aviso de atualização nunca atrapalha o programa
        logger.warning(f"Não foi possível verificar atualizações: {e}")
        return _sem_atualizacao("erro na consulta")
    _cache.update(t=time.monotonic(), dados=dados)
    return dados


# ─────────────────────────── download e troca do .exe ───────────────────────────
def _baixar(url: str, destino: str) -> str:
    """Baixa `url` para `destino` (em streaming) e devolve o SHA-256 (hex minúsculo) do que foi gravado."""
    import httpx
    h = hashlib.sha256()
    with httpx.stream("GET", url, headers={"User-Agent": "LeitorPro-Updater"}, timeout=TIMEOUT_DOWNLOAD, follow_redirects=True) as r:
        r.raise_for_status()
        with open(destino, "wb") as f:
            for bloco in r.iter_bytes(1024 * 256):
                h.update(bloco)
                f.write(bloco)
    return h.hexdigest()


def _sha_esperado(url: str) -> str:
    import httpx
    r = httpx.get(url, headers={"User-Agent": "LeitorPro-Updater"}, timeout=TIMEOUT_CONSULTA, follow_redirects=True)
    r.raise_for_status()
    m = re.search(r"\b([0-9a-fA-F]{64})\b", r.text)
    if not m:
        raise ValueError("o arquivo .sha256 da Release está num formato inesperado")
    return m.group(1).lower()


def gerar_script(pid: int, alvo: str, novo: str, antigo: str) -> str:
    """.bat que espera o programa fechar, guarda o .exe atual como .antigo, coloca o novo e reabre (volta ao antigo se a troca falhar)."""
    ASPAS3 = '"' * 3                                   # find """%PID%""" procura o PID entre aspas na saída CSV do tasklist (não confunde com outros números)
    return f"""@echo off
chcp 65001 >nul
set "PYINSTALLER_RESET_ENVIRONMENT=1"
set "PID={pid}"
set "ALVO={alvo}"
set "NOVO={novo}"
set "ANTIGO={antigo}"
set N=0
:esperar
tasklist /FI "PID eq %PID%" /NH /FO CSV 2>nul | find {ASPAS3}%PID%{ASPAS3} >nul
if errorlevel 1 goto trocar
set /a N+=1
if %N% GEQ 120 goto falhou
ping -n 2 127.0.0.1 >nul
goto esperar
:trocar
set N=0
:tentar
if exist "%ANTIGO%" del /f /q "%ANTIGO%" >nul 2>&1
move /y "%ALVO%" "%ANTIGO%" >nul 2>&1
if not errorlevel 1 goto colocar
set /a N+=1
if %N% GEQ 60 goto falhou
ping -n 2 127.0.0.1 >nul
goto tentar
:colocar
move /y "%NOVO%" "%ALVO%" >nul 2>&1
if errorlevel 1 move /y "%ANTIGO%" "%ALVO%" >nul 2>&1
goto abrir
:falhou
if exist "%NOVO%" del /f /q "%NOVO%" >nul 2>&1
:abrir
start "" "%ALVO%"
(goto) 2>nul & del /f /q "%~f0"
"""


def _encerrar():
    os._exit(0)


def _iniciar_script(bat: str):
    # CREATE_NO_WINDOW (e NÃO DETACHED_PROCESS): o .bat roda num console oculto que os comandos dele herdam — sem janelas piscando e com `ping` esperando de verdade
    flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | getattr(subprocess, "CREATE_NO_WINDOW", 0)
    env = dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT="1")          # o novo .exe não pode reaproveitar a pasta temporária deste
    subprocess.Popen(["cmd", "/c", bat], creationflags=flags, close_fds=True, env=env)


def download_and_install_update(download_url=None) -> bool:
    """Baixa a Release mais recente (a URL do pedido é IGNORADA), confere o SHA-256, grava o .bat auxiliar e agenda o fechamento do programa."""
    if desligado():
        raise RuntimeError("As atualizações automáticas só funcionam no aplicativo desktop.")
    if not IS_FROZEN:
        raise RuntimeError("Atualização automática só no executável (.exe); em desenvolvimento use git pull.")
    if not _aplicando.acquire(blocking=False):
        raise RuntimeError("Já existe uma atualização em andamento.")
    try:
        info = check_for_updates(forcar=True)
        if not info.get("has_update"):
            raise RuntimeError("Não há atualização disponível.")
        alvo = os.path.abspath(sys.executable)
        pasta = os.path.dirname(alvo)
        novo, antigo = alvo + ".novo", alvo + ".antigo"
        esperado = _sha_esperado(info["sha256_url"])
        logger.info(f"Baixando a versão {info['latest_version']}...")
        try:
            obtido = _baixar(info["download_url"], novo)
        except Exception:
            _remover(novo)
            raise
        if obtido != esperado:
            _remover(novo)
            raise RuntimeError("O arquivo baixado não confere com o SHA-256 da Release; a atualização foi cancelada.")
        bat = os.path.join(pasta, "atualizar_leitor.bat")
        with open(bat, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(gerar_script(os.getpid(), alvo, novo, antigo))
        _iniciar_script(bat)
        logger.info("Atualização baixada e conferida; o programa vai fechar e reabrir na nova versão.")
        threading.Timer(1.5, _encerrar).start()           # deixa a resposta HTTP sair antes de fechar
        return True
    finally:
        _aplicando.release()


def _remover(caminho):
    try:
        if os.path.exists(caminho):
            os.remove(caminho)
    except OSError:
        pass


def limpar_residuos(atraso: float = 120.0):
    """Depois de abrir na versão nova: apaga o .antigo (e sobras) da atualização anterior, com atraso — o .antigo é o salva-vidas se a nova não abrir."""
    if not IS_FROZEN or desligado():
        return

    def _limpar():
        alvo = os.path.abspath(sys.executable)
        for sufixo in (".antigo", ".novo"):
            _remover(alvo + sufixo)
        _remover(os.path.join(os.path.dirname(alvo), "atualizar_leitor.bat"))

    t = threading.Timer(atraso, _limpar)
    t.daemon = True
    t.start()
