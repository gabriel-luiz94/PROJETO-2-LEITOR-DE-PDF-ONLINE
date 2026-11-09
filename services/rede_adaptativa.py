"""
services/rede_adaptativa.py — Adaptação automática à rede (TASK-121). SÓ no modo desktop (APP_MODE=desktop); na versão online (servidor) nada daqui roda.

Problema: o programa desktop usa a nuvem (Supabase) primeiro e o banco local só como reserva. Em rede com proxy da empresa (ou com um proxy
"esquecido" de outra rede), cada consulta à nuvem esperava ~21 s de timeout antes de cair no banco local.

O que este módulo faz:
  1. Ao iniciar (antes de qualquer conexão): descobre o proxy configurado (variáveis HTTPS_PROXY/HTTP_PROXY e configuração do Windows), TESTA o
     Supabase de verdade pelo caminho com proxy e pelo direto (em paralelo, timeout curto) e fica com o que responder. Aplica a escolha ao processo
     inteiro (inclusive o chat de IA): proxy → HTTPS_PROXY/HTTP_PROXY; direto → NO_PROXY=* (ignora também o proxy do Windows).
  2. Com o programa aberto: uma vigia leve reavalia sozinha (a cada ~30 s online; a cada 5–30 s offline) e na hora quando uma chamada falha por rede
     (`sinalizar_falha`). Se a rede mudou (casa ↔ empresa) ela troca de caminho; se nenhum caminho responde, entra em MODO OFFLINE.
  3. Disjuntor: enquanto offline, `get_supabase()` devolve None, e todo o programa vai direto ao banco local (sem esperar timeout).
     Ao voltar a nuvem, os ouvintes registrados (cópia/mescla do banco local) rodam em segundo plano.
Proxy com usuário e senha e arquivo PAC (AutoConfigURL) ficam fora.
"""
import os
import socket
import threading
import time
import urllib.request
from urllib.parse import urlparse

from config import APP_MODE, logger

INTERVALO_ONLINE = 30.0            # segundos entre checagens com a nuvem OK
INTERVALO_OFFLINE_MAX = 30.0       # offline: 5 s, 10 s … até 30 s
TIMEOUT_SONDA = 3.0                # cada teste ao Supabase
TIMEOUT_TCP_PROXY = 1.5            # o proxy responde? (conexão TCP)
INTERVALO_MIN_FALHA = 2.0          # várias falhas seguidas disparam uma só reavaliação

_VARS_PROXY = ("HTTPS_PROXY", "HTTP_PROXY", "ALL_PROXY", "https_proxy", "http_proxy", "all_proxy")
_VARS_SEM_PROXY = ("NO_PROXY", "no_proxy")


def _ler_proxy_do_sistema():
    """Proxy configurado (variáveis de ambiente primeiro, depois o registro do Windows) ou None. Lido UMA vez, antes de qualquer alteração."""
    try:
        p = urllib.request.getproxies()
        return p.get("https") or p.get("http") or p.get("all") or None
    except Exception:  # noqa: BLE001
        return None


_PROXY_ORIGINAL = _ler_proxy_do_sistema()

_lock = threading.Lock()
_estado = {"iniciado": False, "online": True, "caminho": "direto", "proxy": None, "desde": time.time(), "ultima_checagem": 0.0, "motivo": ""}
_acordar = threading.Event()
_parar = threading.Event()
_thread = None
_ultima_reavaliacao = 0.0
_ouvintes = {"online": [], "offline": [], "mudanca": []}


# ─────────────────────────── configuração ───────────────────────────
def habilitada() -> bool:
    """Só no modo desktop (e nunca sob teste), com o Supabase configurado; REDE_ADAPTATIVA=0 desliga."""
    if APP_MODE != "desktop" or os.environ.get("REDE_ADAPTATIVA", "1").strip() == "0" or os.environ.get("PYTEST_CURRENT_TEST"):
        return False
    return bool(os.environ.get("SUPABASE_URL", "").strip() and os.environ.get("SUPABASE_KEY", "").strip())


def ativa() -> bool:
    return _estado["iniciado"]


def offline() -> bool:
    """True com a vigia ligada e a nuvem inalcançável: o programa deve usar só o banco local."""
    return _estado["iniciado"] and not _estado["online"]


def estado() -> dict:
    with _lock:
        e = dict(_estado)
    e["proxy"] = _proxy_sem_credenciais(e.get("proxy"))
    e["ativa"] = e.pop("iniciado")
    return e


def registrar_ouvinte(evento: str, funcao):
    """evento: 'online' (nuvem voltou), 'offline' (nuvem caiu) ou 'mudanca' (qualquer troca de estado/caminho; recebe o estado)."""
    _ouvintes[evento].append(funcao)


def _proxy_sem_credenciais(url):
    if not url:
        return None
    try:
        p = urlparse(url)
        return f"{p.hostname}:{p.port}" if p.port else (p.hostname or None)
    except Exception:  # noqa: BLE001
        return None


# ─────────────────────────── testes de rede ───────────────────────────
def _cliente_http(proxy, timeout):
    import httpx
    kw = dict(timeout=httpx.Timeout(timeout), trust_env=False, follow_redirects=False)
    if proxy:
        try:
            return httpx.Client(proxy=proxy, **kw)           # httpx ≥ 0.26
        except TypeError:
            return httpx.Client(proxies=proxy, **kw)         # httpx antigo
    return httpx.Client(**kw)


def _proxy_responde(proxy) -> bool:
    try:
        p = urlparse(proxy)
        with socket.create_connection((p.hostname, p.port or 80), timeout=TIMEOUT_TCP_PROXY):
            return True
    except Exception:  # noqa: BLE001
        return False


def _sondar(proxy) -> bool:
    """O Supabase responde por este caminho (proxy ou direto)? Resposta HTTP (até 401/404) conta; erro de rede/timeout, 5xx e 407 reprovam."""
    url = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
    if not url:
        return False
    try:
        if proxy and not _proxy_responde(proxy):
            return False
        with _cliente_http(proxy, TIMEOUT_SONDA) as c:
            r = c.get(url + "/rest/v1/", headers={"apikey": os.environ.get("SUPABASE_KEY", "").strip()})
        return r.status_code < 500 and r.status_code != 407        # 5xx (nuvem/proxy fora do ar) e 407 (proxy pede login) não valem
    except Exception:  # noqa: BLE001
        return False


def avaliar(proxy_configurado=None):
    """Testa os caminhos possíveis EM PARALELO e escolhe: (online, caminho, proxy). Prefere o proxy configurado quando ele funciona."""
    proxy = proxy_configurado if proxy_configurado is not None else _PROXY_ORIGINAL
    candidatos = ([("proxy", proxy)] if proxy else []) + [("direto", None)]
    resultados = {}

    def _testar(nome, prx):
        resultados[nome] = _sondar(prx)

    ts = [threading.Thread(target=_testar, args=c, daemon=True) for c in candidatos]
    for t in ts:
        t.start()
    for t in ts:
        t.join(TIMEOUT_SONDA + TIMEOUT_TCP_PROXY + 1.0)
    if proxy and resultados.get("proxy"):
        return True, "proxy", proxy
    if resultados.get("direto"):
        return True, "direto", None
    return False, "nenhum", proxy


def _aplicar(caminho, proxy):
    """Aplica o caminho ao processo todo e recria o cliente do Supabase (que lê o ambiente ao ser criado)."""
    for k in _VARS_PROXY + _VARS_SEM_PROXY:
        os.environ.pop(k, None)
    if caminho == "proxy" and proxy:
        for k in ("HTTPS_PROXY", "HTTP_PROXY", "https_proxy", "http_proxy"):
            os.environ[k] = proxy
    elif caminho == "direto":
        os.environ["NO_PROXY"] = os.environ["no_proxy"] = "*"          # httpx: "*" desliga o proxy (inclusive o do Windows)
    try:
        from services.supabase_client import reset_supabase_client
        reset_supabase_client()
    except Exception:  # noqa: BLE001
        pass


# ─────────────────────────── estado e vigia ───────────────────────────
def _descrever(caminho, proxy):
    return {"proxy": f"proxy da empresa ({_proxy_sem_credenciais(proxy)})", "direto": "acesso direto", "nenhum": "sem nuvem"}.get(caminho, caminho)


def _atualizar(online, caminho, proxy, motivo=""):
    """Grava o novo estado; se mudou, aplica o caminho e avisa os ouvintes (fora do lock)."""
    with _lock:
        antes = (_estado["online"], _estado["caminho"], _estado["proxy"])
        _estado.update(online=online, caminho=caminho, proxy=proxy, ultima_checagem=time.time(), motivo=motivo)
        mudou = antes != (online, caminho, proxy)
        if mudou:
            _estado["desde"] = time.time()
    if not mudou:
        return
    if online:
        _aplicar(caminho, proxy)                           # caminho novo, ou a nuvem voltou: cliente novo (conexão limpa)
    logger.info(f"Rede: {'ONLINE' if online else 'OFFLINE (modo local)'} — {_descrever(caminho, proxy)}.")
    for ev in ([("online" if online else "offline")] + ["mudanca"]):
        for f in list(_ouvintes[ev]):
            try:
                f(estado()) if ev == "mudanca" else f()
            except Exception as e:  # noqa: BLE001 — um ouvinte nunca derruba a vigia
                logger.warning(f"Ouvinte de rede ({ev}) falhou: {e}")


def _reavaliar():
    global _ultima_reavaliacao
    _ultima_reavaliacao = time.monotonic()
    atual = estado()
    if atual["online"]:
        prx = _PROXY_ORIGINAL if atual["caminho"] == "proxy" else None
        if _sondar(prx):                                   # teste barato pelo caminho em uso
            with _lock:
                _estado["ultima_checagem"] = time.time()
            return
    online, caminho, proxy = avaliar()
    _atualizar(online, caminho, proxy)


def _loop():
    falhas = 0
    while not _parar.is_set():
        espera = INTERVALO_ONLINE if _estado["online"] else min(INTERVALO_OFFLINE_MAX, 5.0 * (1 + falhas))
        _acordar.wait(espera)
        _acordar.clear()
        if _parar.is_set():
            break
        try:
            _reavaliar()
            falhas = 0 if _estado["online"] else falhas + 1
        except Exception as e:  # noqa: BLE001
            logger.warning(f"Vigia de rede falhou: {e}")


def sinalizar_falha():
    """Uma chamada ao Supabase falhou por rede: reavalia já (em vez de esperar o próximo intervalo)."""
    if _estado["iniciado"] and (time.monotonic() - _ultima_reavaliacao) > INTERVALO_MIN_FALHA:
        _acordar.set()


def iniciar():
    """Chamar UMA vez na inicialização, antes de qualquer conexão. Devolve o estado, ou None se desligada (versão online, sem Supabase, teste)."""
    global _thread
    if not habilitada() or _estado["iniciado"]:
        return None
    online, caminho, proxy = avaliar()
    with _lock:
        _estado.update(iniciado=True, online=online, caminho=caminho, proxy=proxy, desde=time.time(), ultima_checagem=time.time())
    if online:
        _aplicar(caminho, proxy)
    logger.info(f"Rede ao iniciar: {'ONLINE' if online else 'OFFLINE (modo local)'} — {_descrever(caminho, proxy)}"
                f"{' | proxy configurado: ' + (_proxy_sem_credenciais(_PROXY_ORIGINAL) or '') if _PROXY_ORIGINAL else ' | sem proxy configurado'}.")
    _parar.clear()
    _thread = threading.Thread(target=_loop, daemon=True, name="RedeAdaptativa")
    _thread.start()
    return estado()


def parar():
    """Para a vigia e desliga o disjuntor (usado em testes)."""
    _parar.set()
    _acordar.set()
    with _lock:
        _estado.update(iniciado=False, online=True, caminho="direto", proxy=None)
