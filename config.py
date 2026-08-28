"""
config.py — Configurações centralizadas e resolução de caminhos.
Suporta modo dev (python app.py) e modo frozen (PyInstaller .exe).
"""
import os
import sys
import logging
from dotenv import load_dotenv

# Carrega as variáveis do .env
load_dotenv()

# ── Logging ─────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("leitor_pdf")


# ── Resolução de diretórios ─────────────────────────────────────────────────

def _get_base_dir() -> str:
    """Retorna diretório base: ao lado do .exe em modo frozen, ou ao lado do app.py em modo dev."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _get_resource_path(relative: str) -> str:
    """Resolve caminho de recurso empacotado (static, prompt, seed). Funciona em dev e frozen."""
    if getattr(sys, "frozen", False):
        # _MEIPASS = temp onde PyInstaller extrai datas
        meipass = getattr(sys, "_MEIPASS", _get_base_dir())
        # Tenta _MEIPASS primeiro (recurso empacotado), depois pasta do .exe (persistente)
        p1 = os.path.join(meipass, relative)
        if os.path.exists(p1):
            return p1
        return os.path.join(_get_base_dir(), relative)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative)


# ── Caminhos principais ─────────────────────────────────────────────────────

BASE_DIR = _get_base_dir()
DB_PATH = os.path.join(BASE_DIR, "banco_resumo.db")
PROMPT_PATH = _get_resource_path("prompt_rede_eletrica.txt")
SEED_CSV_PATH = _get_resource_path(os.path.join("data", "tabela_seed.csv"))

# ── Diretório static ────────────────────────────────────────────────────────

if getattr(sys, "frozen", False):
    _static_base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.executable)))
else:
    _static_base = os.path.dirname(os.path.abspath(__file__))

STATIC_DIR = os.path.join(_static_base, "static")
# Fallback: se static não estiver em _MEIPASS (ex: .exe portável), tenta ao lado do executável
if not os.path.exists(STATIC_DIR) and getattr(sys, "frozen", False):
    alt_static = os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "static")
    if os.path.exists(alt_static):
        STATIC_DIR = alt_static
if not os.path.exists(STATIC_DIR):
    os.makedirs(STATIC_DIR, exist_ok=True)

# ── Headers no-cache ────────────────────────────────────────────────────────

NO_CACHE_HEADERS = {
    "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
    "Pragma": "no-cache",
    "Expires": "0",
}

# ── Verificação de modo ─────────────────────────────────────────────────────

IS_FROZEN = getattr(sys, "frozen", False)
