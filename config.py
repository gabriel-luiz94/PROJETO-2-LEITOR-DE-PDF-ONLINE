"""
config.py — Configurações centralizadas e resolução de caminhos.
Suporta modo dev (python app.py) e modo frozen (PyInstaller .exe).
"""

# ── Versão da Aplicação ─────────────────────────────────────────────────────
APP_VERSION = "2.8.3"
import os
import sys
import logging
import secrets
import tempfile
from pathlib import Path
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
# No .exe o load_dotenv() acima não enxerga a pasta do executável; carrega o .env ao lado dele
# (não sobrescreve variáveis já definidas, então em dev nada muda).
load_dotenv(os.path.join(BASE_DIR, ".env"))
DB_PATH = os.path.join(BASE_DIR, "banco_resumo.db")
PROMPT_PATH = _get_resource_path("prompt_rede_eletrica.txt")
PROMPT_OBRAS_PATH = _get_resource_path(os.path.join("data", "prompt_obras.txt"))
PROMPT_ASSISTENTE_PATH = _get_resource_path(os.path.join("data", "prompt_assistente.txt"))      # TASK-191: assistente com ferramentas
SKILLS_DIR_PACOTE = _get_resource_path(os.path.join("data", "skills"))                           # skills que acompanham o programa (versionadas)
# Skills escritas/geradas pelo usuário: no .exe ficam AO LADO dele (persistem entre atualizações) e vencem as do pacote; em dev, a mesma pasta do pacote.
SKILLS_DIR_USUARIO = os.path.join(BASE_DIR, "skills") if getattr(sys, "frozen", False) else SKILLS_DIR_PACOTE
SEED_CSV_PATH = _get_resource_path(os.path.join("data", "tabela_seed.csv"))
REGRAS_LEITOR_PROCESSAMENTO_SEED_PATH = _get_resource_path(os.path.join("data", "regras_leitor_processamento_seed.json"))
REGRAS_LEITOR_CLASSIFICACAO_SEED_PATH = _get_resource_path(os.path.join("data", "regras_leitor_classificacao_seed.json"))
REGRAS_LEITOR_PROCESSAMENTO_PE_SEED_PATH = _get_resource_path(os.path.join("data", "regras_leitor_processamento_pernambuco_seed.json"))
REGRAS_LEITOR_CLASSIFICACAO_PE_SEED_PATH = _get_resource_path(os.path.join("data", "regras_leitor_classificacao_pernambuco_seed.json"))
VINCULACAO_PE_SEED_PATH = _get_resource_path(os.path.join("data", "vinculacao_pernambuco_seed.json"))
REGRAS_VINCULACAO_SEED_PATH = _get_resource_path(os.path.join("data", "regras_vinculacao_seed.json"))
REGIONAIS_PE_PATH = _get_resource_path(os.path.join("data", "regionais_pernambuco.json"))      # TASK-165
VALIDACOES_SEED_DIR = _get_resource_path(os.path.join("data", "validacoes"))
REGRAS_DOMINIO_SEED_PATH = os.path.join(VALIDACOES_SEED_DIR, "regras_dominio_seed.json")
AJUSTES_SEED_PATH = os.path.join(VALIDACOES_SEED_DIR, "ajustes_seed.json")

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

# ── Modo de operação ────────────────────────────────────────────────────────
# "desktop" = app local (.exe ou dev), "server" = deploy na nuvem
APP_MODE = os.environ.get("APP_MODE", "desktop" if IS_FROZEN else "desktop")

# ── JWT ─────────────────────────────────────────────────────────────────────
def _segredo_jwt() -> str:
    configurado = os.environ.get("JWT_SECRET", "").strip()
    if len(configurado.encode("utf-8")) >= 32 and configurado != "gere-uma-chave-secreta-forte-aqui":
        return configurado
    if APP_MODE == "server" or os.environ.get("HEADLESS") == "1":
        raise RuntimeError("Configure JWT_SECRET com um segredo aleatório de pelo menos 32 bytes.")

    # Fica no perfil local para que «Lembrar de mim» sobreviva a reinícios e atualizações.
    perfil = Path(os.environ.get("APPDATA") or Path.home() / ".config") / "OrcaCore"
    perfil.mkdir(mode=0o700, parents=True, exist_ok=True)
    destino = perfil / "jwt.key"
    if not destino.is_file():
        with tempfile.NamedTemporaryFile(dir=perfil, mode="w", encoding="ascii", delete=False) as arquivo:
            temporario = Path(arquivo.name)
            arquivo.write(secrets.token_hex(32))
        try:
            # Publica o arquivo completo sem sobrescrever a chave criada por outra instância.
            if os.name == "nt":
                temporario.rename(destino)
            else:
                os.link(temporario, destino)
        except FileExistsError:
            pass
        finally:
            temporario.unlink(missing_ok=True)
    segredo = destino.read_text(encoding="ascii").strip()
    if len(segredo) != 64 or any(c not in "0123456789abcdef" for c in segredo):
        raise RuntimeError("O arquivo local de sessão do OrçaCore está inválido.")
    return segredo


JWT_SECRET = _segredo_jwt()

# ── URLs do servidor central (para auto-update e sync) ──────────────────────
SERVER_URL = os.environ.get("SERVER_URL", "")
MANUAL_PATH = _get_resource_path(os.path.join("data", "manual_regras_e_ajustes.md"))
