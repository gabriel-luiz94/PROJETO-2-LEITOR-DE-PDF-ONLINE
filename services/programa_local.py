"""
services/programa_local.py — o plugin do nanoCAD entrega a obra ao GridReader instalado no computador (TASK-143).

Antes, "Abrir na aplicação" abria uma aba no navegador. Agora:

  * Programa ABERTO: o plugin acha o servidor local (GET /api/plugin/desktop) e manda a obra (POST /api/plugin/receber, só do próprio
    computador). A obra fica guardada aqui, em memória (PENDENTES), e a janela do programa vem para a frente em
    /resumo?plugin_pendente=<id>; o Resumo, já com o usuário logado, grava a obra para ele (POST /api/plugin/obra) e a abre.
  * Programa FECHADO: ao abrir, o .exe registra onde está e em que porta (%APPDATA%\\OrcamentoNanoCAD\\programa.json). O plugin grava a obra
    num arquivo e abre o .exe com "--abrir-obra <arquivo>"; a janela começa no login e, depois de entrar, abre a obra.

TASK-162 (integração com o PROJETO DE TRABALHO): o plugin grava o projeto `<desenho>.obra` na pasta do desenho e entrega só o CAMINHO
(POST /api/plugin/receber {chave, caminho} · ou "--abrir-projeto <caminho>" com o programa fechado). A janela abre esse projeto pelo
mesmo caminho do menu 📁 Projeto → Abrir (Abrir/Mesclar/Sobrescrever); nada é gravado no banco.
"""
import json
import os
import sys
import threading
import time
import uuid

from config import logger

VALIDADE_S = 3600
LIMITE_BYTES = 20 * 1024 * 1024
_pendentes: dict = {}
_trava = threading.Lock()


def pasta_dados() -> str:
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "OrcamentoNanoCAD")


def arquivo_registro() -> str:
    return os.path.join(pasta_dados(), "programa.json")


def registrar(porta: int, versao: str) -> None:
    """O .exe diz ao plugin onde está e em que porta atende (só o programa empacotado)."""
    if not getattr(sys, "frozen", False) and not os.environ.get("GRIDREADER_REGISTRAR"):
        return
    try:
        os.makedirs(pasta_dados(), exist_ok=True)
        with open(arquivo_registro(), "w", encoding="utf-8") as f:
            json.dump({"exe": sys.executable, "porta": int(porta), "versao": versao, "pid": os.getpid()}, f, ensure_ascii=False, indent=2)
    except OSError as e:
        logger.warning(f"Não deu para registrar o programa para o plugin do nanoCAD: {e}")


def validar(chave, envelope) -> str:
    """Erro (texto) ou '' se a obra recebida tem a forma do .obra.json."""
    if not isinstance(envelope, dict) or envelope.get("tipo") != "obra" or not isinstance(envelope.get("dados"), dict):
        return "a obra recebida não está no formato do .obra.json"
    if len(json.dumps(envelope, ensure_ascii=False)) > LIMITE_BYTES:
        return "obra grande demais"
    if not isinstance(chave, str) or not chave.strip():
        return "falta a chave do desenho"
    return ""


def guardar(chave: str, envelope: dict) -> str:
    """Guarda a obra recebida até a janela buscá-la; devolve o id."""
    agora = time.time()
    with _trava:
        for k in [k for k, v in _pendentes.items() if agora - v["quando"] > VALIDADE_S]:
            del _pendentes[k]
        i = uuid.uuid4().hex
        _pendentes[i] = {"chave": chave.strip(), "envelope": envelope, "quando": agora}
    return i


def validar_projeto(caminho) -> tuple:
    """(caminho absoluto, código do projeto de regras, '') ou (None, None, erro): `.obra` existente e legível no formato do programa."""
    from services import projeto_obra
    if not isinstance(caminho, str) or not caminho.lower().endswith(projeto_obra.EXTENSAO):
        return None, None, "o projeto precisa ser um arquivo .obra"
    caminho = os.path.abspath(caminho)
    if not os.path.isfile(caminho):
        return None, None, "o arquivo do projeto não existe"
    try:
        with open(caminho, "rb") as f:
            bruto = f.read(projeto_obra.LIMITE_BYTES + 1)
        carga = json.loads(bruto.decode("utf-8-sig"))
        codigo = str(carga.get("projeto") or "").strip()
        projeto_obra.ler(bruto, codigo)                              # mesmo exame do menu Abrir
        return caminho, codigo, ""
    except projeto_obra.ErroProjeto as e:
        return None, None, " ".join(e.mensagens)
    except (OSError, ValueError, UnicodeDecodeError, AttributeError):
        return None, None, "o arquivo do projeto está ilegível"


def guardar_projeto(chave: str, caminho: str, codigo: str) -> str:
    """Guarda o pedido de abrir o projeto `caminho` até a janela buscá-lo; devolve o id."""
    agora = time.time()
    with _trava:
        i = uuid.uuid4().hex
        _pendentes[i] = {"chave": (chave or "").strip(), "caminho": caminho, "projeto": codigo, "quando": agora}
    return i


def ler_argumento_projeto(argv) -> str:
    """'--abrir-projeto <caminho.obra>' na linha de comando do .exe (aberto pelo plugin com o programa fechado)."""
    for k, a in enumerate(argv):
        if a == "--abrir-projeto" and k + 1 < len(argv):
            return argv[k + 1]
        if a.startswith("--abrir-projeto="):
            return a.split("=", 1)[1]
    return ""


def pegar(i: str):
    with _trava:
        v = _pendentes.get(i or "")
        if v and time.time() - v["quando"] > VALIDADE_S:
            del _pendentes[i]
            return None
        return v


def url_pendente(i: str) -> str:
    return f"/resumo?plugin_pendente={i}"


def ler_argumento(argv) -> str:
    """'--abrir-obra <arquivo>' na linha de comando do .exe (aberto pelo plugin)."""
    for k, a in enumerate(argv):
        if a == "--abrir-obra" and k + 1 < len(argv):
            return argv[k + 1]
        if a.startswith("--abrir-obra="):
            return a.split("=", 1)[1]
    return ""


def carregar_arquivo(caminho: str) -> str:
    """Lê o arquivo que o plugin gravou ({chave, envelope}), guarda a obra e apaga o arquivo. Devolve o id ('' se não deu)."""
    try:
        with open(caminho, encoding="utf-8") as f:
            d = json.load(f)
        erro = validar(d.get("chave"), d.get("envelope"))
        if erro:
            logger.warning(f"Obra do nanoCAD ignorada ({caminho}): {erro}")
            return ""
        return guardar(d["chave"], d["envelope"])
    except (OSError, ValueError) as e:
        logger.warning(f"Obra do nanoCAD ilegível ({caminho}): {e}")
        return ""
    finally:
        try:
            os.remove(caminho)
        except OSError:
            pass
