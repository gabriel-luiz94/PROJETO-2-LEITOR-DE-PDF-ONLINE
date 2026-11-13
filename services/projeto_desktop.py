"""
services/projeto_desktop.py — Projeto de trabalho (`.obra` numa pasta) na janela desktop (pywebview), TASK-154.

`ApiProjeto` estende a ponte de arquivos: o JavaScript chama `pywebview.api.projeto_*`. Toda a lógica está em funções de módulo
(`criar`, `gravar`, `abrir`, `assinatura_atual`) para testar sem janela. Nenhuma função levanta exceção: devolve {ok: False, erro}.
Escrita atômica (arquivo temporário + os.replace) e trava por caminho; só grava arquivos `.obra`.
"""
import json
import os
import threading
from datetime import datetime

import config
from config import logger
from services import projeto_obra
from services.arquivos_desktop import ApiArquivos, _escolhido

_travas: dict = {}
_ultimas: dict = {}                    # caminho → assinatura da última gravação FEITA POR ESTE PROCESSO (a tela pode perder a resposta ao fechar no meio de uma gravação)
_travas_mestre = threading.Lock()


def _trava(caminho: str) -> threading.Lock:
    with _travas_mestre:
        return _travas.setdefault(os.path.normcase(os.path.abspath(caminho)), threading.Lock())


def _chave(caminho: str) -> str:
    return os.path.normcase(os.path.abspath(caminho))


def _caminho_valido(caminho) -> str:
    if not isinstance(caminho, str) or not caminho.lower().endswith(projeto_obra.EXTENSAO):
        raise projeto_obra.ErroProjeto(["O projeto precisa ser um arquivo .obra."])
    return os.path.abspath(caminho)


def _carregar_json(caminho: str):
    with open(caminho, "rb") as f:
        bruto = f.read(projeto_obra.LIMITE_BYTES + 1)
    if len(bruto) > projeto_obra.LIMITE_BYTES:
        raise projeto_obra.ErroProjeto([f"Arquivo grande demais (máx. {projeto_obra.LIMITE_BYTES // (1024 * 1024)} MB)."])
    return bruto


def _gravar_atomico(caminho: str, texto: str) -> str:
    dados = texto.encode("utf-8")
    temporario = caminho + ".tmp"
    with open(temporario, "wb") as f:
        f.write(dados)
        f.flush()
        os.fsync(f.fileno())
    os.replace(temporario, caminho)
    return projeto_obra.assinatura(dados)


def _modificado_em(caminho: str) -> str:
    try:
        return datetime.fromtimestamp(os.path.getmtime(caminho)).isoformat(timespec="seconds")
    except OSError:
        return ""


def _falha(e) -> dict:
    if isinstance(e, projeto_obra.ErroProjeto):
        return {"ok": False, "erro": " ".join(e.mensagens)}
    logger.warning(f"projeto: operação falhou ({type(e).__name__})")          # sem o conteúdo do projeto no log
    return {"ok": False, "erro": f"Não foi possível acessar o arquivo ({type(e).__name__})."}


def criar(pasta: str, parcial: dict, projeto_codigo: str = "") -> dict:
    """Cria a pasta (se faltar) e `<nome>.obra` nela. Recusa se o arquivo já existe."""
    try:
        if not isinstance(pasta, str) or not pasta.strip():
            raise projeto_obra.ErroProjeto(["Escolha a pasta do projeto."])
        texto = projeto_obra.montar(parcial, None, projeto_codigo)
        pasta = os.path.abspath(pasta.strip())
        os.makedirs(pasta, exist_ok=True)
        caminho = os.path.join(pasta, projeto_obra.nome_arquivo(parcial.get("nome")))
        with _trava(caminho):
            if os.path.exists(caminho):
                return {"ok": False, "existe": True, "caminho": caminho, "erro": "Já existe um projeto com esse nome nessa pasta."}
            assin = _gravar_atomico(caminho, texto)
            _ultimas[_chave(caminho)] = assin
            return {"ok": True, "caminho": caminho, "assinatura": assin}
    except Exception as e:  # noqa: BLE001 — a ponte nunca derruba a janela
        return _falha(e)


def gravar(caminho: str, parcial: dict, esperada: str = "", forcar: bool = False, projeto_codigo: str = "") -> dict:
    """Atualiza o `.obra` com o que mudou (o resto é mantido). Se o arquivo foi alterado por fora (assinatura diferente da esperada),
    devolve {conflito: True} sem gravar — a tela decide; `forcar` grava mesmo assim."""
    try:
        caminho = _caminho_valido(caminho)
        with _trava(caminho):
            existente, atual = None, ""
            if os.path.exists(caminho):
                bruto = _carregar_json(caminho)
                atual = projeto_obra.assinatura(bruto)
                if esperada and atual != esperada and atual != _ultimas.get(_chave(caminho)) and not forcar:       # igual à última gravação deste processo = não houve mudança por fora
                    return {"ok": False, "conflito": True, "assinatura_atual": atual, "modificado_em": _modificado_em(caminho),
                            "erro": "O arquivo do projeto foi alterado por outro programa ou computador."}
                try:
                    existente = json.loads(bruto.decode("utf-8-sig"))
                except (ValueError, UnicodeDecodeError):
                    existente = None
            elif not os.path.isdir(os.path.dirname(caminho)):
                raise projeto_obra.ErroProjeto(["A pasta do projeto não existe mais."])
            texto = projeto_obra.montar(parcial, existente, projeto_codigo)
            assin = _gravar_atomico(caminho, texto)
            _ultimas[_chave(caminho)] = assin
            return {"ok": True, "caminho": caminho, "assinatura": assin}
    except Exception as e:  # noqa: BLE001
        return _falha(e)


def abrir(caminho: str, projeto_selecionado: str) -> dict:
    """Lê e valida o `.obra` (ou `.obra.json` antigo): {ok, caminho, assinatura, projeto: {nome, numero_obra, sap, dados, resultado, legado}}."""
    try:
        if not isinstance(caminho, str) or not os.path.isfile(caminho):
            raise projeto_obra.ErroProjeto(["O arquivo do projeto não foi encontrado."])
        bruto = _carregar_json(caminho)
        return {"ok": True, "caminho": os.path.abspath(caminho), "assinatura": projeto_obra.assinatura(bruto),
                "projeto": projeto_obra.ler(bruto, projeto_selecionado)}
    except Exception as e:  # noqa: BLE001
        return _falha(e)


def assinatura_atual(caminho: str) -> dict:
    try:
        caminho = _caminho_valido(caminho)
        if not os.path.isfile(caminho):
            return {"ok": True, "existe": False}
        return {"ok": True, "existe": True, "assinatura": projeto_obra.assinatura(_carregar_json(caminho))}
    except Exception as e:  # noqa: BLE001
        return _falha(e)


# ───────── rascunho do projeto default (TASK-158) ─────────
# No projeto default o salvamento automático grava um rascunho numa pasta fixa junto do aplicativo, para o usuário recuperar depois.
# `atual` = o que está sendo feito; `anterior` = o último estado da sessão anterior (guardado quando uma nova sessão começa a gravar).
RASCUNHOS = {"atual": "rascunho_default.obra", "anterior": "rascunho_default_anterior.obra"}


def pasta_rascunho() -> str:
    return os.path.join(config.BASE_DIR, "rascunhos")


def rascunho_gravar(parcial: dict, projeto_codigo: str = "", nova_sessao: bool = False, pasta: str = "") -> dict:
    """Grava o rascunho do default (só Resumo). `nova_sessao`: guarda antes o rascunho atual como "anterior"."""
    try:
        if not str(projeto_codigo or "").strip():
            raise projeto_obra.ErroProjeto(["Projeto de regras não informado."])                # sem ele o arquivo não poderia ser aberto depois
        pasta = pasta or pasta_rascunho()
        os.makedirs(pasta, exist_ok=True)
        atual, anterior = os.path.join(pasta, RASCUNHOS["atual"]), os.path.join(pasta, RASCUNHOS["anterior"])
        texto = projeto_obra.montar({"nome": "Rascunho (projeto default)", "dados": (parcial or {}).get("dados") or {}}, None, projeto_codigo)
        with _trava(atual):
            if nova_sessao and os.path.isfile(atual):
                os.replace(atual, anterior)
            return {"ok": True, "caminho": atual, "assinatura": _gravar_atomico(atual, texto)}
    except Exception as e:  # noqa: BLE001
        return _falha(e)


def rascunho_info(pasta: str = "") -> dict:
    """Rascunhos que existem: [{id, salvo_em, itens}] (mais recente primeiro)."""
    try:
        pasta = pasta or pasta_rascunho()
        saida = []
        for id_, nome in RASCUNHOS.items():
            caminho = os.path.join(pasta, nome)
            if not os.path.isfile(caminho):
                continue
            try:
                carga = json.loads(_carregar_json(caminho).decode("utf-8-sig"))
                dados = carga.get("dados") or {}
                itens = sum(len(dados.get(k) or []) for k in ("cabos", "outros"))
                saida.append({"id": id_, "salvo_em": (carga.get("projeto_trabalho") or {}).get("salvo_em", ""), "itens": itens})
            except (ValueError, OSError, UnicodeDecodeError):
                continue
        return {"ok": True, "rascunhos": saida}
    except Exception as e:  # noqa: BLE001
        return _falha(e)


def rascunho_ler(id_: str, projeto_selecionado: str, pasta: str = "") -> dict:
    nome = RASCUNHOS.get(id_)
    if not nome:
        return {"ok": False, "erro": "Rascunho desconhecido."}
    return abrir(os.path.join(pasta or pasta_rascunho(), nome), projeto_selecionado)


def _parcial(meta_json) -> dict:
    try:
        meta = json.loads(meta_json) if isinstance(meta_json, str) else meta_json
    except ValueError:
        raise projeto_obra.ErroProjeto(["Dados do projeto inválidos."])
    if not isinstance(meta, dict):
        raise projeto_obra.ErroProjeto(["Dados do projeto inválidos."])
    return meta


class ApiProjeto(ApiArquivos):
    """js_api da janela: arquivos (TASK-142) + projeto de trabalho. Só métodos públicos aparecem para o JavaScript."""

    def projeto_escolher_pasta(self, inicial=""):
        try:
            if self._janela is None:
                return {"ok": False, "erro": "janela indisponível"}
            r = _escolhido(self._janela.create_file_dialog(self._dialogo("FOLDER"), directory=inicial if isinstance(inicial, str) and os.path.isdir(inicial) else ""))
            return {"ok": True, "pasta": r[0]} if r else {"cancelado": True}
        except Exception as e:  # noqa: BLE001
            return _falha(e)

    def projeto_criar(self, pasta, meta_json, projeto_codigo=""):
        try:
            return criar(pasta, _parcial(meta_json), projeto_codigo)
        except Exception as e:  # noqa: BLE001
            return _falha(e)

    def projeto_escolher_arquivo(self, projeto_selecionado=""):
        """Caixa "Abrir" (.obra e .obra.json antigo) e já lê o escolhido."""
        try:
            if self._janela is None:
                return {"ok": False, "erro": "janela indisponível"}
            r = _escolhido(self._janela.create_file_dialog(self._dialogo("OPEN"), allow_multiple=False,
                                                           file_types=("Projeto OrçaCore (*.obra;*.json)", "Todos os arquivos (*.*)")))
            return abrir(r[0], projeto_selecionado) if r else {"cancelado": True}
        except Exception as e:  # noqa: BLE001
            return _falha(e)

    def projeto_abrir(self, caminho, projeto_selecionado=""):
        return abrir(caminho, projeto_selecionado)

    def projeto_gravar(self, caminho, meta_json, assinatura_esperada="", forcar=False, projeto_codigo=""):
        try:
            return gravar(caminho, _parcial(meta_json), assinatura_esperada, bool(forcar), projeto_codigo)
        except Exception as e:  # noqa: BLE001
            return _falha(e)

    def projeto_assinatura(self, caminho):
        return assinatura_atual(caminho)

    def projeto_rascunho_gravar(self, meta_json, projeto_codigo="", nova_sessao=False):
        try:
            return rascunho_gravar(_parcial(meta_json), projeto_codigo, bool(nova_sessao))
        except Exception as e:  # noqa: BLE001
            return _falha(e)

    def projeto_rascunho_info(self):
        return rascunho_info()

    def projeto_rascunho_ler(self, id_rascunho, projeto_selecionado=""):
        return rascunho_ler(id_rascunho, projeto_selecionado)
