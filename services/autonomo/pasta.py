"""
services/autonomo/pasta.py — Pasta monitorada e fila do modo autônomo (TASK-031, fase C).

`entrada/<código-do-projeto>/arquivo.dxf|pdf` → pipeline (um arquivo por vez) → original movido para `processados/<projeto>/` (ou `erros/<projeto>/`
com um `.erro.txt`); resultados em `saida/` e obra no banco. Regras de segurança:
  • só processa arquivo ESTÁVEL (mesmo tamanho e data em duas varreduras separadas por `estabilizacao_s`) e que dê para abrir — evita pegar cópia em andamento;
  • a falha de um arquivo nunca derruba a fila nem apaga o original; erro vira `erros/` + `.erro.txt` + linha no histórico;
  • pasta de projeto inexistente, arquivo solto na raiz e extensão não suportada vão para `erros/` com o motivo (nada fica parado para sempre);
  • o mesmo conteúdo no mesmo projeto não é processado duas vezes (vai para `processados/<projeto>/duplicados/`);
  • `TRAVA` serializa o pipeline (vigia + ações da API): um arquivo por vez.
Limitação conhecida: o tempo de cada arquivo é limitado pelo leitor (60 s por chamada) e pelo tamanho (200 MB); não há cancelamento do pipeline inteiro.
"""
import logging
import os
import shutil
import threading
import time
from datetime import datetime

import database
from services.autonomo import config_autonomo, execucoes, pipeline, saida

log = logging.getLogger(__name__)
TRAVA = threading.RLock()
IGNORAR_PREFIXOS = (".", "~$")
IGNORAR_SUFIXOS = (".tmp", ".part", ".partial", ".crdownload", ".download", ".lock", ".erro.txt")


def _ignorado(nome: str) -> bool:
    n = nome.lower()
    return n.startswith(IGNORAR_PREFIXOS) or n.endswith(IGNORAR_SUFIXOS)


def _mover(origem: str, pasta_destino: str) -> str:
    """Move sem sobrescrever (sufixo de data se o nome já existe). PermissionError = arquivo EM USO (no Windows, outro programa com o arquivo
    aberto): quem chama deixa o arquivo onde está e tenta de novo na próxima varredura."""
    os.makedirs(pasta_destino, exist_ok=True)
    nome = os.path.basename(origem)
    destino = os.path.join(pasta_destino, nome)
    if os.path.exists(destino):
        stem, ext = os.path.splitext(nome)
        destino = os.path.join(pasta_destino, f"{stem}-{datetime.now().strftime('%Y%m%d-%H%M%S-%f')}{ext}")
    try:
        os.rename(origem, destino)                      # mesmo volume: atômico; falha com PermissionError se o arquivo está em uso
    except PermissionError:
        raise
    except OSError:                                     # volumes diferentes: copia e depois apaga a origem
        shutil.copy2(origem, destino)
        try:
            os.remove(origem)
        except PermissionError:
            os.remove(destino)                          # não deixa cópia órfã: o arquivo continua na entrada e será tentado de novo
            raise
    return destino


def _projeto_existe(codigo: str) -> bool:
    conn = database.get_connection()
    try:
        return conn.execute("SELECT 1 FROM projetos WHERE codigo = ?", (codigo,)).fetchone() is not None
    finally:
        conn.close()


def _abre(caminho: str) -> bool:
    try:
        with open(caminho, "rb") as f:
            f.read(1)
        return True
    except OSError:
        return False


class Vigia:
    def __init__(self, relogio=time.time):
        self._relogio = relogio
        self._vistos = {}                 # caminho -> (tamanho, mtime_ns, instante da 1ª observação)
        self._adiados = {}                # caminho -> movimento pendente (arquivo em uso): {pasta, exec_id, erro, atualizar}
        self._parar = threading.Event()
        self._thread = None
        self._estado = {"ultima_varredura": None, "processando": None, "na_fila": 0, "processados": 0, "erros": 0, "ultimo_erro": None}
        self._estado_trava = threading.Lock()
        self._varredura = threading.RLock()   # uma varredura por vez (a thread da vigia e `POST /varrer` não se atropelam)

    # ── estado ──────────────────────────────────────────────────────────────
    def _set(self, **kw):
        with self._estado_trava:
            self._estado.update(kw)

    def status(self) -> dict:
        with self._estado_trava:
            return {**self._estado, "ligado": self.ativo()}

    def ativo(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    # ── ciclo de vida ───────────────────────────────────────────────────────
    def iniciar(self) -> None:
        if self.ativo():
            return
        self._parar.clear()
        self._thread = threading.Thread(target=self._laco, name="autonomo-vigia", daemon=True)
        self._thread.start()

    def parar(self, espera: float = 10.0) -> None:
        self._parar.set()
        t = self._thread
        if t is not None and t is not threading.current_thread():
            t.join(espera)
        if t is None or not t.is_alive():
            self._thread = None

    def _laco(self):
        while not self._parar.is_set():
            cfg = config_autonomo.carregar()
            try:
                self.varrer(cfg)
            except Exception:  # noqa: BLE001 — a vigia nunca morre por um erro de varredura
                log.exception("Falha na varredura do modo autônomo")
                self._set(ultimo_erro="Falha inesperada na varredura (veja o log).")
            self._parar.wait(max(1, int(cfg.get("intervalo_s") or 5)))

    # ── uma varredura ───────────────────────────────────────────────────────
    def _candidatos(self, entrada: str):
        """[(caminho, projeto|None)] de arquivos no 1º nível de cada subpasta (projeto) e soltos na raiz (projeto None)."""
        saida_ = []
        try:
            itens = sorted(os.scandir(entrada), key=lambda e: e.name)
        except FileNotFoundError:
            return saida_
        for e in itens:
            if _ignorado(e.name):
                continue
            if e.is_file():
                saida_.append((e.path, None))
            elif e.is_dir():
                try:
                    for f in sorted(os.scandir(e.path), key=lambda x: x.name):
                        if f.is_file() and not _ignorado(f.name):
                            saida_.append((f.path, e.name))
                except OSError:
                    continue
        return saida_

    def _pronto(self, caminho: str, estabilizacao: float, agora: float) -> bool:
        try:
            st = os.stat(caminho)
        except OSError:
            self._vistos.pop(caminho, None)
            return False
        atual = (st.st_size, st.st_mtime_ns)
        visto = self._vistos.get(caminho)
        if visto is None or visto[:2] != atual:
            self._vistos[caminho] = (*atual, agora)       # primeira vez (ou ainda mudando): reinicia a contagem
            return False
        return agora - visto[2] >= estabilizacao and _abre(caminho)

    def varrer(self, cfg: dict = None) -> list:
        """Uma passada: processa os arquivos prontos (um por vez). Devolve [{arquivo, projeto, status, execucao}]."""
        with self._varredura:
            return self._varrer(cfg)

    def _varrer(self, cfg: dict = None) -> list:
        cfg = cfg or config_autonomo.carregar()
        p = config_autonomo.garantir_pastas(cfg)
        agora = self._relogio()
        self._retomar_movimentos()
        candidatos = [(c, proj) for c, proj in self._candidatos(p["entrada"]) if c not in self._adiados]
        existentes = {c for c, _ in candidatos}
        for k in [k for k in self._vistos if k not in existentes]:
            del self._vistos[k]
        prontos = [(c, proj) for c, proj in candidatos if self._pronto(c, cfg["estabilizacao_s"], agora)]
        self._set(ultima_varredura=datetime.now().isoformat(timespec="seconds"), na_fila=len(candidatos))
        feitos = []
        for caminho, projeto in prontos:
            if self._parar.is_set():
                break
            feitos.append(self._tratar(caminho, projeto, cfg, p))
            self._vistos.pop(caminho, None)
        self._set(processando=None)
        return feitos

    # ── movimentos (um arquivo em uso não pode ser movido: espera, sem erro e sem reprocessar) ──
    def _concluir_movimento(self, m: dict, destino: str) -> None:
        if m.get("erro"):
            try:
                with open(destino + ".erro.txt", "w", encoding="utf-8") as f:
                    f.write(f"{datetime.now().isoformat(timespec='seconds')}\n{m['erro']}\n")
            except OSError:
                pass
        if m.get("exec_id") and m.get("atualizar", True):
            execucoes.atualizar(m["exec_id"], arquivo_caminho=destino)

    def _mover_ou_adiar(self, caminho: str, pasta: str, exec_id: str = None, erro: str = None, atualizar: bool = True):
        """Devolve o destino, ou None se o arquivo está em uso (fica na pasta e o movimento é repetido a cada varredura)."""
        m = {"pasta": pasta, "exec_id": exec_id, "erro": erro, "atualizar": atualizar}
        try:
            destino = _mover(caminho, pasta)
        except PermissionError:
            self._adiados[caminho] = m
            return None
        self._concluir_movimento(m, destino)
        return destino

    def _retomar_movimentos(self) -> None:
        for caminho, m in list(self._adiados.items()):
            if not os.path.exists(caminho):
                del self._adiados[caminho]
                continue
            try:
                destino = _mover(caminho, m["pasta"])
            except PermissionError:
                continue                                  # ainda em uso
            except OSError:
                log.exception("Falha ao mover %s", caminho)
                continue
            del self._adiados[caminho]
            self._concluir_movimento(m, destino)

    def _para_erros(self, caminho: str, projeto, motivo: str, p: dict, exec_id: str = None, cfg_user: str = None) -> dict:
        if not exec_id:      # falha ANTES do pipeline (projeto inexistente, tipo inválido…): também aparece no histórico
            try:
                try:
                    h = pipeline.hash_arquivo(caminho)
                except OSError:
                    h = ""
                exec_id = execucoes.criar(os.path.basename(caminho), h, projeto or "", cfg_user or "", status="erro")
                execucoes.atualizar(exec_id, mensagem=motivo, relatorio={"etapas": [], "erro": motivo})
            except Exception:  # noqa: BLE001 — o histórico é complemento; o arquivo vai para erros/ do mesmo jeito
                log.exception("Falha ao registrar o erro no histórico")
                exec_id = None
        self._mover_ou_adiar(caminho, os.path.join(p["erros"], saida.nome_seguro(projeto, "sem_projeto")), exec_id=exec_id, erro=motivo)
        with self._estado_trava:
            self._estado["erros"] += 1
            self._estado["ultimo_erro"] = f"{os.path.basename(caminho)}: {motivo}"
        return {"arquivo": os.path.basename(caminho), "projeto": projeto, "status": "erro", "execucao": exec_id, "mensagem": motivo}

    def _tratar(self, caminho: str, projeto, cfg: dict, p: dict) -> dict:
        nome = os.path.basename(caminho)
        try:
            if projeto is None:
                return self._para_erros(caminho, None, "Arquivo solto na pasta de entrada: coloque-o em uma subpasta com o código do projeto.", p, cfg_user=cfg.get("user_id"))
            if not nome.lower().endswith(pipeline.EXTENSOES):
                return self._para_erros(caminho, projeto, f"Tipo de arquivo não suportado: '{os.path.splitext(nome)[1] or '(sem extensão)'}' (use {', '.join(pipeline.EXTENSOES)}). Arquivo: {nome}", p, cfg_user=cfg.get("user_id"))
            if not _projeto_existe(projeto):
                return self._para_erros(caminho, projeto, f"Projeto '{projeto}' não está cadastrado (a subpasta deve ter o código do projeto).", p, cfg_user=cfg.get("user_id"))
            if not cfg.get("user_id"):
                return {"arquivo": nome, "projeto": projeto, "status": "parado", "execucao": None,
                        "mensagem": "Escolha o usuário dono das obras na configuração do modo autônomo."}   # fica na pasta; não é erro do arquivo
            self._set(processando=nome)
            with TRAVA:
                ex = pipeline.processar_arquivo(caminho, projeto, cfg["user_id"], p["saida"])
            if ex.get("duplicado"):
                self._mover_ou_adiar(caminho, os.path.join(p["processados"], saida.nome_seguro(projeto), "duplicados"), exec_id=ex["id"], atualizar=False)
                return {"arquivo": nome, "projeto": projeto, "status": "duplicado", "execucao": ex["id"]}
            if ex["status"] == "erro":
                return self._para_erros(caminho, projeto, ex["mensagem"] or "Erro ao processar.", p, ex["id"])
            self._mover_ou_adiar(caminho, os.path.join(p["processados"], saida.nome_seguro(projeto)), exec_id=ex["id"])
            with self._estado_trava:
                self._estado["processados"] += 1
            return {"arquivo": nome, "projeto": projeto, "status": ex["status"], "execucao": ex["id"]}
        except Exception as e:  # noqa: BLE001 — nada derruba a fila
            log.exception("Falha ao tratar %s", nome)
            try:
                return self._para_erros(caminho, projeto, f"Erro inesperado: {type(e).__name__}", p, cfg_user=cfg.get("user_id"))
            except Exception:  # noqa: BLE001
                return {"arquivo": nome, "projeto": projeto, "status": "erro", "execucao": None, "mensagem": "Falha ao mover o arquivo."}


def reprocessar(exec_id: str) -> dict:
    """Roda de novo o arquivo de uma execução anterior (que está em `processados/`/`erros/`), criando uma NOVA execução."""
    ex = execucoes.buscar(exec_id)
    if not ex:
        raise pipeline.ErroPipeline("Execução não encontrada.")
    caminho = ex.get("arquivo_caminho")
    if not caminho or not os.path.isfile(caminho):
        raise pipeline.ErroPipeline("O arquivo original não está mais disponível para reprocessar.")
    cfg = config_autonomo.carregar()
    p = config_autonomo.garantir_pastas(cfg)
    with TRAVA:
        novo = pipeline.processar_arquivo(caminho, ex["projeto_codigo"], cfg.get("user_id") or ex["user_id"], p["saida"], forcar=True)
    return novo


_vigia = None
_trava_vigia = threading.Lock()


def obter_vigia() -> Vigia:
    global _vigia
    with _trava_vigia:
        if _vigia is None:
            _vigia = Vigia()
        return _vigia
