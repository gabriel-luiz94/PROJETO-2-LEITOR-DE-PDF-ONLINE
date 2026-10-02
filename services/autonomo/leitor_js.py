"""
services/autonomo/leitor_js.py — Executa o MESMO motor do leitor da tela (static/regras_leitor_engine.js) no backend (TASK-031, fase A).

O arquivo JS é a fonte única da verdade: nada dele é copiado para Python. Ele roda num motor QuickJS embutido, sem acesso a
arquivos nem rede, **dentro de um processo separado** (trabalhador): o limite de tempo do QuickJS NÃO interrompe uma regex com
backtracking catastrófico (testado), então quem controla o tempo é o processo principal, que mata e recria o trabalhador se uma
regra cadastrada travar. Assim uma regra ruim derruba só aquele arquivo, nunca o programa.

`quickjs` só é importado dentro do trabalhador — quem nunca usa o modo autônomo não carrega a dependência.
O trabalhador é um subprocesso (`python -m services.autonomo.trabalhador_js`, protocolo de linhas JSON) e não `multiprocessing`: o
`spawn` reimportaria o módulo principal (`app.py`) dentro do filho. NOTA (executável PyInstaller): `sys.executable` é o próprio .exe,
então ao ligar o modo autônomo no executável `app.py` precisa tratar o argumento `--trabalhador-leitor-js` chamando
`services.autonomo.trabalhador_js.main()` e saindo, antes de abrir a janela (fases C/D; não feito aqui para não tocar no `app.py`).

Itens de entrada = os da extração: {pagina, texto, cor, layer}. A "vizinhança" das regras enxerga a lista inteira, como na tela.
"""
import json
import os
import queue
import subprocess
import sys
import threading

import config

CAMINHO_MOTOR = os.path.join(config.STATIC_DIR, "regras_leitor_engine.js")
LIMITE_MEMORIA = 256 * 1024 * 1024
LIMITE_SEGUNDOS = 60   # tempo máximo de UMA chamada; lido a cada chamada (os testes reduzem)

class ErroLeitorJS(RuntimeError):
    pass


_RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _comando_trabalhador() -> list:
    if getattr(sys, "frozen", False):
        return [sys.executable, "--trabalhador-leitor-js"]
    return [sys.executable, "-m", "services.autonomo.trabalhador_js"]


class LeitorJS:
    def __init__(self, caminho: str = None):
        with open(caminho or CAMINHO_MOTOR, "r", encoding="utf-8") as f:
            self._fonte = f.read()
        self._trava = threading.Lock()
        self._proc = None
        self._fila = None
        self._iniciar()

    def _iniciar(self):
        self._proc = subprocess.Popen(_comando_trabalhador(), stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                                      text=True, encoding="utf-8", cwd=_RAIZ, bufsize=1)
        fila = queue.Queue()
        self._fila = fila

        def ler(saida):          # thread leitora: permite esperar a resposta com limite de tempo
            try:
                for linha in saida:
                    fila.put(linha)
            finally:
                fila.put(None)   # EOF: o trabalhador morreu
        threading.Thread(target=ler, args=(self._proc.stdout,), daemon=True).start()
        self._enviar({"fonte": self._fonte})
        resp = self._receber(LIMITE_SEGUNDOS)
        if not resp.get("ok"):
            self.fechar()
            raise ErroLeitorJS(resp.get("erro", "Falha ao iniciar o motor do leitor."))

    def _enviar(self, obj):
        try:
            self._proc.stdin.write(json.dumps(obj, ensure_ascii=True) + "\n")
            self._proc.stdin.flush()
        except (BrokenPipeError, OSError, ValueError):
            self.fechar()
            raise ErroLeitorJS("O processo do motor do leitor terminou inesperadamente.")

    def _receber(self, limite) -> dict:
        try:
            linha = self._fila.get(timeout=limite)
        except queue.Empty:
            self.fechar()
            raise ErroLeitorJS(f"O motor do leitor passou de {limite}s (regra cadastrada muito lenta?). O trabalhador foi reiniciado.")
        if linha is None:
            self.fechar()
            raise ErroLeitorJS("O processo do motor do leitor terminou inesperadamente.")
        return json.loads(linha)

    def fechar(self):
        proc, self._proc = self._proc, None
        if proc is not None:
            try:
                proc.kill()
                proc.wait(5)
            except Exception:  # noqa: BLE001
                pass
            for f in (proc.stdin, proc.stdout):
                try:
                    f.close()
                except Exception:  # noqa: BLE001
                    pass

    def _chamar(self, nome, carga):
        with self._trava:
            if self._proc is None or self._proc.poll() is not None:
                self._iniciar()
            self._enviar({"fn": nome, "carga": json.dumps(carga, ensure_ascii=False)})
            resp = self._receber(LIMITE_SEGUNDOS)
            if not resp.get("ok"):
                raise ErroLeitorJS(resp.get("erro", "Erro no motor do leitor."))
            return json.loads(resp["r"])

    def processar_lote(self, itens: list, regras_proc: list, regras_cls: list, modo: str = "extracao", indices: list = None) -> list:
        """Roda processar+classificar em cada item (ou só em `indices`). modo 'extracao' = passo da tela do Leitor
        (script.js); 'resumo' = o passo do Resumo (resumo.js, `texto || ''`). Devolve [{operacao, ativo, entidade}]."""
        if modo not in ("extracao", "resumo"):
            raise ValueError("modo precisa ser 'extracao' ou 'resumo'")
        return self._chamar("lote", {"itens": itens, "indices": indices, "proc": regras_proc or [], "cls": regras_cls or [], "modo": modo})

    def classificar_entidade(self, ativo: str, regras_cls: list) -> str:
        """Equivalente de autoClassifyEntidade (resumo.js): só o texto do ativo, operação neutra 'I'."""
        if not ativo:
            return "0"
        r = self._chamar("classificar", {"operacao": "I", "ativo": ativo, "cor": "", "layer": "", "texto": ativo, "cls": regras_cls or []})
        return r["entidade"]


_leitor = None
_trava_global = threading.Lock()


def obter_leitor() -> LeitorJS:
    """Leitor compartilhado (criado sob demanda)."""
    global _leitor
    with _trava_global:
        if _leitor is None or _leitor._proc is None:
            _leitor = LeitorJS()
        return _leitor
