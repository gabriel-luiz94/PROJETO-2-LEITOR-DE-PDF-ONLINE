"""
services/autonomo/leitor_js.py — Executa o MESMO motor do leitor da tela (static/regras_leitor_engine.js) no backend (TASK-031, fase A).

O arquivo JS é a fonte única da verdade: nada dele é copiado para Python. Ele roda num motor QuickJS embutido, sem acesso a
arquivos nem rede, **dentro de um processo separado** (trabalhador): o limite de tempo do QuickJS NÃO interrompe uma regex com
backtracking catastrófico (testado), então quem controla o tempo é o processo principal, que mata e recria o trabalhador se uma
regra cadastrada travar. Assim uma regra ruim derruba só aquele arquivo, nunca o programa.

`quickjs` só é importado dentro do trabalhador — quem nunca usa o modo autônomo não carrega a dependência.
NOTA (build Windows/PyInstaller): o trabalhador usa `multiprocessing` (spawn); ao ligar o modo autônomo no executável, `app.py`
precisa chamar `multiprocessing.freeze_support()` no início do bloco `__main__` (fases C/D).

Itens de entrada = os da extração: {pagina, texto, cor, layer}. A "vizinhança" das regras enxerga a lista inteira, como na tela.
"""
import json
import multiprocessing
import os
import threading

import config

CAMINHO_MOTOR = os.path.join(config.STATIC_DIR, "regras_leitor_engine.js")
LIMITE_MEMORIA = 256 * 1024 * 1024
LIMITE_SEGUNDOS = 60   # tempo máximo de UMA chamada; lido a cada chamada (os testes reduzem)

# Cola em JS: mesma montagem de `engineItem` de script.js (updateRowLogic) e resumo.js (computeRowLogic).
_COLA = """
var Motor = module.exports;
function __lote(json) {
    var a = JSON.parse(json), itens = a.itens, idx = a.indices, P = a.proc, C = a.cls, resumo = a.modo === 'resumo';
    var alvo = idx || itens.map(function (_, i) { return i; });
    return JSON.stringify(alvo.map(function (i) {
        var it = itens[i];
        return Motor.processarEClassificar({
            texto: resumo ? (it.texto || '') : it.texto, cor: it.cor || '#000000', layer: it.layer || '',
            pagina: it.pagina, index: i, allItems: itens
        }, P, C);
    }));
}
function __classificar(json) {
    var a = JSON.parse(json);
    return JSON.stringify(Motor.classificar(a.operacao, a.ativo, a.cor, a.layer, a.texto, a.cls));
}
"""


class ErroLeitorJS(RuntimeError):
    pass


def _trabalhador(conn, fonte: str):
    """Corre no processo filho: carrega o motor e atende pedidos (nome_da_funcao, json) até receber None."""
    try:
        import quickjs
    except ImportError:
        conn.send(("erro", "dependência 'quickjs' não instalada (pip install quickjs)"))
        return
    try:
        ctx = quickjs.Context()
        ctx.set_memory_limit(LIMITE_MEMORIA)
        ctx.eval("var module = {exports: {}};")  # o arquivo exporta via module.exports (modo Node)
        ctx.eval(fonte)
        ctx.eval(_COLA)
        funcoes = {"lote": ctx.eval("__lote"), "classificar": ctx.eval("__classificar")}
    except Exception as e:  # noqa: BLE001 — qualquer falha de carga vira erro claro no processo principal
        conn.send(("erro", f"Falha ao carregar o motor do leitor: {e}"))
        return
    conn.send(("ok", None))
    while True:
        pedido = conn.recv()
        if pedido is None:
            return
        nome, carga = pedido
        try:
            conn.send(("ok", str(funcoes[nome](carga))))
        except Exception as e:  # noqa: BLE001
            conn.send(("erro", f"Erro no motor do leitor: {e}"))


class LeitorJS:
    def __init__(self, caminho: str = None):
        with open(caminho or CAMINHO_MOTOR, "r", encoding="utf-8") as f:
            self._fonte = f.read()
        self._trava = threading.Lock()
        self._proc = None
        self._conn = None
        self._iniciar()

    def _iniciar(self):
        ctx = multiprocessing.get_context("spawn")
        pai, filho = ctx.Pipe()
        proc = ctx.Process(target=_trabalhador, args=(filho, self._fonte), daemon=True)
        proc.start()
        filho.close()
        self._proc, self._conn = proc, pai
        estado, msg = self._receber(LIMITE_SEGUNDOS)
        if estado != "ok":
            self.fechar()
            raise ErroLeitorJS(msg)

    def _receber(self, limite):
        if not self._conn.poll(limite):
            self.fechar()
            raise ErroLeitorJS(f"O motor do leitor passou de {limite}s (regra cadastrada muito lenta?). O trabalhador foi reiniciado.")
        try:
            return self._conn.recv()
        except (EOFError, OSError):
            self.fechar()
            raise ErroLeitorJS("O processo do motor do leitor terminou inesperadamente.")

    def fechar(self):
        try:
            if self._proc is not None and self._proc.is_alive():
                self._proc.terminate()
                self._proc.join(5)
        finally:
            if self._conn is not None:
                self._conn.close()
            self._proc = self._conn = None

    def _chamar(self, nome, carga):
        with self._trava:
            if self._proc is None or not self._proc.is_alive():
                self._iniciar()
            self._conn.send((nome, json.dumps(carga, ensure_ascii=False)))
            estado, msg = self._receber(LIMITE_SEGUNDOS)
            if estado != "ok":
                raise ErroLeitorJS(msg)
            return json.loads(msg)

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
