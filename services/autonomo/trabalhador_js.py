"""
services/autonomo/trabalhador_js.py — Processo trabalhador do leitor (QuickJS). Não importar; é executado como processo filho por
services/autonomo/leitor_js.py (`python -m services.autonomo.trabalhador_js`).

Protocolo: uma linha JSON por mensagem em stdin/stdout. 1ª mensagem recebida: {"fonte": "<js do motor>"}; depois {"fn": "lote"|"classificar",
"carga": "<json>"}; resposta {"ok": true, "r": "<json>"} ou {"ok": false, "erro": "..."}; EOF encerra. stdout é só do protocolo.
"""
import json
import sys

LIMITE_MEMORIA = 256 * 1024 * 1024

COLA = """
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


def _enviar(obj):
    sys.__stdout__.write(json.dumps(obj, ensure_ascii=True) + "\n")
    sys.__stdout__.flush()


def main():
    stdout_real = sys.stdout
    sys.stdout = sys.stderr          # qualquer print perdido não corrompe o protocolo
    try:
        import quickjs
    except ImportError:
        _enviar({"ok": False, "erro": "dependência 'quickjs' não instalada (pip install quickjs)"})
        return
    try:
        primeira = json.loads(sys.stdin.readline())
        ctx = quickjs.Context()
        ctx.set_memory_limit(LIMITE_MEMORIA)
        ctx.eval("var module = {exports: {}};")  # o arquivo exporta via module.exports (modo Node)
        ctx.eval(primeira["fonte"])
        ctx.eval(COLA)
        funcoes = {"lote": ctx.eval("__lote"), "classificar": ctx.eval("__classificar")}
    except Exception as e:  # noqa: BLE001
        _enviar({"ok": False, "erro": f"Falha ao carregar o motor do leitor: {e}"})
        return
    _enviar({"ok": True})
    for linha in sys.stdin:
        if not linha.strip():
            continue
        try:
            pedido = json.loads(linha)
            _enviar({"ok": True, "r": str(funcoes[pedido["fn"]](pedido["carga"]))})
        except Exception as e:  # noqa: BLE001
            _enviar({"ok": False, "erro": f"Erro no motor do leitor: {e}"})
    sys.stdout = stdout_real


if __name__ == "__main__":
    main()
