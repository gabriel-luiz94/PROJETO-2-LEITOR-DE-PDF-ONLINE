"""
services/autonomo/trabalhador_js.py — Processo trabalhador do leitor (QuickJS). Não importar; é executado como processo filho por
services/autonomo/leitor_js.py (`python -m services.autonomo.trabalhador_js`).

Protocolo: uma linha JSON por mensagem num SOCKET local (127.0.0.1:<porta>; argumentos `<porta> <token>`). O trabalhador envia {"token"};
recebe {"fonte": "<js do motor>"}; depois {"fn": "lote"|"classificar", "carga": "<json>"}; responde {"ok": true, "r": "<json>"} ou
{"ok": false, "erro": "..."}; EOF encerra. (Socket e não stdin/stdout: no .exe `--windowed` o stdin/stdout do Python não existem.)
"""
import json
import socket
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


def _servir(porta: int, token: str):
    """Conecta no processo principal (127.0.0.1) e atende pedidos. Usa SOCKET e não stdin/stdout: no .exe do PyInstaller sem console
    (`--windowed`) o stdin/stdout do Python não existem, mas um socket funciona igual em qualquer sistema."""
    sock = socket.create_connection(("127.0.0.1", porta), timeout=30)
    sock.settimeout(None)
    leitura = sock.makefile("r", encoding="utf-8", newline="\n")
    escrita = sock.makefile("w", encoding="utf-8", newline="\n")

    def enviar(obj):
        escrita.write(json.dumps(obj, ensure_ascii=True) + "\n")
        escrita.flush()

    enviar({"token": token})                       # o principal só aceita quem sabe o token
    try:
        import quickjs
    except ImportError:
        enviar({"ok": False, "erro": "dependência 'quickjs' não instalada (pip install quickjs)"})
        return
    try:
        primeira = json.loads(leitura.readline())
        ctx = quickjs.Context()
        ctx.set_memory_limit(LIMITE_MEMORIA)
        ctx.eval("var module = {exports: {}};")  # o arquivo exporta via module.exports (modo Node)
        ctx.eval(primeira["fonte"])
        ctx.eval(COLA)
        funcoes = {"lote": ctx.eval("__lote"), "classificar": ctx.eval("__classificar")}
    except Exception as e:  # noqa: BLE001
        enviar({"ok": False, "erro": f"Falha ao carregar o motor do leitor: {e}"})
        return
    enviar({"ok": True})
    for linha in leitura:
        if not linha.strip():
            continue
        try:
            pedido = json.loads(linha)
            enviar({"ok": True, "r": str(funcoes[pedido["fn"]](pedido["carga"]))})
        except Exception as e:  # noqa: BLE001
            enviar({"ok": False, "erro": f"Erro no motor do leitor: {e}"})


def main(argv=None):
    """`<porta> <token>` (depois de `--trabalhador-leitor-js` quando chamado pelo app.py/.exe, ou direto com `-m`)."""
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--trabalhador-leitor-js" in argv:
        argv = argv[argv.index("--trabalhador-leitor-js") + 1:]
    _servir(int(argv[0]), argv[1])


if __name__ == "__main__":
    main()
