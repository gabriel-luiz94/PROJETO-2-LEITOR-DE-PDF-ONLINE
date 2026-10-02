"""
scripts/autonomo_verificar_ambiente.py — Diagnóstico do modo autônomo no SEU computador (TASK-031, roteiro de testes no Windows).

Uso (na pasta do projeto):   python scripts\\autonomo_verificar_ambiente.py
Não altera banco nem arquivos do projeto (só cria e apaga uma pasta temporária). Imprime OK / FALHA / AVISO por verificação e, no fim,
um bloco "COPIE ISTO" para você colar na conversa. Código de saída 0 = nenhuma FALHA.
"""
import os
import platform
import shutil
import sys
import tempfile
import time

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
RESULTADOS = []


def registrar(estado, nome, detalhe=""):
    RESULTADOS.append((estado, nome, detalhe))
    print(f"[{estado:5}] {nome}" + (f" — {detalhe}" if detalhe else ""))


def verificar(nome, fn, critico=True):
    try:
        detalhe = fn()
        registrar("OK", nome, detalhe or "")
        return True
    except Exception as e:  # noqa: BLE001
        registrar("FALHA" if critico else "AVISO", nome, f"{type(e).__name__}: {e}")
        return False


def v_python():
    ok = sys.version_info[:2] == (3, 11)
    if not ok:
        raise RuntimeError(f"Python {platform.python_version()} (o executável é gerado com 3.11)")
    return f"Python {platform.python_version()} {platform.architecture()[0]} em {platform.platform()}"


def v_import_quickjs():
    import quickjs
    return f"quickjs em {os.path.dirname(quickjs.__file__)}"


def v_eval():
    import quickjs
    ctx = quickjs.Context()
    r = ctx.eval("[1,2,3].map(function (x) { return x * 2; }).join('-') + ' ' + /^(\\d+)-x/i.test('12-X')")
    if str(r) != "2-4-6 true":
        raise RuntimeError(f"resposta inesperada: {r!r}")
    return "QuickJS executa JavaScript e regex"


REGRAS = [{"fase": 1, "ordem": 1, "modo": "DEFINIR", "texto_regex": "AFASTADOR", "ativo_template": "1-AF", "parar": True}]
ITEM = {"pagina": 1, "texto": "AFASTADOR", "cor": "#ff0000", "layer": ""}
leitor = None


def v_leitor():
    global leitor
    from services.autonomo.leitor_js import LeitorJS
    t = time.time()
    leitor = LeitorJS()
    inicio = time.time() - t
    t = time.time()
    r = leitor.processar_lote([ITEM], REGRAS, [], "extracao")
    if r[0]["ativo"] != "1-AF":
        raise RuntimeError(f"resultado inesperado: {r}")
    return f"processo auxiliar subiu em {inicio:.2f}s; 1ª chamada em {(time.time() - t) * 1000:.0f} ms"


def v_motor_real():
    import json
    from services.autonomo.leitor_js import LeitorJS  # noqa: F401
    with open(os.path.join(RAIZ, "data", "regras_leitor_processamento_seed.json"), encoding="utf-8") as f:
        proc = json.load(f)
    with open(os.path.join(RAIZ, "data", "regras_leitor_classificacao_seed.json"), encoding="utf-8") as f:
        cls = json.load(f)
    itens = [{"pagina": 1, "texto": t, "cor": "#ff0000", "layer": ""} for t in ("AFASTADOR", "INST. 01 - U4", "10 METROS", "CAA 2 ABC 35 m", "TROCAR 2 RS M AC")]
    r = leitor.processar_lote(itens, proc, cls, "extracao")
    return "regras de semente: " + "; ".join(f"{i['texto']!r}→{x['ativo']!r}/{x['entidade']}" for i, x in zip(itens, r))


def v_regex_catastrofica():
    import services.autonomo.leitor_js as m
    antigo = m.LIMITE_SEGUNDOS
    m.LIMITE_SEGUNDOS = 3
    lt = m.LeitorJS()
    try:
        ruim = [{"fase": 1, "ordem": 1, "modo": "DEFINIR", "texto_regex": "^(a+)+$", "ativo_template": "X", "parar": True}]
        t = time.time()
        try:
            lt.processar_lote([{"pagina": 1, "texto": "a" * 40 + "b", "cor": "#ff0000", "layer": ""}], ruim, [], "extracao")
            raise RuntimeError("era esperado estourar o tempo")
        except m.ErroLeitorJS as e:
            if "passou de" not in str(e):
                raise
        gasto = time.time() - t
        r = lt.processar_lote([ITEM], REGRAS, [], "extracao")      # reiniciou sozinho?
        if r[0]["ativo"] != "1-AF":
            raise RuntimeError("não se recuperou")
        return f"regra travada interrompida em {gasto:.1f}s e o motor se recuperou"
    finally:
        m.LIMITE_SEGUNDOS = antigo
        lt.fechar()


def v_sem_console():
    p = leitor._proc
    if p.stdin is not None or p.stdout is not None or p.poll() is not None:
        raise RuntimeError("o processo auxiliar não deveria depender de stdin/stdout")
    return f"processo auxiliar PID {p.pid} sem stdin/stdout (canal por socket local)"


def v_processo_morre_com_o_pai():
    import subprocess
    codigo = ("import sys, time; sys.path.insert(0, r'%s'); from services.autonomo.leitor_js import LeitorJS; lt = LeitorJS(); "
              "print(lt._proc.pid, flush=True); time.sleep(60)" % RAIZ)
    pai = subprocess.Popen([sys.executable, "-c", codigo], stdout=subprocess.PIPE, text=True)
    pid_filho = int(pai.stdout.readline())
    pai.kill()
    pai.wait()
    limite = time.time() + 10
    while time.time() < limite:
        if not _pid_vivo(pid_filho):
            return f"depois que o programa principal morreu, o processo auxiliar {pid_filho} encerrou sozinho"
        time.sleep(0.3)
    raise RuntimeError(f"o processo auxiliar {pid_filho} ficou órfão (encerre-o no Gerenciador de Tarefas)")


def _pid_vivo(pid):
    if os.name == "nt":
        import subprocess
        saida = subprocess.run(["tasklist", "/FI", f"PID eq {pid}"], capture_output=True, text=True).stdout
        return str(pid) in saida
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def v_pastas_com_acento():
    base = tempfile.mkdtemp(prefix="Área Autônomo ")
    try:
        origem = os.path.join(base, "entrada", "027")
        os.makedirs(origem)
        arq = os.path.join(origem, "obra ção ñ.dxf")
        with open(arq, "wb") as f:
            f.write(b"x" * 1000)
        destino = os.path.join(base, "processados", "027")
        os.makedirs(destino)
        shutil.move(arq, os.path.join(destino, os.path.basename(arq)))
        if not os.path.exists(os.path.join(destino, "obra ção ñ.dxf")):
            raise RuntimeError("o arquivo não chegou ao destino")
        return f"mover com acentos e espaços funciona ({base})"
    finally:
        shutil.rmtree(base, ignore_errors=True)


def v_arquivo_aberto():
    base = tempfile.mkdtemp(prefix="trava ")
    try:
        arq = os.path.join(base, "a.dxf")
        with open(arq, "wb") as f:
            f.write(b"x")
        aberto = open(arq, "rb")
        try:
            try:
                shutil.move(arq, os.path.join(base, "b.dxf"))
                return "mover um arquivo ABERTO funcionou (comportamento de Linux/Mac; no Windows costuma falhar)"
            except PermissionError:
                return "mover um arquivo aberto dá PermissionError (esperado no Windows): o programa deixa o arquivo na pasta e tenta de novo"
        finally:
            aberto.close()
    finally:
        shutil.rmtree(base, ignore_errors=True)


def v_pasta_padrao():
    import config
    destino = os.path.join(config.BASE_DIR, "autonomo")
    existia = os.path.isdir(destino)
    teste = os.path.join(destino, "_teste_escrita")
    os.makedirs(teste, exist_ok=True)
    os.rmdir(teste)
    if not existia:
        os.rmdir(destino)
    return f"pasta padrão do modo autônomo gravável: {destino}"


def v_dependencias():
    import ezdxf, pymupdf  # noqa: E401,F401
    return f"ezdxf {ezdxf.__version__}, pymupdf {pymupdf.__doc__.splitlines()[0] if pymupdf.__doc__ else 'ok'}"


def v_node():
    achou = shutil.which("node")
    if not achou:
        raise RuntimeError("node não encontrado (só é necessário para os testes de paridade do pytest; o programa não usa)")
    return achou


if __name__ == "__main__":
    print("== Diagnóstico do modo autônomo ==")
    verificar("Python 3.11", v_python)
    ok_q = verificar("Dependência quickjs instalada", v_import_quickjs)
    if ok_q:
        verificar("QuickJS executa JavaScript", v_eval)
        if verificar("Processo auxiliar do leitor sobe e responde", v_leitor):
            verificar("Motor com as regras de semente", v_motor_real)
            verificar("Processo auxiliar sem stdin/stdout (como no .exe sem console)", v_sem_console)
            leitor.fechar()
        verificar("Regra travada (regex catastrófica) é interrompida", v_regex_catastrofica)
        verificar("Processo auxiliar não fica órfão quando o programa fecha", v_processo_morre_com_o_pai)
    verificar("Pastas com acento e espaço", v_pastas_com_acento)
    verificar("Mover arquivo aberto", v_arquivo_aberto, critico=False)
    verificar("Pasta padrão gravável", v_pasta_padrao)
    verificar("ezdxf e pymupdf", v_dependencias)
    verificar("node (opcional, só para o pytest)", v_node, critico=False)
    print("\n===== COPIE ISTO =====")
    print(f"{platform.platform()} | Python {platform.python_version()} | frozen={getattr(sys, 'frozen', False)}")
    for estado, nome, detalhe in RESULTADOS:
        print(f"[{estado}] {nome}" + (f" — {detalhe}" if estado != "OK" else ""))
    sys.exit(1 if any(e == "FALHA" for e, _, _ in RESULTADOS) else 0)
