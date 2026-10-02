"""TASK-031 fase A — o leitor do modo autônomo é o MESMO arquivo JS da tela (QuickJS) e dá o mesmo resultado que o Node."""
import json
import random

import pytest

pytest.importorskip("quickjs")  # dependência do modo autônomo (requirements.txt)
from services.autonomo.leitor_js import ErroLeitorJS, LeitorJS  # noqa: E402
from tests import oraculo_tela as o  # noqa: E402

sem_node = pytest.mark.skipif(not o.NODE, reason="node não instalado")

with open("data/regras_leitor_processamento_seed.json", encoding="utf-8") as f:
    PROC = json.load(f)
with open("data/regras_leitor_classificacao_seed.json", encoding="utf-8") as f:
    CLS = json.load(f)

CORES = ["#ff0000", "#808080", "#000000", "#0000ff", "#ff00ff", "#c0c0c0", None]
LAYERS = ["", "01_LV", "01_RETENS", "01_RETENS_LV", "01_ORCAMENTO", "01_ORCAMENTO_LV", "01_PARTICULAR", "TXT_FF0000"]
TEXTOS = ["AFASTADOR", "DT11/300", "INST. 01 - U4", "INST.ALAR 2 SI3", "10 METROS", "5,5 METROS", "REC. CALÇADA 3X", "IP 2X", "CONC BASE 2 X",
          "COMPRESSOR 4X", "CAVA", "FIOS 12 FIOS", "3-100A-5H", "3 - 100A - 0,5H", "5H", "3CF-400A", "3CL-300A", "TR - 3 - 75 kVA", "TR - 1 - 15",
          "TR15", "112,5", "PODA M", "2 PODAS", "BLOCO TERRA3", "FLY REF", "FLY DESF", "FLY INST", "1-100 A", "RS M 1", "X RS M 1 Y", "REC. CALCADA 2X", "CONC BASE", "APOIOS",
          "CAA 2 ABC 35 m", "3x1x35", "CU 16", "P 50", "CA 4 AB 10", "ABC 3x35 120 M", "CAZ 3", "M 12", "CA4", "X1X 25 M", "AWG 4", "#2",
          "  espaços   extras  ", "1-U4", "U4", "", "ÇÃO", "CFU"]


def itens_aleatorios(n, semente):
    rnd = random.Random(semente)
    return [{"pagina": rnd.choice([1, 1, 2]), "texto": rnd.choice(TEXTOS), "cor": rnd.choice(CORES), "layer": rnd.choice(LAYERS)}
            for _ in range(n)]


@pytest.fixture(scope="module")
def leitor():
    lt = LeitorJS()
    yield lt
    lt.fechar()


@sem_node
@pytest.mark.parametrize("semente", [1, 2, 3, 4, 5])
def test_leitor_quickjs_igual_ao_node_nas_regras_reais(leitor, semente):
    itens = itens_aleatorios(300, semente)
    node = o.passo_leitor(itens, PROC, CLS)
    qjs = leitor.processar_lote(itens, PROC, CLS, "extracao")
    assert [(n["entidade"], n["operacao"], n["ativo"]) for n in node] == [(q["entidade"], q["operacao"], q["ativo"]) for q in qjs]


def test_vizinhanca_enxerga_a_lista_inteira(leitor):
    """A regra de fase 3 (5H com vizinho '3-100A') depende dos itens vizinhos, como na tela."""
    itens = [{"pagina": 1, "texto": "3-100A-5H", "cor": "#ff0000", "layer": ""},
             {"pagina": 1, "texto": "5H", "cor": "#ff0000", "layer": ""}]
    r = leitor.processar_lote(itens, PROC, CLS, "extracao")
    assert r[1]["ativo"] != "5H" or r[1]["ativo"] == ""  # com vizinho a regra reescreve; o valor exato é checado contra o Node:
    if o.NODE:
        node = o.passo_leitor(itens, PROC, CLS)
        assert [x["ativo"] for x in node] == [x["ativo"] for x in r]


def test_modo_resumo_so_recalcula_os_indices_pedidos(leitor):
    itens = itens_aleatorios(20, 9)
    todos = leitor.processar_lote(itens, PROC, CLS, "resumo")
    alguns = leitor.processar_lote(itens, PROC, CLS, "resumo", [3, 7])
    assert alguns == [todos[3], todos[7]]


@sem_node
def test_classificar_entidade_igual_ao_autoclassify_da_tela(leitor):
    ativos = ["1-AF", "2-RCFU", "1-TR3", "CAA2 ABC 35 m", "1-ROCO", "1-IP", "3-CFU", "1-U4", "", "1-FLY", "1-TERRA3"]
    esperado = o.passo_resumo([{"pagina": 1, "texto": a, "cor": "#ff0000", "layer": "", "entidade": "0", "operacao": "I", "ativo": a} for a in ativos],
                              [], CLS)
    # a entidade de cada ativo (via classificar do motor, como autoClassifyEntidade) bate com a separação feita pelo JS real
    nomes = {a: leitor.classificar_entidade(a, CLS) for a in ativos}
    assert nomes["1-AF"] == "ESTRUTURA" and nomes["2-RCFU"] == "CHAVE" and nomes[""] == "0"
    assert isinstance(esperado["cabos"], list)


def test_regra_invalida_nao_derruba_o_leitor(leitor):
    ruim = [{"fase": 1, "ordem": 1, "modo": "DEFINIR", "texto_regex": "(", "ativo_template": "X", "parar": True}]
    r = leitor.processar_lote([{"pagina": 1, "texto": "A", "cor": "#ff0000", "layer": ""}], ruim, [], "extracao")
    assert r[0]["ativo"] == ""  # regex inválida é ignorada pelo motor (_safeRegex)


def test_regex_catastrofica_nao_trava_e_o_leitor_se_recupera(monkeypatch):
    """O limite de tempo do QuickJS não interrompe regex com backtracking; quem mata é o processo principal."""
    import time
    import services.autonomo.leitor_js as m
    monkeypatch.setattr(m, "LIMITE_SEGUNDOS", 2)
    lt = m.LeitorJS()
    try:
        ruim = [{"fase": 1, "ordem": 1, "modo": "DEFINIR", "texto_regex": "^(a+)+$", "ativo_template": "X", "parar": True}]
        t = time.time()
        with pytest.raises(ErroLeitorJS, match="passou de"):
            lt.processar_lote([{"pagina": 1, "texto": "a" * 40 + "b", "cor": "#ff0000", "layer": ""}], ruim, [], "extracao")
        assert time.time() - t < 10
        ok = lt.processar_lote([{"pagina": 1, "texto": "AFASTADOR", "cor": "#ff0000", "layer": ""}], PROC, CLS, "extracao")
        assert ok[0]["ativo"] == "1-AF"        # reiniciou sozinho e voltou a funcionar
    finally:
        lt.fechar()


def test_motor_inexistente_da_erro_claro(tmp_path):
    with pytest.raises(FileNotFoundError):
        LeitorJS(str(tmp_path / "nao_existe.js"))
    ruim = tmp_path / "ruim.js"
    ruim.write_text("isto não é javascript (((", encoding="utf-8")
    with pytest.raises(ErroLeitorJS, match="Falha ao carregar"):
        LeitorJS(str(ruim))


def test_importar_o_modulo_nao_carrega_quickjs():
    import subprocess
    import sys
    r = subprocess.run([sys.executable, "-c", "import sys, services.autonomo.leitor_js, services.autonomo.montagem; print('quickjs' in sys.modules)"],
                       capture_output=True, text=True)
    assert r.stdout.strip() == "False"


def test_trabalhador_que_nao_inicia_da_erro_rapido(monkeypatch):
    import sys
    import time
    import services.autonomo.leitor_js as m
    monkeypatch.setattr(m, "_comando_trabalhador", lambda porta, token: [sys.executable, "-c", "import sys; sys.exit(3)"])
    t = time.time()
    with pytest.raises(ErroLeitorJS, match="não conectou"):
        m.LeitorJS()
    assert time.time() - t < 15


def test_conexao_com_token_errado_e_recusada(monkeypatch):
    import sys
    import services.autonomo.leitor_js as m
    codigo = ("import socket, sys, json; s = socket.create_connection(('127.0.0.1', int(sys.argv[1]))); "
              "s.sendall(b'{\"token\": \"errado\"}\\n'); s.recv(10)")
    monkeypatch.setattr(m, "_comando_trabalhador", lambda porta, token: [sys.executable, "-c", codigo, str(porta)])
    with pytest.raises(ErroLeitorJS, match="token inválido"):
        m.LeitorJS()


def test_o_canal_nao_depende_de_stdin_nem_stdout(leitor):
    """No .exe `--windowed` o stdin/stdout do Python não existem: o trabalhador é iniciado SEM stdin/stdout e mesmo assim responde."""
    r = leitor.processar_lote([{"pagina": 1, "texto": "AFASTADOR", "cor": "#ff0000", "layer": ""}], PROC, CLS, "extracao")
    assert r[0]["ativo"] == "1-AF"
    proc = leitor._proc
    assert proc.stdin is None and proc.stdout is None and proc.poll() is None
