"""TASK-039 (plugin): no nanoCAD 5 a API declara Document.CommandInProgress mas não implementa — a consulta abria o «API Logger»
(MGD · NOT_IMPLEMENTED · get_CommandInProgress) nos botões do painel e o Salvar da montagem manual não seguia. A DLL do nanoCAD 5 não usa mais."""
import os

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLUGIN = os.path.join(RAIZ, "plugin_nanocad")


def test_consulta_do_comando_em_andamento_so_fora_do_nanocad5():
    src = open(os.path.join(PLUGIN, "src", "Plugin", "Tema.cs"), encoding="utf-8").read()
    i = src.index("d.CommandInProgress")
    antes = src[:i]
    assert antes.rindex("#if !NANOCAD5") > antes.rindex("public static void Executar(") and "#endif" in src[i:i + 200]


def test_dll_do_nanocad5_nao_referencia_command_in_progress():
    with open(os.path.join(PLUGIN, "appload", "nanoCAD5", "OrcamentoNanoCAD.dll"), "rb") as f:
        assert b"get_CommandInProgress" not in f.read()
    with open(os.path.join(PLUGIN, "appload", "nanoCAD26", "OrcamentoNanoCAD.dll"), "rb") as f:
        assert b"get_CommandInProgress" in f.read()                                  # no nanoCAD 26 a verificação continua


# ── etapa 2 (0.16.6): nanoCAD 26 travava no «Abrir no OrçaCore» / «Adicionar anotações ao OrçaCore» ──
def _src(*partes):
    with open(os.path.join(PLUGIN, "src", *partes), encoding="utf-8") as f:
        return f.read()


def test_abrir_no_orcacore_conversa_em_segundo_plano_e_volta_para_a_tela():
    src = _src("Plugin", "ProgramaNoComputador.cs")
    i = src.index("public static void AbrirProjeto(")
    corpo = src[i:src.index("private static string Entregar(", i)]
    assert "ThreadPool.QueueUserWorkItem" in corpo and "Aplicacao.NaTela(() => Aplicacao.Msg(msg))" in corpo
    assert "ProgramaLocal.EntregarProjeto" not in corpo                                   # a entrega não roda mais na thread do nanoCAD
    assert "abrir no OrçaCore: porta " in src and "anotações: lote " in src               # tempos no log.txt
    apl = _src("Plugin", "Aplicacao.cs")
    assert "public static void NaTela(Action acao)" in apl and "PrepararTela();" in apl


def test_chamadas_ao_orcacore_local_sem_proxy_nem_keepalive():
    src = _src("Nucleo", "ClienteAplicacao.cs")
    i = src.index("if (req.RequestUri.IsLoopback)")
    assert "req.Proxy = null;" in src[i:i + 500] and "req.KeepAlive = false;" in src[i:i + 500]
