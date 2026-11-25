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
