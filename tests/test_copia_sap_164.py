"""TASK-164 — visão SAP estilo Excel: cópia idêntica à do Excel, colunas da cópia e cópia em blocos."""
import json
import os
import subprocess

import pytest

import config
from tests import oraculo_tela as o

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado")
MOD = os.path.join(config.STATIC_DIR, "orcamento_sap.js")
HTML = os.path.join(config.STATIC_DIR, "resultado_orcamento.html")


def js(expr, **entrada):
    codigo = f"const S=require({json.dumps(MOD)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({expr}));"
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


M = [["M", "1", "1,50", "1600"], ["N", "X", "2", "1000"]]


def test_texto_excel_termina_toda_linha_com_crlf():
    assert js("S.textoCopia(e.m,'excel')", m=M) == "M\t1\t1,50\t1600\r\nN\tX\t2\t1000\r\n"
    assert js("S.textoCopia(e.m,'texto')", m=M) == "M\t1\t1,50\t1600\r\nN\tX\t2\t1000\r\n"


def test_formato_lf_e_o_de_antes_sem_quebra_final():
    assert js("S.textoCopia(e.m,'lf')", m=M) == "M\t1\t1,50\t1600\nN\tX\t2\t1000"


def test_celula_com_tab_ou_quebra_nao_desalinha():
    assert js("S.textoCopia(e.m,'excel')", m=[["a\tb\nc", "d"]]) == "a b c\td\r\n"


def test_html_e_tabela_escapada():
    h = js("S.htmlCopia(e.m)", m=[["<A&B>", "1000"]])
    assert "<table>" in h and "&lt;A&amp;B&gt;" in h and "<td>1000</td>" in h and "StartFragment" in h


def test_formato_invalido_vira_excel_e_so_excel_usa_html():
    assert js("[S.formatoValido('x'),S.formatoValido('lf'),S.formatoUsaHtml('excel'),S.formatoUsaHtml('texto'),S.formatoUsaHtml('lf')]") == ["excel", "lf", True, False, False]


def test_colunas_da_copia_padrao_e_ordem_fixa():
    assert js("S.colunasDaCopia(null)") == ["mdo", "codigo", "qtd", "centro"]
    assert js("S.colunasDaCopia(['centro','mdo'])") == ["mdo", "centro"]
    assert js("S.colunasDaCopia([])") == ["mdo", "codigo", "qtd", "centro"]
    assert js("S.colunasDaCopia(['ref','x'])") == ["mdo", "codigo", "qtd", "centro"]   # a coluna "#" nunca entra


def test_andamento_dos_blocos():
    a = js("S.andamentoBlocos(48,10,10)")
    assert (a["bloco"], a["blocos"], a["faltam"], a["ini"], a["fim"], a["terminou"]) == (1, 5, 38, 10, 20, False)
    u = js("S.andamentoBlocos(15,10,10)")
    assert (u["ini"], u["fim"]) == (10, 15)                       # o último bloco é o que sobrou
    t = js("S.andamentoBlocos(15,10,15)")
    assert t["terminou"] and t["faltam"] == 0 and t["percentual"] == 100
    assert js("S.andamentoBlocos(0,10,0)")["total"] == 0


def test_bloco_anterior():
    assert js("S.blocoAnterior(20,10)") == {"ini": 10, "fim": 20}
    assert js("S.blocoAnterior(0,10)") is None


def test_tela_usa_os_recursos_novos():
    h = open(HTML, encoding="utf-8").read()
    for chave in ("sap-copia-barra", "copiarBlocoSap", "SAP.textoCopia", "clipboardData.setData", "sap-rn", "editarCelulaSap", "data-acao=\"bloco-prox\""):
        assert chave in h, chave
    assert 'contenteditable="true" class="sap-mdo"' not in h      # a célula só edita depois de 2 cliques / F2
