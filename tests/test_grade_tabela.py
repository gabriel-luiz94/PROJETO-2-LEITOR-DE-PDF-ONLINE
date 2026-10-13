"""TASK-067 — lógica pura da grade estilo Excel (static/grade_tabela.js) no Node."""
import json
import os
import subprocess

import pytest

import config
from tests import oraculo_tela as o

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado")
MODULO = os.path.join(config.STATIC_DIR, "grade_tabela.js")


def js(expr, **entrada):
    codigo = f"const G=require({json.dumps(MODULO)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({expr}));"
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def colar(dados, texto, tipo="outros", linha=0, coluna=0, **extra):
    return js("(()=>{const d=e.dados;const r=G.aplicarColagem(d,e.texto,{tipo:e.tipo,linha:e.linha,coluna:e.coluna,maxLinhas:e.max});return {d,r};})()",
              dados=dados, texto=texto, tipo=tipo, linha=linha, coluna=coluna, max=extra.get("max"))


def L(op, ativo, ent="0", **kw):
    return {"entidade": ent, "operacao": op, "ativo": ativo, **kw}


def test_parse_tsv_crlf_aspas_e_linhas_vazias():
    r = js("G.parseColagem(e.t)", t='I\tA 1\r\n\r\n"R"\t"B\n2"\r\n')
    assert r["linhas"] == [["I", "A 1"], ["R", "B\n2"]]


def test_eh_colagem_multipla():
    assert js("[G.ehColagemMultipla('abc'),G.ehColagemMultipla('abc\\n'),G.ehColagemMultipla('a\\nb'),G.ehColagemMultipla('a\\tb')]") == [False, False, True, True]


def test_uma_coluna_vai_para_ativo_mesmo_comecando_na_operacao():
    r = colar([L("I", "X")], "1-u4\n2-u5\n3-u6", linha=0, coluna=0)
    assert [x["ativo"] for x in r["d"]] == ["1-U4", "2-U5", "3-U6"]
    assert [x["operacao"] for x in r["d"]] == ["I", "I", "I"] and r["r"]["criadas"] == 2


def test_uma_coluna_de_operacoes_vai_para_operacao():
    r = colar([L("I", "A"), L("I", "B")], "r\n*m", coluna=0)
    assert [x["operacao"] for x in r["d"]] == ["R", "*M"]


def test_duas_colunas_operacao_e_ativo_sobrescreve_e_estende():
    d = [L("I", "VELHO1"), L("I", "VELHO2"), L("R", "FICA")]
    r = colar(d, "R\t1-A\n*M\t2-B\nI\t3-C", linha=1, coluna=0)
    assert [(x["operacao"], x["ativo"]) for x in r["d"]] == [("I", "VELHO1"), ("R", "1-A"), ("*M", "2-B"), ("I", "3-C")]
    assert r["r"]["alteradas"] == [1, 2, 3] and r["r"]["criadas"] == 1


def test_operacao_invalida_mantem_a_atual_e_conta():
    r = colar([L("R", "A")], "ZZ\t1-X", coluna=0)
    assert r["d"][0]["operacao"] == "R" and r["d"][0]["ativo"] == "1-X" and r["r"]["opsInvalidas"] == 1


def test_cabos_tres_colunas_e_sinal_negativo():
    r = colar([L("I", "A", "CABO", qtdAtivos=5)], "R\tCAA 2 80 m\t-3", tipo="cabos", coluna=0)
    x = r["d"][0]
    assert x["operacao"] == "R" and x["ativo"] == "CAA 2 80 M" and x["qtdAtivos"] == "-3"


def test_coluna_sobrando_e_descartada_em_outros():
    r = colar([L("I", "A")], "I\tX\tEXTRA", tipo="outros", coluna=0)
    assert r["d"][0] == L("I", "X")


def test_linha_nova_de_cabos_nasce_cabo_e_herda_operacao():
    r = colar([L("R", "A", "CABO")], "1\n2", tipo="cabos", coluna=1)
    assert r["d"][1]["entidade"] == "CABO" and r["d"][1]["operacao"] == "R"


def test_tabela_vazia_cria_linhas_com_I():
    r = colar([], "a\nb", coluna=1)
    assert [(x["operacao"], x["ativo"]) for x in r["d"]] == [("I", "A"), ("I", "B")]


def test_limite_de_linhas():
    r = colar([], "\n".join(f"A{i}" for i in range(20)), coluna=1, max=5)
    assert len(r["d"]) == 5 and r["r"]["truncado"] is True


def test_intervalo_em_tsv_e_ida_e_volta():
    d = [L("I", "A", qtdAtivos=1), L("R", "B", qtdAtivos=2)]
    t = js("G.textoDoIntervalo(e.d,'cabos',0,0,1,2)", d=d)
    assert t == "I\tA\t1\nR\tB\t2"
    r = colar([L("M", "X", "CABO"), L("M", "Y", "CABO")], t, tipo="cabos", coluna=0)
    assert [(x["operacao"], x["ativo"]) for x in r["d"]] == [("I", "A"), ("R", "B")]


def test_limpar_intervalo_so_esvazia_ativo():
    r = js("(()=>{const d=e.d;const a=G.limparIntervalo(d,'outros',0,0,1,1);return {d,a};})()", d=[L("R", "A"), L("I", "B")])
    assert r["a"] == [0, 1] and [x["ativo"] for x in r["d"]] == ["", ""] and [x["operacao"] for x in r["d"]] == ["R", "I"]


def test_operacao_e_ativo_colados_com_o_cursor_na_coluna_ativo():
    r = colar([L("I", "A"), L("I", "B")], "R\t1-X\n*M\t2-Y", coluna=1)
    assert [(x["operacao"], x["ativo"]) for x in r["d"]] == [("R", "1-X"), ("*M", "2-Y")]


def test_ativo_e_quantidade_em_cabos_com_o_cursor_no_ativo():
    r = colar([L("I", "A", "CABO")], "CAA 2 80 m\t-3", tipo="cabos", coluna=1)
    assert r["d"][0]["ativo"] == "CAA 2 80 M" and r["d"][0]["qtdAtivos"] == "-3" and r["d"][0]["operacao"] == "I"
