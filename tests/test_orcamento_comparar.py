"""TASK-071 — lógica pura do modo comparação (static/orcamento_comparar.js)."""
import json
import os
import subprocess

import pytest

import config
from tests import oraculo_tela as o

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado")
MOD = os.path.join(config.STATIC_DIR, "orcamento_comparar.js")


def js(expr, **entrada):
    codigo = f"const C=require({json.dumps(MOD)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({expr}));"
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def A(op, mdo, cod, qtd, desc=""):
    return {"operacao": op, "mdo": mdo, "codigo": cod, "qtd": qtd, "desc": desc}


def test_parse_rec_consolida_repetidos_e_traduz_codigos():
    t = "100;2,5;M;IN;ODI;LM\n100;1;M;IN;ODI;LM\n200;3;S;DR;ODD;LM\n\nlinha-ruim\n"
    r = js("C.parseRec(e.t)", t=t)
    assert [(x["codigo"], x["operacao"], x["mdo"], x["qtd"]) for x in r] == [("100", "I", "MATERIAL", 3.5), ("200", "R", "MAO-DE-OBRA", 3.0)]


def test_chave_normaliza_operacao_mdo_e_codigo():
    assert js("[C.chave('i','material',' ab1 '), C.chave('R','M','AB1'), C.chave('','','x'), C.chave('I','N','x'), C.chave('I','MÃO-DE-OBRA','x')]") == \
        ["I|MATERIAL|AB1", "R|MATERIAL|AB1", "I|MAO-DE-OBRA|X", "I|MAO-DE-OBRA|X", "I|MAO-DE-OBRA|X"]


def test_os_quatro_status_e_o_zero_do_outro_lado():
    atual = [A("I", "MATERIAL", "1", 5), A("I", "MATERIAL", "2", 3), A("I", "MATERIAL", "3", 9)]
    ref = [A("I", "MATERIAL", "1", 5), A("I", "MATERIAL", "2", 4), A("I", "MATERIAL", "4", 7, "só ref")]
    r = js("C.comparar(e.a, e.r)", a=atual, r=ref)
    assert [(l["status"], l["ref"]) for l in r["linhas"]] == [("igual", 5), ("diferente", 4), ("so_atual", 0)]      # só no atual: referência = 0
    assert [(f["codigo"], f["qtd"], f["desc"]) for f in r["fantasmas"]] == [("4", 7, "só ref")]                       # só na referência: atual = 0
    assert r["contagem"] == {"igual": 1, "diferente": 1, "so_atual": 1, "so_referencia": 1}


def test_mesma_chave_com_operacao_diferente_sao_itens_distintos():
    r = js("C.comparar(e.a, e.r)", a=[A("I", "MATERIAL", "1", 5)], r=[A("R", "MATERIAL", "1", 5)])
    assert r["linhas"][0]["status"] == "so_atual" and len(r["fantasmas"]) == 1 and r["fantasmas"][0]["operacao"] == "R"


def test_igualdade_em_duas_casas_e_negativos():
    ig = lambda a, b: js("C.igualQtd(e.a, e.b)", a=a, b=b)
    assert ig(1.004, 1.0) is True and ig(1.006, 1.0) is False and ig(-3, -3.0) is True and ig(0, 0.001) is True
    r = js("C.comparar(e.a, e.r)", a=[A("I", "M", "1", -2)], r=[A("I", "M", "1", -2)])
    assert r["linhas"][0]["status"] == "igual"


def test_linhas_repetidas_comparam_a_soma_e_mostram_a_referencia_so_na_primeira():
    atual = [A("I", "MATERIAL", "1", 2), A("I", "MATERIAL", "1", 3), A("I", "MATERIAL", "9", 1)]
    r = js("C.comparar(e.a, e.r)", a=atual, r=[A("I", "MATERIAL", "1", 5)])
    assert [(l["status"], l["ref"]) for l in r["linhas"]] == [("igual", 5), ("igual", None), ("so_atual", 0)]
    assert r["contagem"]["igual"] == 1 and r["contagem"]["so_atual"] == 1


def test_referencia_vazia_deixa_tudo_so_no_atual_e_atual_vazio_tudo_fantasma():
    assert js("C.comparar(e.a, []).contagem", a=[A("I", "M", "1", 1)]) == {"igual": 0, "diferente": 0, "so_atual": 1, "so_referencia": 0}
    r = js("C.comparar([], e.r)", r=[A("I", "M", "1", 1), A("R", "N", "2", 2)])
    assert len(r["fantasmas"]) == 2 and r["linhas"] == []


def test_fantasma_entra_na_ordem_padrao_mdo_operacao_descricao():
    linhas = [A("I", "MAO-DE-OBRA", "1", 1, "ALFA"), A("I", "MAO-DE-OBRA", "2", 1, "ZETA"), A("I", "MATERIAL", "3", 1, "BETA"), A("R", "MATERIAL", "4", 1, "ALFA")]
    pos = lambda item: js("C.posicaoInsercao(e.l, e.i, C.comparadorPadrao)", l=linhas, i=item)
    assert pos(A("I", "MAO-DE-OBRA", "x", 0, "MEIO")) == 1                  # entre ALFA e ZETA (mão de obra primeiro)
    assert pos(A("I", "MATERIAL", "x", 0, "ZZZ")) == 3                      # material I depois de BETA e antes do material R
    assert pos(A("R", "MATERIAL", "x", 0, "ZZZ")) == 4                      # no fim
    assert pos(A("I", "MAO-DE-OBRA", "x", 0, "AAA")) == 0
