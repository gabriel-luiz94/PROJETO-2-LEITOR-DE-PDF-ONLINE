"""TASK-124 — poste cadastrado na base com o nome COMPACTO como ativo (``121000R, R12/1000``) continua sendo poste e preenche o parâmetro
``poste`` das estruturas da linha, no formato compacto e no antigo; paridade Python × JS."""
import json
import shutil
import subprocess

import pytest

from services import anotacao_outros as ao
from services import orcamento_calc as calc

NODE = shutil.which("node")
cl = calc._classico_outros


def L(ativo, comp="", cod="", fi=1, mdo="MATERIAL"):
    return {"ativo": ativo, "desc_ativo": ativo and "D", "componente": comp, "projeto": "PERNAMBUCO", "mdo": mdo, "codigo": cod, "desc_codigo": f"d {cod}",
            "fator_i": fi, "fator_r": 0.0, "filtro": "", "origem": "OUTROS", "lv": "A"}


BASE = [
    L("121000R, R12/1000", "POSTE R 12/1000", "3302058"),                 # poste com o nome compacto como ATIVO (e não RC12/1000)
    L("RC10/300, 10300R", "RC10/300", "rc300"),                          # poste no padrão antigo
    L("N1=N1(poste, tipocabo)", "N1", "n1a"),
    L("N1(poste=%1000R|%D|%R, tipocabo=A)", "N1", "n1b"),
]
CASOS = ["121000R, N1", "R12/1000, N1", "1-121000R 1-N1", "10300R, N1", "N1, 121000R", "121000R"]


def ctx():
    return ao.contexto_da_base(BASE, "PERNAMBUCO")


@pytest.mark.parametrize("nome", ["121000R", "R12/1000", "RC12/1000", "DT11/600", "9400D"])
def test_nomes_de_poste(nome):
    assert ao.eh_poste(nome)


@pytest.mark.parametrize("nome", ["N1", "CE3 PR", "10D", "ATERINT", "500R", ""])
def test_nao_poste(nome):
    assert not ao.eh_poste(nome)


@pytest.mark.parametrize("linha", ["121000R, N1", "R12/1000, N1", "1-121000R 1-N1"])
def test_poste_com_nome_compacto_preenche_o_parametro(linha):
    r = ao.interpretar_linha(linha, "I", cl, ctx())
    assert r["erros"] == []
    assert [i["ativo"] for i in r["itens"]] == ["121000R", "N1(POSTE=121000R)"]


def test_poste_a_direita_nao_preenche_e_poste_sozinho_vale():
    assert [i["ativo"] for i in ao.interpretar_linha("N1, 121000R", "I", cl, ctx())["itens"]] == ["N1", "121000R"]
    assert [i["ativo"] for i in ao.interpretar_linha("121000R", "I", cl, ctx())["itens"]] == ["121000R"]


def test_operacao_m_ignora_o_poste_compacto():
    r = ao.interpretar_linha("121000R, N1", "M", cl, ctx())
    assert [(i["ativo"], i["operacao"]) for i in r["itens"]] == [("N1(POSTE=121000R)", "I")]
    assert ao.codigo_estrutura("121000R, N1", cl, ctx()) == "N1"


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_com_o_js():
    cod = """
const AO = require('./static/anotacao_outros.js'); const AP = require('./static/ativo_parametros.js');
const d = JSON.parse(require('fs').readFileSync(0, 'utf8'));
const classico = t => t.split(/\\s+/).filter(Boolean).map(x => { const m = x.match(/^(\\d+(?:[.,]\\d+)?)-(.+)$/); return { ativo: m ? m[2] : x, qtd: m ? parseFloat(m[1]) : 1 }; });
const ctx = AO.contextoDaBase(d.base, 'PERNAMBUCO');
console.log(JSON.stringify(d.casos.map(c => AO.interpretarLinha(c, 'I', classico, ctx).itens.map(i => [i.ativo, i.qtd]))));
"""
    r = subprocess.run([NODE, "-e", cod], input=json.dumps({"casos": CASOS, "base": BASE}), capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    js = json.loads(r.stdout)
    for c, itens in zip(CASOS, js):
        py = [[i["ativo"], i["qtd"]] for i in ao.interpretar_linha(c, "I", lambda t: [{"ativo": x.split("-", 1)[-1], "qtd": float(x.split("-", 1)[0]) if "-" in x else 1.0} for x in t.split()], ctx())["itens"]]
        assert py == itens, c
