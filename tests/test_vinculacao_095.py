"""TASK-095 — só operação I, M, *I e *M entram no vínculo (R e *R ficam de fora) e o cabo aceita estruturas fictícias. Python = JS."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import vinculacao_regras as vr

NODE = shutil.which("node")
CFG = {"cabos": [{"padrao": "^C\\d+$", "valor_cabo": "C", "pontas": {"I": 2, "M": 1}}],
       "estruturas": [{"padrao": "^N1$", "qtd_cabos": 2}, {"padrao": "^N3$", "qtd_cabos": 1, "amarracao": True}]}


def primeiro(t):
    return (t or "").split(" ")[0].upper()


def cab(a, op="I", v=None, fict=None):
    d = {"entidade": "CABO", "operacao": op, "ativo": a}
    if v is not None:
        d["vinculoEstruturas"] = v
    if fict:
        d["estruturasFicticias"] = fict
    return d


def est(a, op="I"):
    return {"entidade": "ESTRUTURA", "operacao": op, "ativo": a}


def test_participa_vinculo():
    assert all(vr.participa_vinculo({"operacao": o}) for o in ("I", "*I", "M", "*M", "i", "", None))
    assert not any(vr.participa_vinculo({"operacao": o}) for o in ("R", "*R", "r", " *r "))


def test_cabo_e_estrutura_r_nao_entram_na_analise():
    a = vr.analisar([cab("C16", "R", [0]), cab("C25", "*R"), cab("C35", "I", [1])], [est("N1", "I"), est("N1", "*R")], CFG, primeiro, primeiro)
    assert set(a["cabos"]) == {2}                    # só o cabo I
    assert set(a["estruturas"]) == {0}               # só a estrutura I
    assert a["estruturas"][0]["tem"] == 0            # o vínculo do cabo R não conta
    assert vr.estruturas_disponiveis([est("N1", "R"), est("N1", "*M")]) == [1]


def test_estrutura_r_nao_recebe_vinculo_do_script():
    cabos = [dict(cab("C16", "I"), _x=1, _y=0)]
    outros = [dict(est("N1", "R"), _x=1, _y=0), dict(est("N1", "I"), _x=5, _y=0)]
    r = vr.executar_script({"versao": 1, "passos": [{"modo": "proximidade", "pontas": 1}]}, cabos, outros, CFG, primeiro, primeiro, True)
    assert r["vinculos"][0]["estruturas"] == [1] and cabos[0]["vinculoEstruturas"] == [1]


def test_estrutura_ficticia_ocupa_uma_ponta_do_cabo():
    a = vr.analisar([cab("C16", "I", [0], ["N1"])], [est("N1")], CFG, primeiro, primeiro)
    assert a["cabos"][0]["tem"] == 2 and a["cabos"][0]["estado"] == "gr" and a["cabos"][0]["ficticias"] == 1
    a = vr.analisar([cab("C16", "I", None, ["N1"])], [est("N1")], CFG, primeiro, primeiro)
    assert a["cabos"][0]["tem"] == 1 and a["cabos"][0]["estado"] == "pu"                  # 1 de 2 pontas
    a = vr.analisar([cab("C16", "I", None, ["T9"])], [], {"cabos": CFG["cabos"], "estruturas": [{"padrao": "^T9$", "cabos": "^A"}]}, primeiro, primeiro)
    assert a["cabos"][0]["estado"] == "rd" and a["cabos"][0]["erros"]                    # a estrutura fictícia não aceita o cabo
    a = vr.analisar([cab("C16", "I", None, ["N1"])], [], CFG, primeiro, primeiro)
    assert a["cabos"][0]["estado"] == "pu"                                                 # fictícia não precisa de linha em Outros


def test_script_considera_ficticias_do_cabo():
    cabos = [dict(cab("C16", "I", None, ["N3"]), _x=0, _y=0), dict(cab("C25", "I"), _x=0, _y=0)]
    outros = [dict(est("N1"), _x=1, _y=0), dict(est("N1"), _x=2, _y=0)]
    r = vr.executar_script({"versao": 1, "passos": [{"modo": "proximidade"}]}, cabos, outros, CFG, primeiro, primeiro, True)
    assert r["passos"][0]["ja_tinham"] == 1                       # o cabo com estrutura fictícia já tem vínculo
    assert [v["cabo"] for v in r["vinculos"]] == [1]


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_js():
    js = """
    const VR = require(process.argv[1]); const p = t => (t || '').split(' ')[0].toUpperCase(); const cfg = %s;
    const cab = (a, op, v, f) => { const d = { entidade: 'CABO', operacao: op, ativo: a }; if (v) d.vinculoEstruturas = v; if (f) d.estruturasFicticias = f; return d; };
    const est = (a, op) => ({ entidade: 'ESTRUTURA', operacao: op, ativo: a });
    const a1 = VR.analisar([cab('C16', 'R', [0]), cab('C25', '*R'), cab('C35', 'I', [1])], [est('N1', 'I'), est('N1', '*R')], cfg, p, p);
    const a2 = VR.analisar([cab('C16', 'I', [0], ['N1'])], [est('N1', 'I')], cfg, p, p);
    const a3 = VR.analisar([cab('C16', 'I', null, ['T9'])], [], { cabos: cfg.cabos, estruturas: [{ padrao: '^T9$', cabos: '^A' }] }, p, p);
    const cabos = [Object.assign(cab('C16', 'I', null, ['N3']), { _x: 0, _y: 0 }), Object.assign(cab('C25', 'I'), { _x: 0, _y: 0 })];
    const outros = [Object.assign(est('N1', 'I'), { _x: 1, _y: 0 }), Object.assign(est('N1', 'I'), { _x: 2, _y: 0 })];
    const sc = VR.executarScript({ versao: 1, passos: [{ modo: 'proximidade' }] }, cabos, outros, cfg, p, p, true);
    console.log(JSON.stringify({ a1, a2, a3, sc: sc.vinculos, ja: sc.passos[0].ja_tinham }));
    """ % json.dumps(CFG)
    s = json.loads(subprocess.run([NODE, "-e", js, os.path.join(config.STATIC_DIR, "vinculacao_regras.js")], capture_output=True, text=True, check=True).stdout)
    py1 = vr.analisar([cab("C16", "R", [0]), cab("C25", "*R"), cab("C35", "I", [1])], [est("N1", "I"), est("N1", "*R")], CFG, primeiro, primeiro)
    py2 = vr.analisar([cab("C16", "I", [0], ["N1"])], [est("N1", "I")], CFG, primeiro, primeiro)
    py3 = vr.analisar([cab("C16", "I", None, ["T9"])], [], {"cabos": CFG["cabos"], "estruturas": [{"padrao": "^T9$", "cabos": "^A"}]}, primeiro, primeiro)
    for js_a, py in ((s["a1"], py1), (s["a2"], py2), (s["a3"], py3)):
        assert {int(k): v for k, v in js_a["cabos"].items()} == py["cabos"]
        assert {int(k): v for k, v in js_a["estruturas"].items()} == py["estruturas"]
        assert js_a["links"] == py["links"]
    assert s["ja"] == 1 and [v["cabo"] for v in s["sc"]] == [1]
