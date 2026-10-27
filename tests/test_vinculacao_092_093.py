"""TASK-092 (Fases/Tensão/Nível do cabo a partir do texto original) e TASK-093 (estrutura com 1 cabo real + 1 fictício opcional): Python = JS."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import vinculacao_regras as vr

NODE = shutil.which("node")
TEXTOS = ["3xA35P(7,94)-13,8kV 45 m", "1xS1/0-7,9kV 20 m", "3xC16(16) 30 m", "RAMA25 12 m", "2xC25 8 m", "3xA70-0,6kV 10 m", "S1/0 - 13.8 kV", ""]


def primeiro(txt):
    return (txt or "").split(" ")[0].upper()


def test_derivar_cabo_fases_tensao_nivel():
    d = vr.derivar_cabo("3xA35P(7,94)-13,8kV 45 m")
    assert d == {"fases": 3, "tensao": "13,8kV", "nivel": "MT"}
    assert vr.derivar_cabo("1xS1/0-7,9kV 20 m")["nivel"] == "MT"
    assert vr.derivar_cabo("3xC16(16) 30 m") == {"fases": 3, "tensao": "", "nivel": "BT"}
    assert vr.derivar_cabo("RAMA25 12 m") == {"fases": 1, "tensao": "", "nivel": "BT"}          # sem Nx = 1x implícito
    assert vr.derivar_cabo("3xA70-0,6kV 10 m")["nivel"] == "BT"                                    # até 1 kV é BT


def test_regra_de_cabo_filtra_por_fases_nivel_e_texto():
    cfg = vr.normalizar_config({"cabos": [
        {"padrao": "^A35P$", "nivel": "MT", "fases": "3", "valor_cabo": "P"},
        {"padrao": "^A35P$", "valor_cabo": "X"},
        {"padrao": "^S1/0$", "padrao_texto": "7,9kV", "valor_cabo": "M79"}]})
    mt = vr.extra_cabo({"textoOriginal": "3xA35P(7,94)-13,8kV 45 m"})
    bt = vr.extra_cabo({"textoOriginal": "3xA35P 45 m"})
    assert vr.info_cabo("A35P", "I", cfg, mt)["valor_cabo"] == "P"
    assert vr.info_cabo("A35P", "I", cfg, bt)["valor_cabo"] == "X"                              # BT não casa a regra MT
    assert vr.info_cabo("A35P", "I", cfg)["valor_cabo"] == "P"                                  # sem extra (cabo fictício): filtros ignorados
    assert vr.info_cabo("S1/0", "I", cfg, vr.extra_cabo({"textoOriginal": "1xS1/0-7,9kV"}))["valor_cabo"] == "M79"
    assert vr.info_cabo("S1/0", "I", cfg, vr.extra_cabo({"textoOriginal": "1xS1/0"}))["regra"] is None
    # o que foi digitado na linha vale mais que o derivado do texto
    assert vr.extra_cabo({"textoOriginal": "3xC16", "fases": "1", "nivel": "mt"}) == {"texto": "3xC16", "descricao": "3XC16", "fases": 1, "tensao": "", "nivel": "MT"}


def test_validar_config_filtros_novos():
    ok = {"cabos": [{"padrao": "A", "fases": "1,3", "nivel": "MT", "padrao_texto": "13,8kV"}], "estruturas": [{"padrao": "SI3", "qtd_min": 1, "qtd_max": 2, "max_reais": 1}]}
    assert vr.validar_config(ok) == []
    ruim = {"cabos": [{"padrao": "A", "fases": "x", "nivel": "AT", "padrao_texto": "("}], "estruturas": [{"padrao": "SI3", "qtd_min": 3, "qtd_max": 2, "max_reais": 0}]}
    erros = vr.validar_config(ruim)
    assert any("fases" in e for e in erros) and any("nivel" in e for e in erros) and any("padrao_texto" in e for e in erros)
    assert any("qtd_min" in e and "maior" in e for e in erros) and any("max_reais" in e for e in erros)


CFG_SI3 = {"cabos": [{"padrao": "^C\\d+$", "valor_cabo": "C"}, {"padrao": "^A\\d+$", "valor_cabo": "A"}],
           "estruturas": [{"padrao": "^SI3$", "qtd_min": 1, "qtd_max": 2, "max_reais": 1, "amarracao": True}]}


def analisa(cabos, outros, cfg=CFG_SI3):
    return vr.analisar(cabos, outros, cfg, primeiro, primeiro)


def cab(ativo, v=None):
    d = {"entidade": "CABO", "operacao": "I", "ativo": ativo}
    if v is not None:
        d["vinculoEstruturas"] = v
    return d


def est(ativo, fict=None):
    d = {"entidade": "ESTRUTURA", "operacao": "I", "ativo": ativo}
    if fict:
        d["cabosFicticios"] = fict
    return d


def test_si3_um_real_completa_e_segundo_real_e_bloqueado():
    a = analisa([cab("C16", [0])], [est("SI3")])
    assert a["estruturas"][0]["estado"] == "gr" and a["estruturas"][0]["precisa"] == 1 and a["estruturas"][0]["maximo"] == 2
    a = analisa([cab("C16", [0]), cab("C25", [0])], [est("SI3")])
    assert a["estruturas"][0]["estado"] == "rd" and "fictício" in a["estruturas"][0]["erros"][0]
    assert any(not l["ok"] for l in a["links"])


def test_si3_segundo_cabo_ficticio_nao_vincula_o_cabo():
    a = analisa([cab("C16", [0]), cab("A35")], [est("SI3", ["A35"])])
    e = a["estruturas"][0]
    assert e["estado"] == "gr" and e["tem"] == 2 and e["ficticios"] == 1 and not e["erros"]
    assert a["cabos"][1]["tem"] == 0                                           # a SI3 NÃO conta como estrutura vinculada do cabo fictício
    assert not vr.achados([cab("C16", [0])], [est("SI3")], CFG_SI3, primeiro, primeiro)      # 1 cabo é o bastante


def test_si3_sem_cabo_continua_pendente():
    a = analisa([cab("C16")], [est("SI3")])
    assert a["estruturas"][0]["estado"] == "k"


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_js():
    js = """
    const VR = require(process.argv[1]); const p = t => (t || '').split(' ')[0].toUpperCase();
    const cfg = %s, textos = %s;
    const cab = (a, v) => { const d = { entidade: 'CABO', operacao: 'I', ativo: a }; if (v) d.vinculoEstruturas = v; return d; };
    const est = (a, f) => { const d = { entidade: 'ESTRUTURA', operacao: 'I', ativo: a }; if (f) d.cabosFicticios = f; return d; };
    const casos = [[[cab('C16', [0])], [est('SI3')]], [[cab('C16', [0]), cab('C25', [0])], [est('SI3')]], [[cab('C16', [0]), cab('A35')], [est('SI3', ['A35'])]]];
    console.log(JSON.stringify({ d: textos.map(VR.derivarCabo), a: casos.map(([c, o]) => { const r = VR.analisar(c, o, cfg, p, p); return { e: r.estruturas, c: r.cabos, l: r.links }; }),
        v: VR.validarConfig({ cabos: [{ padrao: 'A', fases: 'x', nivel: 'AT', padrao_texto: '(' }], estruturas: [{ padrao: 'SI3', qtd_min: 3, qtd_max: 2, max_reais: 0 }] }) }));
    """ % (json.dumps(CFG_SI3), json.dumps(TEXTOS))
    caminho = os.path.join(config.BASE_DIR if hasattr(config, "BASE_DIR") else ".", "static", "vinculacao_regras.js")
    saida = json.loads(subprocess.run([NODE, "-e", js, os.path.abspath(caminho)], capture_output=True, text=True, check=True).stdout)
    assert saida["d"] == [vr.derivar_cabo(t) for t in TEXTOS]
    casos = [([cab("C16", [0])], [est("SI3")]), ([cab("C16", [0]), cab("C25", [0])], [est("SI3")]), ([cab("C16", [0]), cab("A35")], [est("SI3", ["A35"])])]
    for js_r, (c, o) in zip(saida["a"], casos):
        py = vr.analisar(c, o, CFG_SI3, primeiro, primeiro)
        assert {int(k): v for k, v in js_r["e"].items()} == py["estruturas"]
        assert {int(k): v for k, v in js_r["c"].items()} == py["cabos"]
        assert js_r["l"] == py["links"]
    corta = lambda erros: [e.split(" (")[0] if "regex inválida" in e else e for e in erros]      # a mensagem do motor de regex difere entre JS e Python
    assert corta(saida["v"]) == corta(vr.validar_config({"cabos": [{"padrao": "A", "fases": "x", "nivel": "AT", "padrao_texto": "("}], "estruturas": [{"padrao": "SI3", "qtd_min": 3, "qtd_max": 2, "max_reais": 0}]}))
