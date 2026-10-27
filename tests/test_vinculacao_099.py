"""TASK-099 — script de vínculo por coordenadas com RESERVA por sorteio (modo `aleatorio`, passo `completar`). Python = JS."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import vinculacao_regras as vr

NODE = shutil.which("node")
CFG = {"cabos": [{"padrao": "^C\\d+$", "pontas": {"I": 2, "M": 1}}],
       "estruturas": [{"padrao": "^N1$", "qtd_cabos": 2, "lados": "LIVRE"}, {"padrao": "^T3$", "qtd_cabos": 1, "amarracao": True}]}
SCRIPT = json.load(open(os.path.join(config.BASE_DIR if hasattr(config, "BASE_DIR") else ".", "data", "vinculacao_script_coordenadas_com_sorteio.json"), encoding="utf-8"))


def primeiro(t):
    return (t or "").split(" ")[0].upper()


def cab(a, op="I", xy=None, v=None):
    d = {"entidade": "CABO", "operacao": op, "ativo": a}
    if xy:
        d["_x"], d["_y"] = xy
    if v is not None:
        d["vinculoEstruturas"] = v
    return d


def est(a, xy=None):
    d = {"entidade": "ESTRUTURA", "operacao": "I", "ativo": a}
    if xy:
        d["_x"], d["_y"] = xy
    return d


def cenario():
    outros = [est("N1", (0, 0)), est("N1", (10, 0)), est("N1", (500, 0)), est("T3", (1000, 0)), est("N1", (2000, 0)), est("T3", (3000, 0))]
    cabos = [cab("C16", "I", (5, 0)),            # entre as 2 primeiras: coordenadas resolvem
             cab("C25", "I"),                    # sem coordenada: sorteio
             cab("C35", "I", (505, 0)),          # só 1 estrutura no raio (500): parcial, o sorteio completa
             cab("C50", "M"),                    # M sem coordenada: sorteio de amarração
             cab("C70", "M", (1002, 0))]         # M com coordenada: amarração mais próxima
    return cabos, outros


def roda(script=SCRIPT, aplicar=True):
    cabos, outros = cenario()
    r = vr.executar_script(script, cabos, outros, CFG, primeiro, primeiro, aplicar)
    return r, cabos


def test_script_valido():
    assert vr.validar_script(SCRIPT) == []
    assert vr.validar_script({"passos": [{"modo": "aleatorio", "semente": "x"}]})
    assert vr.validar_script({"passos": [{"modo": "aleatorio", "completar": "sim"}]})


def test_coordenadas_primeiro_e_sorteio_so_para_o_que_sobrou():
    r, cabos = roda()
    assert cabos[0]["vinculoEstruturas"] == [0, 1]                 # coordenadas (passo 2)
    assert cabos[4]["vinculoEstruturas"] == [3]                    # M perto da T3 (passo 1)
    assert cabos[2]["vinculoEstruturas"][0] == 2 and len(cabos[2]["vinculoEstruturas"]) == 2          # parcial pelas coordenadas, completado pelo sorteio
    assert len(cabos[1]["vinculoEstruturas"]) == 2                  # sem coordenada: tudo no sorteio
    assert cabos[3]["vinculoEstruturas"] == [5]                    # M sem coordenada: sorteia a única amarração livre (a T3 perto já ficou com o cabo M das coordenadas)
    assert all(c.get("vinculoOrigem") == "script" for c in cabos)


def test_o_sorteio_respeita_capacidade_e_tabelas():
    r, cabos = roda()
    a = vr.analisar(cabos, cenario()[1], CFG, primeiro, primeiro)
    assert all(l["ok"] for l in a["links"])


def test_sorteio_e_reproduzivel_e_a_semente_muda_o_resultado():
    def vinculos(sem):
        s = json.loads(json.dumps(SCRIPT))
        for p in s["passos"]:
            if p["modo"] == "aleatorio":
                p["semente"] = sem
        cabos, outros = cenario()
        outros += [est("N1", (3000, 0)), est("N1", (4000, 0)), est("N1", (5000, 0))]
        vr.executar_script(s, cabos, outros, CFG, primeiro, primeiro, True)
        return [c.get("vinculoEstruturas") for c in cabos]
    assert vinculos(7) == vinculos(7)
    assert len({json.dumps(vinculos(s)) for s in range(1, 12)}) > 1


def test_completar_so_ganha_o_que_falta():
    cabos = [cab("C16", "I", None, [0])]
    outros = [est("N1"), est("N1"), est("N1")]
    s = {"versao": 1, "passos": [{"modo": "aleatorio", "completar": True, "pontas": 2}]}
    r = vr.executar_script(s, cabos, outros, CFG, primeiro, primeiro, True)
    assert cabos[0]["vinculoEstruturas"][0] == 0 and len(cabos[0]["vinculoEstruturas"]) == 2 and r["vinculos"][0]["completo"]
    cabos = [cab("C16", "I", None, [0])]
    r = vr.executar_script({"versao": 1, "passos": [{"modo": "aleatorio", "pontas": 2}]}, cabos, outros, CFG, primeiro, primeiro, True)      # sem `completar`: já tem vínculo
    assert cabos[0]["vinculoEstruturas"] == [0] and r["passos"][0]["ja_tinham"] == 1


def test_mesmo_poste_nao_vale_como_as_duas_pontas_no_sorteio():
    cabos = [cab("C16", "I")]
    outros = [est("N1"), est("N1")]
    outros = [dict(outros[0], ativo="9400D, N1, N1")]
    r = vr.executar_script({"versao": 1, "passos": [{"modo": "aleatorio", "pontas": 2, "aceitar_parcial": True}]}, cabos, outros, CFG, primeiro, lambda a, row=None: "N1",
                           True, lambda a: ["N1", "N1"])
    assert len(cabos[0]["vinculoEstruturas"]) == 1                     # duas estruturas do mesmo poste não completam as 2 pontas


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_js_incluindo_o_sorteio():
    js = """
    const VR = require(process.argv[1]); const p = t => (t || '').split(' ')[0].toUpperCase(); const cfg = %s, script = %s;
    const cab = (a, op, xy) => Object.assign({ entidade: 'CABO', operacao: op, ativo: a }, xy ? { _x: xy[0], _y: xy[1] } : {});
    const est = (a, xy) => Object.assign({ entidade: 'ESTRUTURA', operacao: 'I', ativo: a }, xy ? { _x: xy[0], _y: xy[1] } : {});
    const saida = [];
    for (const sem of [1, 2, 3, 7, 12345, -5, 2147483647]) {
        const s = JSON.parse(JSON.stringify(script)); s.passos.forEach(x => { if (x.modo === 'aleatorio') x.semente = sem; });
        const outros = [est('N1', [0, 0]), est('N1', [10, 0]), est('N1', [500, 0]), est('T3', [1000, 0]), est('N1', [2000, 0]), est('N1', [3000, 0]), est('N1', [4000, 0])];
        const cabos = [cab('C16', 'I', [5, 0]), cab('C25', 'I'), cab('C35', 'I', [505, 0]), cab('C50', 'M'), cab('C70', 'M', [1002, 0]), cab('C80', 'I')];
        const r = VR.executarScript(s, cabos, outros, cfg, p, p, true);
        saida.push({ v: cabos.map(c => c.vinculoEstruturas || null), r: r.vinculos, passos: r.passos, conf: r.conflitos, sem: r.sem_candidato });
    }
    console.log(JSON.stringify(saida));
    """ % (json.dumps(CFG), json.dumps(SCRIPT))
    s = json.loads(subprocess.run([NODE, "-e", js, os.path.join(config.STATIC_DIR, "vinculacao_regras.js")], capture_output=True, text=True, check=True).stdout)
    for sem, js_r in zip([1, 2, 3, 7, 12345, -5, 2147483647], s):
        sc = json.loads(json.dumps(SCRIPT))
        for x in sc["passos"]:
            if x["modo"] == "aleatorio":
                x["semente"] = sem
        outros = [est("N1", (0, 0)), est("N1", (10, 0)), est("N1", (500, 0)), est("T3", (1000, 0)), est("N1", (2000, 0)), est("N1", (3000, 0)), est("N1", (4000, 0))]
        cabos = [cab("C16", "I", (5, 0)), cab("C25", "I"), cab("C35", "I", (505, 0)), cab("C50", "M"), cab("C70", "M", (1002, 0)), cab("C80", "I")]
        r = vr.executar_script(sc, cabos, outros, CFG, primeiro, primeiro, True)
        assert js_r["v"] == [c.get("vinculoEstruturas") for c in cabos], sem
        assert js_r["r"] == r["vinculos"] and js_r["passos"] == r["passos"] and js_r["conf"] == r["conflitos"] and js_r["sem"] == r["sem_candidato"], sem
