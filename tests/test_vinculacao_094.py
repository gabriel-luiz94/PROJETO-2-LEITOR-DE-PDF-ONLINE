"""TASK-094 — várias estruturas numa linha de Outros (`9400D, N1, N4`): cada uma é uma unidade vinculável (id = linha, linha + 0,01, …). Python = JS."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import vinculacao_regras as vr

NODE = shutil.which("node")
CFG = {"cabos": [{"padrao": "^C\\d+$", "valor_cabo": "C"}, {"padrao": "^A\\d+$", "valor_cabo": "A"}],
       "estruturas": [{"padrao": "^N1$", "qtd_cabos": 2, "lados": "LIVRE"}, {"padrao": "^N4$", "qtd_cabos": 1, "lados": "LIVRE"}], "ordem_maior": ["C", "P", "A"]}


def primeiro(txt):
    return (txt or "").split(" ")[0].upper()


def codigos(ativo):
    """Itens depois do poste: `9400D, N1, N4` → [N1, N4]; linha só com o poste → []."""
    partes = [p.strip().upper() for p in (ativo or "").split(",")]
    return [p for p in partes[1:] if p]


def codigo(ativo, row=None):
    c = codigos(ativo)
    return c[0] if c else (ativo or "").split(",")[0].strip().upper()


def cab(a, v=None, xy=None):
    d = {"entidade": "CABO", "operacao": "I", "ativo": a}
    if v is not None:
        d["vinculoEstruturas"] = v
    if xy:
        d["_x"], d["_y"] = xy
    return d


def est(a, xy=None, **kw):
    d = {"entidade": "ESTRUTURA", "operacao": "I", "ativo": a, **kw}
    if xy:
        d["_x"], d["_y"] = xy
    return d


def analisa(cabos, outros):
    return vr.analisar(cabos, outros, CFG, primeiro, codigo, codigos)


def test_ids_de_unidade():
    assert vr.id_unidade(12, 0) == 12 and vr.id_unidade(12, 2) == 12.02
    assert vr.linha_de(12.02) == 12 and vr.slot_de(12.02) == 2 and vr.linha_de(7) == 7 and vr.slot_de(7) == 0
    assert vr.estruturas_disponiveis([est("9400D, N1, N4"), {"entidade": "CABO"}, est("9200D, N1")], codigo, codigos) == [0, 0.01, 2]
    assert vr.estruturas_disponiveis([est("9400D, N1, N4")]) == [0]                   # sem codigos_fn: como antes (uma por linha)


def test_cada_estrutura_da_linha_tem_o_seu_estado():
    a = analisa([cab("C16", [0]), cab("C25", [0]), cab("A35", [0.01])], [est("9400D, N1, N4")])
    assert set(a["estruturas"]) == {0, 0.01}
    n1, n4 = a["estruturas"][0], a["estruturas"][0.01]
    assert (n1["codigo"], n1["tem"], n1["estado"]) == ("N1", 2, "gr")
    assert (n4["codigo"], n4["tem"], n4["estado"]) == ("N4", 1, "gr")
    assert {(l["cabo"], l["estrutura"]) for l in a["links"]} == {(0, 0), (1, 0), (2, 0.01)}


def test_a_n4_nao_aceita_segundo_cabo_e_ficticio_e_por_estrutura():
    a = analisa([cab("C16", [0.01]), cab("C25", [0.01])], [est("9400D, N1, N4")])
    assert a["estruturas"][0.01]["estado"] == "rd"                                    # N4 aceita 1 cabo
    assert a["estruturas"][0]["estado"] == "k"                                        # a N1 da mesma linha segue sem cabo
    a = analisa([cab("C16", [0])], [est("9400D, N1, N4", cabosFicticios=["C25"], cabosFicticiosSlots={"1": ["A35"]})])
    assert a["estruturas"][0]["ficticios"] == 1 and a["estruturas"][0]["tem"] == 2      # fictício da 1ª estrutura
    assert a["estruturas"][0.01]["ficticios"] == 1 and a["estruturas"][0.01]["tem"] == 1


def test_parametro_por_linha_e_achados_usam_a_linha():
    outros = [est("9200D, N1"), est("9400D, N1, N4")]
    cabos = [cab("C16", [1.01])]
    assert vr.parametros_por_linha(cabos, outros, CFG, primeiro, codigo, codigos) == {1: {"CABO": "C"}}
    ach = vr.achados([cab("C16", [1])], outros, CFG, primeiro, codigo, codigos)
    assert any(x["linha_id"] == "OUTROS-1" and "[N4]" in x["mensagem"] for x in ach)           # N4 sem cabo: mensagem diz qual estrutura
    assert all(x["linha_id"] in ("OUTROS-0", "OUTROS-1") for x in ach)


def test_script_nao_usa_duas_estruturas_do_mesmo_poste_como_as_duas_pontas():
    cfg = {"cabos": [{"padrao": "^C\\d+$", "pontas": {"I": 2}}], "estruturas": [{"padrao": "^N[14]$", "qtd_cabos": 2}]}
    outros = [est("9400D, N1, N4", xy=(10, 0)), est("9500D, N1", xy=(30, 0))]
    cabos = [cab("C16", xy=(15, 0))]
    script = {"versao": 1, "passos": [{"modo": "proximidade", "pontas": 2}]}
    r = vr.executar_script(script, cabos, outros, cfg, primeiro, codigo, True, codigos)
    assert r["vinculos"][0]["estruturas"] == [0, 1]          # N1 do 1º poste + N1 do 2º (a N4 do 1º poste foi pulada)
    assert cabos[0]["vinculoEstruturas"] == [0, 1] and cabos[0]["vinculoOrigem"] == "script"


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_js():
    js = """
    const VR = require(process.argv[1]); const p = t => (t || '').split(' ')[0].toUpperCase();
    const cods = a => (a || '').split(',').map(x => x.trim().toUpperCase()).slice(1).filter(Boolean);
    const cod = (a, row) => { const c = cods(a); return c.length ? c[0] : (a || '').split(',')[0].trim().toUpperCase(); };
    const cfg = %s;
    const cab = (a, v, xy) => { const d = { entidade: 'CABO', operacao: 'I', ativo: a }; if (v) d.vinculoEstruturas = v; if (xy) { d._x = xy[0]; d._y = xy[1]; } return d; };
    const est = (a, xy, extra) => Object.assign({ entidade: 'ESTRUTURA', operacao: 'I', ativo: a }, xy ? { _x: xy[0], _y: xy[1] } : {}, extra || {});
    const o1 = [est('9400D, N1, N4', null, { cabosFicticios: ['C25'], cabosFicticiosSlots: { 1: ['A35'] } }), est('9200D, N1')];
    const c1 = [cab('C16', [0, 0.01]), cab('C25', [0]), cab('A35', [1])];
    const o2 = [est('9400D, N1, N4', [10, 0]), est('9500D, N1', [30, 0])], c2 = [cab('C16', null, [15, 0])];
    const cfg2 = { cabos: [{ padrao: '^C\\\\d+$', pontas: { I: 2 } }], estruturas: [{ padrao: '^N[14]$', qtd_cabos: 2 }] };
    const a = VR.analisar(c1, o1, cfg, p, cod, cods);
    const sc = VR.executarScript({ versao: 1, passos: [{ modo: 'proximidade', pontas: 2 }] }, c2, o2, cfg2, p, cod, true, cods);
    console.log(JSON.stringify({ a: { e: a.estruturas, c: a.cabos, l: a.links }, par: VR.parametrosPorLinha(c1, o1, cfg, p, cod, cods), ach: VR.achados(c1, o1, cfg, p, cod, cods), sc: sc.vinculos, v: c2[0].vinculoEstruturas,
        disp: VR.estruturasDisponiveis(o1, cod, cods) }));
    """ % json.dumps(CFG)
    caminho = os.path.join(config.STATIC_DIR, "vinculacao_regras.js")
    saida = json.loads(subprocess.run([NODE, "-e", js, caminho], capture_output=True, text=True, check=True).stdout)
    o1 = [est("9400D, N1, N4", cabosFicticios=["C25"], cabosFicticiosSlots={"1": ["A35"]}), est("9200D, N1")]
    c1 = [cab("C16", [0, 0.01]), cab("C25", [0]), cab("A35", [1])]
    py = vr.analisar(c1, o1, CFG, primeiro, codigo, codigos)
    assert {float(k): v for k, v in saida["a"]["e"].items()} == py["estruturas"]
    assert {int(k): v for k, v in saida["a"]["c"].items()} == py["cabos"]
    assert saida["a"]["l"] == py["links"]
    assert {int(k): v for k, v in saida["par"].items()} == vr.parametros_por_linha(c1, o1, CFG, primeiro, codigo, codigos)
    assert saida["ach"] == vr.achados(c1, o1, CFG, primeiro, codigo, codigos)
    assert saida["disp"] == vr.estruturas_disponiveis(o1, codigo, codigos)
    assert saida["sc"][0]["estruturas"] == [0, 1] and saida["v"] == [0, 1]
