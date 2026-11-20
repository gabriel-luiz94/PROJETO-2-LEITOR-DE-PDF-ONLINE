"""TASK-170 — nível do cabo fictício, `lados: DERIVACAO`, `cabos_vaos` e `define` (parâmetros automáticos pela linha do poste / 1º cabo). Python = JS."""
import copy
import json
import os
import shutil
import subprocess

import pytest

import config
from services import anotacao_outros as ao
from services import vinculacao_regras as vr

NODE = shutil.which("node")
RL = {"padrao": "RL, RL_MT, RL_REAP", "qtd_cabos": 2, "parametros": "cabo, cabo2", "ocupa": False, "herda": "POSTE", "nivel": "MT", "ordem": "MAIOR_PRIMEIRO", "distintos": True,
      "lados": "LIVRE", "parametro": "", "cabos": ""}
LINHA_RL = "RL, RL_MT, RL_REAP"
CFG = {
    "versao": 2,
    "cabos": [
        {"padrao": "^S04$", "grupo": "BT nua", "valor_cabo": "C", "nivel": "BT", "estruturas": "LN3, IBIT"},
        {"padrao": "^S04$", "grupo": "CAA MT", "valor_cabo": "A", "nivel": "MT", "estruturas": "N3, B4MT_RL, RL"},
        {"padrao": "^RAMA25$", "grupo": "Multiplex", "valor_cabo": "", "nivel": "BT", "estruturas": "FLBIM"},
        {"padrao": "^(C\\d+|A\\d+P|(S|W)\\d+(/\\d+)?)$", "grupo": "MT", "valor_cabo": "P", "operacoes": ["I", "M"], "pontas": {"I": 2, "M": 1}},
        {"padrao": "^PB\\d+$", "grupo": "Multiplex BT", "valor_cabo": "", "nivel": "BT", "operacoes": ["I"]},
    ],
    "estruturas": [
        {"padrao": "^B4MT_RL$", "qtd_cabos": 2, "lados": "LIVRE", "parametro": "", "cabos": ""},
        {"padrao": "^N3$", "qtd_cabos": 1, "amarracao": True, "lados": "LIVRE", "parametro": "", "cabos": ""},
        {"padrao": "^LN3$", "qtd_cabos": 1, "amarracao": True, "lados": "LIVRE", "parametro": "", "cabos": ""},
        {"padrao": "^FLBIM$", "qtd_cabos": 1, "amarracao": True, "lados": "LIVRE", "parametro": "", "cabos": ""},
        {"padrao": "^SDBIT$", "qtd_min": 1, "qtd_max": 3, "lados": "DERIVACAO", "parametros": "cabobt", "amarracao": True, "cabos": "", "parametro": ""},
        {"padrao": "^IBIT$", "cabos_vaos": ["^PB\\d+$", "^(C\\d+|(S|W)\\d+(/\\d+)?)$"], "amarracao": True, "lados": "LIVRE", "parametros": "cabo, cabo2", "parametro": ""},
        dict(RL),
        {"padrao": "PARA_RAIO", "qtd_cabos": 2, "parametros": "cabo, cabo2", "ocupa": False, "herda": "POSTE", "nivel": "MT", "ordem": "MAIOR_PRIMEIRO", "distintos": True,
         "lados": "LIVRE", "parametro": "", "cabos": "", "define": [{"linha_tem": LINHA_RL, "param": "tipo=RL"}]},
        {"padrao": "ATRINT_EQ, ATRINT-EQ", "ocupa": False, "herda": "POSTE", "nivel": "MT", "cabos": "^(C\\d+|A\\d+P|(S|W)\\d+(/\\d+)?)$", "ordem": "MAIOR_PRIMEIRO", "distintos": True,
         "parametros": "", "lados": "LIVRE", "parametro": "",
         "define": [{"linha_tem": LINHA_RL, "cabo": "^C120$", "param": "tipo=RL_CU120"}, {"linha_tem": LINHA_RL, "cabo": "^C\\d+$", "param": "tipo=RL_CU"},
                    {"linha_tem": LINHA_RL, "cabo": "^S4/0$", "param": "tipo=RL_CAA40"}, {"linha_tem": LINHA_RL, "cabo": "^(S|W)\\d+(/\\d+)?$", "param": "tipo=RL_CAA"},
                    {"linha_tem": LINHA_RL, "param": "tipo=RL"}]},
    ],
    "ordem_maior": ["C", "P", "A"],
}


def primeiro(txt):
    return (txt or "").replace(",", " ").split()[0].upper() if (txt or "").strip() else ""


def codigos(ativo):
    return [x.strip().upper() for x in (ativo or "").split(",")][1:]


def cod_est(ativo, row=None):
    c = codigos(ativo)
    return c[0] if c else primeiro(ativo)


def poste(x, ativo, y=0, **k):
    d = {"entidade": "POSTE", "operacao": "I", "ativo": ativo, "_x": x, "_y": y + 5, "_poste": [x, y]}
    d.update(k)
    return d


def cabo(cod, a=None, b=None, vinc=None, **k):
    mt = not cod.startswith(("PB", "RAMA")) and cod != "S04" or cod == "S04"
    tx = f"3x{cod}-13,8kV 45 m" if (mt and cod != "S04") else f"{cod} 20 m"
    d = {"entidade": "CABO", "operacao": "I", "ativo": f"{cod} 45", "textoOriginal": tx}
    if a is not None:
        d.update({"_geo": [[a, 0], [b, 0]], "_x": (a + b) / 2, "_y": 5})
    if vinc is not None:
        d["vinculoEstruturas"] = vinc
    d.update(k)
    return d


def analisa(cabos, outros, cfg=CFG):
    a = vr.analisar(cabos, outros, cfg, primeiro, cod_est, codigos)
    return a, vr.parametros_por_estrutura(cabos, outros, cfg, primeiro, cod_est, codigos)


# ── A. nível do cabo fictício ──
def test_ficticio_s04_so_entra_no_rl_se_for_cabo_mt_da_estrutura():
    # N3 com cabo real A35P + fictício S04 (regra "CAA MT") → entra no RL (MT)
    a, par = analisa([cabo("A35P", 0, 100, [0])], [poste(100, "12600D, N3, RL", **{"cabosFicticios": ["S04"]})])
    assert "S04" in a["estruturas"][0.01]["le_do_poste"]["cabos"]
    # o mesmo S04 fictício numa LN3 (regra "BT nua") é BT → não entra
    b, _ = analisa([cabo("A35P", 0, 100, [0])], [poste(100, "12600D, LN3, RL", **{"cabosFicticios": ["S04"]})])
    assert "S04" not in b["estruturas"][0.01]["le_do_poste"]["cabos"]


def test_ficticio_rama25_na_flbim_e_bt():
    a, par = analisa([cabo("A35P", 0, 100, [1])], [poste(100, "12600D, N3, FLBIM, RL", **{"cabosFicticios": [], "cabosFicticiosSlots": {"1": ["RAMA25"]}})])
    rl = a["estruturas"][0.02]
    assert rl["le_do_poste"]["cabos"] == ["A35P"]                                          # o RAMA25 fictício da FLBIM é BT: o RL (só MT) não o lê


def test_nivel_do_ficticio_helper():
    assert vr.info_cabo_ficticio("S04", "LN3", "I", vr.normalizar_config(CFG))["nivel"] == "BT"
    assert vr.info_cabo_ficticio("S04", "N3", "I", vr.normalizar_config(CFG))["nivel"] == "MT"
    assert vr.info_cabo_ficticio("A35P", "N3", "I", vr.normalizar_config(CFG))["nivel"] is None          # regra sem nível: como antes


# ── B. lados DERIVACAO ──
@pytest.mark.parametrize("ordem", [[0, 1, 2], [1, 0, 2], [2, 0, 1]])
def test_derivacao_tronco_igual_e_derivacao_em_qualquer_ordem_de_ligacao(ordem):
    todos = [("PB70", 0, 100), ("PB70", 100, 200), ("PB35", 100, 150)]
    cabos = [cabo(todos[i][0], todos[i][1], todos[i][2], [0]) for i in ordem]
    a, par = analisa(cabos, [poste(100, "11400D, SDBIT")])
    assert a["estruturas"][0]["estado"] == "gr" and par[0]["SDBIT"]["CABOBT"] == ["PB35"]


def test_derivacao_tronco_diferente_fica_vermelho():
    a, _ = analisa([cabo("PB70", 0, 100, [0]), cabo("PB120", 100, 200, [0]), cabo("PB35", 100, 150, [0])], [poste(100, "11400D, SDBIT")])
    assert a["estruturas"][0]["estado"] == "rd" and any("tronco precisam ser iguais" in e for e in a["estruturas"][0]["erros"])


def test_derivacao_so_tronco_falta_a_derivacao_e_dois_diferentes_avisa():
    a, par = analisa([cabo("PB70", 0, 100, [0]), cabo("PB70", 100, 200, [0])], [poste(100, "11400D, SDBIT")])
    assert a["estruturas"][0]["estado"] == "pu" and any("falta o cabo da derivação" in m for m in a["estruturas"][0]["avisos_vinculo"]) and 0 not in par
    b, parb = analisa([cabo("PB70", 0, 100, [0]), cabo("PB35", 100, 200, [0])], [poste(100, "11400D, SDBIT")])
    assert b["estruturas"][0]["estado"] == "gr" and parb[0]["SDBIT"]["CABOBT"] == ["PB70"] and any("usado o 1º ligado" in m for m in b["estruturas"][0]["avisos_vinculo"])


def test_vincular_por_poste_nao_liga_segundo_tronco_diferente():
    outros = [poste(100, "11400D, SDBIT")]
    cabos = [cabo("PB70", 0, 100), cabo("PB120", 100, 200), cabo("PB35", 100, 150)]
    r = vr.vincular_por_poste(cabos, outros, CFG, primeiro, cod_est, codigos, aplicar=True)
    ligados = [c["vinculoEstruturas"] for c in cabos if c["vinculoEstruturas"]]
    a, _ = analisa(cabos, outros)
    assert a["estruturas"][0]["estado"] != "rd" and len(ligados) == 2          # o 3º cabo (que faria o tronco ser diferente) ficou de fora


# ── C. cabos por vão ──
def test_ibit_um_multiplex_e_uma_bt_nua():
    a, par = analisa([cabo("S04", 0, 100, [0]), cabo("PB35", 100, 200, [0])], [poste(100, "11400D, IBIT")])
    assert a["estruturas"][0]["estado"] == "gr" and a["estruturas"][0]["tem"] == 2
    assert par[0]["IBIT"]["CABO"] == ["PB35"] and par[0]["IBIT"]["CABO2"] == ["S04"]          # a ordem dos parâmetros segue a dos vãos (multiplex primeiro)
    b, _ = analisa([cabo("PB35", 0, 100, [0]), cabo("PB35", 100, 200, [0])], [poste(100, "11400D, IBIT")])
    assert b["estruturas"][0]["estado"] == "rd" and any("cada vão aceita um cabo" in e for e in b["estruturas"][0]["erros"])


def test_ibit_vincular_por_poste_recusa_o_segundo_multiplex():
    outros = [poste(100, "11400D, IBIT")]
    cabos = [cabo("PB35", 0, 100), cabo("PB35", 100, 200), cabo("S04", 100, 150)]
    vr.vincular_por_poste(cabos, outros, CFG, primeiro, cod_est, codigos, aplicar=True)
    assert [c["vinculoEstruturas"] for c in cabos] == [[0], [], [0]]


# ── D. define ──
def poste_rl(ativo_extra, cabos_poste):
    return [poste(100, f"12600D, B4MT_RL, RL, PARA_RAIO, ATRINT_EQ{ativo_extra}")], cabos_poste


@pytest.mark.parametrize("cods,esperado", [(["A185P", "A35P"], "RL"), (["C35"], "RL_CU"), (["C120", "C35"], "RL_CU120"), (["S1/0"], "RL_CAA"), (["S4/0"], "RL_CAA40")])
def test_atrint_eq_variante_pelo_cabo_mt_do_poste(cods, esperado):
    cabos = [cabo(c, i * 100, (i + 1) * 100, [0]) for i, c in enumerate(cods)]
    a, par = analisa(cabos, [poste(100, "12600D, B4MT_RL, RL, PARA_RAIO, ATRINT_EQ")])
    assert par[0]["ATRINT_EQ"]["TIPO"] == [esperado] and a["estruturas"][0.03]["define"]["TIPO"]["valor"] == esperado
    assert par[0]["PARA_RAIO"]["TIPO"] == ["RL"] and a["estruturas"][0.02]["define"]["TIPO"]["origem"] == "linha: RL"


def test_define_origem_mostra_o_cabo():
    a, _ = analisa([cabo("C35", 0, 100, [0])], [poste(100, "12600D, B4MT_RL, RL, ATRINT_EQ")])
    assert a["estruturas"][0.02]["define"]["TIPO"]["origem"] == "cabo C35"


def test_sem_rl_na_linha_nada_e_definido_e_o_motivo_fica_registrado():
    a, par = analisa([cabo("A35P", 0, 100, [0])], [poste(100, "12600D, N3, PARA_RAIO, ATRINT_EQ")])
    pr, at = a["estruturas"][0.01], a["estruturas"][0.02]
    assert "TIPO" not in (par.get(0, {}).get("PARA_RAIO") or {}) and pr["define"] == {} and pr["define_faltou"] == [{"param": "TIPO", "motivo": "a linha não tem RL"}]
    assert at["define_faltou"][0]["param"] == "TIPO"


def test_define_nao_sobrescreve_o_vinculo_nem_o_digitado():
    cfg = copy.deepcopy(CFG)
    for r in cfg["estruturas"]:
        if r["padrao"] == "PARA_RAIO":
            r["parametros"] = "tipo=ativo"
    a, par = analisa([cabo("A35P", 0, 100, [0])], [poste(100, "12600D, B4MT_RL, RL, PARA_RAIO")], cfg)
    assert par[0]["PARA_RAIO"]["TIPO"] == ["A35P"]                                           # o vínculo de cabo já deu `tipo`: o define não sobrescreve
    assert ao._com_param("PARA_RAIO(TIPO=PADRAO)", "TIPO", "RL", {"PARA_RAIO": {"nomes": ["TIPO"]}}) == "PARA_RAIO(TIPO=PADRAO)"          # o digitado vence


def test_validar_config_dos_campos_novos():
    erros = vr.validar_config({"estruturas": [{"padrao": "X", "lados": "XPTO", "cabos_vaos": "a", "define": [{"param": "tipo"}, {"linha_tem": "RL", "param": "tipo=RL"}, 5]}]})
    assert any("lados" in e for e in erros) and any("cabos_vaos" in e for e in erros) and sum("define" in e for e in erros) >= 3
    assert vr.validar_config({"estruturas": [{"padrao": "X", "lados": "DERIVACAO", "cabos_vaos": ["^A$"], "define": [{"linha_tem": "RL, RL_MT", "cabo": "^C\\d+$", "param": "tipo=RL_CU"}]}]}) == []


_JS = r"""
const v = require(%(v)s); const e = JSON.parse(require('fs').readFileSync(0, 'utf8'));
const primeiro = t => (String(t || '').trim() ? String(t).replace(/,/g, ' ').trim().split(/\s+/)[0].toUpperCase() : '');
const codigos = a => String(a || '').split(',').map(x => x.trim().toUpperCase()).slice(1);
const codEst = (a) => { const c = codigos(a); return c.length ? c[0] : primeiro(a); };
console.log(JSON.stringify(e.casos.map(c => { const cabos = JSON.parse(JSON.stringify(c.cabos)), outros = JSON.parse(JSON.stringify(c.outros));
  const r = c.vincular ? v.vincularPorPoste(cabos, outros, e.cfg, primeiro, codEst, codigos, undefined, true) : null;
  return { r, a: v.analisar(cabos, outros, e.cfg, primeiro, codEst, codigos), p: v.parametrosPorEstrutura(cabos, outros, e.cfg, primeiro, codEst, codigos), ach: v.achados(cabos, outros, e.cfg, primeiro, codEst, codigos), cabos }; })));
"""


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_com_o_js():
    casos = [
        {"vincular": False, "cabos": [cabo("A35P", 0, 100, [0])], "outros": [poste(100, "12600D, N3, RL", **{"cabosFicticios": ["S04"]})]},
        {"vincular": False, "cabos": [cabo("A35P", 0, 100, [0])], "outros": [poste(100, "12600D, LN3, RL", **{"cabosFicticios": ["S04"]})]},
        {"vincular": False, "cabos": [cabo("PB70", 0, 100, [0]), cabo("PB70", 100, 200, [0]), cabo("PB35", 100, 150, [0])], "outros": [poste(100, "11400D, SDBIT")]},
        {"vincular": False, "cabos": [cabo("PB70", 0, 100, [0]), cabo("PB120", 100, 200, [0]), cabo("PB35", 100, 150, [0])], "outros": [poste(100, "11400D, SDBIT")]},
        {"vincular": False, "cabos": [cabo("PB70", 0, 100, [0]), cabo("PB70", 100, 200, [0])], "outros": [poste(100, "11400D, SDBIT")]},
        {"vincular": False, "cabos": [cabo("PB70", 0, 100, [0]), cabo("PB35", 100, 200, [0])], "outros": [poste(100, "11400D, SDBIT")]},
        {"vincular": True, "cabos": [cabo("PB70", 0, 100), cabo("PB120", 100, 200), cabo("PB35", 100, 150)], "outros": [poste(100, "11400D, SDBIT")]},
        {"vincular": False, "cabos": [cabo("S04", 0, 100, [0]), cabo("PB35", 100, 200, [0])], "outros": [poste(100, "11400D, IBIT")]},
        {"vincular": False, "cabos": [cabo("PB35", 0, 100, [0]), cabo("PB35", 100, 200, [0])], "outros": [poste(100, "11400D, IBIT")]},
        {"vincular": True, "cabos": [cabo("PB35", 0, 100), cabo("PB35", 100, 200), cabo("S04", 100, 150)], "outros": [poste(100, "11400D, IBIT")]},
        {"vincular": False, "cabos": [cabo("C120", 0, 100, [0]), cabo("C35", 100, 200, [0])], "outros": [poste(100, "12600D, B4MT_RL, RL, PARA_RAIO, ATRINT_EQ")]},
        {"vincular": False, "cabos": [cabo("S4/0", 0, 100, [0])], "outros": [poste(100, "12600D, B4MT_RL, RL, PARA_RAIO, ATRINT_EQ")]},
        {"vincular": False, "cabos": [cabo("A35P", 0, 100, [0])], "outros": [poste(100, "12600D, N3, PARA_RAIO, ATRINT_EQ")]},
        # Equipamento sem cabo vinculado na linha encontra pela geometria o cabo que chega ao poste.
        {"vincular": False, "cabos": [cabo("A35P", 0, 99), cabo("C35", 108, 200)],
         "outros": [poste(100, "12600D, RL"), poste(200, "12600D, N3")]},
    ]
    rr = subprocess.run([NODE, "-e", _JS % {"v": json.dumps(os.path.join(config.STATIC_DIR, "vinculacao_regras.js"))}], input=json.dumps({"casos": casos, "cfg": CFG}), capture_output=True, text=True, timeout=30)
    assert rr.returncode == 0, rr.stderr
    js = json.loads(rr.stdout)
    assert js[-1]["a"]["estruturas"]["0"]["le_do_poste"]["cabos"] == ["A35P"]
    for c, got in zip(casos, js):
        cabos, outros = copy.deepcopy(c["cabos"]), copy.deepcopy(c["outros"])
        r = vr.vincular_por_poste(cabos, outros, CFG, primeiro, cod_est, codigos, aplicar=True) if c["vincular"] else None
        a = vr.analisar(cabos, outros, CFG, primeiro, cod_est, codigos)
        assert got["r"] == r, c
        assert {float(k): v for k, v in got["a"]["estruturas"].items()} == {float(k): v for k, v in a["estruturas"].items()}, c
        assert {int(k): v for k, v in got["p"].items()} == vr.parametros_por_estrutura(cabos, outros, CFG, primeiro, cod_est, codigos), c
        assert got["ach"] == vr.achados(cabos, outros, CFG, primeiro, cod_est, codigos), c
        assert got["cabos"] == cabos, c
