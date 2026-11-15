"""TASK-169 — equipamento (RL, CFU, para-raio…) herda os cabos do poste sem ocupar pontas; `distintos`/`ordem`/`nivel`/`se_faltar`. Python = JS."""
import copy
import json
import os
import shutil
import subprocess

import pytest

import config
from services import vinculacao_regras as vr

NODE = shutil.which("node")


def est(padrao, **k):
    d = {"padrao": padrao, "lados": "LIVRE", "parametro": "", "cabos": ""}
    d.update(k)
    return d


CFG = {
    "versao": 2,
    "cabos": [
        {"padrao": "^A\\d+P$", "grupo": "MT", "valor_cabo": "P", "operacoes": ["I", "M"], "pontas": {"I": 2, "M": 1}, "estruturas": "B4MT_RL, CE1, CE23, N3", "nivel": "MT"},
        {"padrao": "^PB\\d+$", "grupo": "BT", "valor_cabo": "", "operacoes": ["I"], "estruturas": "B2", "nivel": "BT"},
    ],
    "estruturas": [
        est("^B4MT_RL$", qtd_cabos=2),
        est("^CE1$", qtd_cabos=2, lados="IGUAIS"),
        est("^CE23$", qtd_cabos=3, parametros="cabo, cabo2", distintos=True),
        est("^N3$", qtd_cabos=1, amarracao=True),
        est("^B2$", qtd_cabos=1),
        est("RL, RL_REAP", qtd_cabos=2, parametros="cabo, cabo2", ocupa=False, herda="POSTE", nivel="MT", ordem="MAIOR_PRIMEIRO", distintos=True, se_faltar="AVISAR"),
        est("CFU", qtd_cabos=2, parametros="cabo, cabo2", ocupa=False, herda="POSTE", nivel="MT", ordem="MAIOR_PRIMEIRO", distintos=True, se_faltar="BLOQUEAR"),
        est("PARA_RAIO", qtd_cabos=1, parametros="cabo", ocupa=False, herda="POSTE", nivel="MT", distintos=True),
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


def poste(x, ativo, y=0):
    return {"entidade": "POSTE", "operacao": "I", "ativo": ativo, "_x": x, "_y": y + 5, "_poste": [x, y]}


def cabo(a, b, cod="A35P", y=0, **k):
    tx = f"3x{cod}(7,94)-13,8kV 45 m" if cod.startswith("A") else f"{cod} 20 m"
    d = {"entidade": "CABO", "operacao": "I", "ativo": f"{cod} 45", "textoOriginal": tx, "_x": (a + b) / 2, "_y": y + 5, "_geo": [[a, y], [b, y]]}
    d.update(k)
    return d


def roda(cabos, outros, vincular=True):
    c, o = copy.deepcopy(cabos), copy.deepcopy(outros)
    r = vr.vincular_por_poste(c, o, CFG, primeiro, cod_est, codigos, aplicar=True) if vincular else None
    a = vr.analisar(c, o, CFG, primeiro, cod_est, codigos)
    return r, a, vr.parametros_por_estrutura(c, o, CFG, primeiro, cod_est, codigos), c


D35, D185 = "3XA35P_7.94_-13.8KV_45", "3XA185P"          # (descrição completa dos cabos, sem o comprimento)


def test_rl_herda_os_dois_cabos_e_a_estrutura_fica_verde_com_as_pontas():
    outros = [poste(0, "9100D, N3"), poste(100, "12600D, B4MT_RL, RL"), poste(200, "9100D, N3")]
    cabos = [cabo(0, 100, "A35P"), cabo(100, 200, "A185P")]
    r, a, par, c = roda(cabos, outros)
    assert a["estruturas"][1]["estado"] == "gr" and a["estruturas"][1]["tem"] == 2          # B4MT_RL com os 2 cabos
    assert c[0]["vinculoEstruturas"] == [0, 1] and c[1]["vinculoEstruturas"] == [1, 2]          # cada cabo com as 2 pontas
    rl = a["estruturas"][1.01]
    assert rl["ocupa"] is False and rl["estado"] == "gr" and rl["tem"] == 2
    p = par[1]["RL"]
    assert p["CABO"][0].startswith("3XA185P") and p["CABO2"][0].startswith("3XA35P")          # tronco (maior) em cabo, ramal em cabo2
    assert rl["le_do_poste"]["cabos"] == ["A185P", "A35P"]
    assert not any("RL" in x["mensagem"] for x in r["achados"] if x["tipo"] in ("incompleta", "sobra", "escolha"))


def test_rl_com_vinculo_manual_sem_geometria():
    outros = [poste(100, "12600D, B4MT_RL, RL")]
    cabos = [dict(cabo(0, 100, "A35P"), vinculoEstruturas=[0], vinculoOrigem="manual"), dict(cabo(100, 200, "A185P"), vinculoEstruturas=[0], vinculoOrigem="manual")]
    for x in cabos:
        x.pop("_geo")
    _, a, par, _c = roda(cabos, outros, vincular=False)
    assert a["estruturas"][0.01]["tem"] == 2 and par[0]["RL"]["CABO"][0].startswith("3XA185P")


def test_cfu_cabo_passante_so_cabo_e_derivacao_tronco_e_ramal():
    pas = [poste(100, "12600D, CE1, CFU")]
    _, a, par, _c = roda([cabo(0, 100), cabo(100, 200)], pas)
    assert list(par[0]["CFU"]) == ["CABO"] and a["estruturas"][0.01]["tem"] == 1                 # os dois lados = 1 só cabo (distintos)
    der = [poste(100, "12600D, CE23, CFU")]
    _, a2, par2, _c2 = roda([cabo(0, 100, "A185P"), cabo(100, 200, "A185P"), cabo(100, 150, "A35P")], der)
    assert par2[0]["CFU"]["CABO"][0].startswith("3XA185P") and par2[0]["CFU"]["CABO2"][0].startswith("3XA35P")
    assert par2[0]["CE23"]["CABO"][0].startswith("3XA185P") and par2[0]["CE23"]["CABO2"][0].startswith("3XA35P")          # distintos numa estrutura que ocupa


def test_para_raio_filtra_o_nivel():
    outros = [poste(100, "12600D, N3, PARA_RAIO")]
    cabos = [cabo(0, 100, "A35P"), cabo(100, 200, "PB35")]
    _, a, par, _c = roda(cabos, outros)
    assert par[0]["PARA_RAIO"]["CABO"][0].startswith("3XA35P") and a["estruturas"][0.01]["tem"] == 1


def test_sem_cabo_avisa_ou_bloqueia():
    r, a, par, _ = roda([], [poste(100, "12600D, RL"), poste(300, "12700D, CFU")])
    rl = a["estruturas"][0]
    assert rl["estado"] == "k" and any("o poste 12600D não tem cabo MT ligado" in x for x in rl["avisos"]) and 0 not in par
    cfu = a["estruturas"][1]
    assert cfu["estado"] == "rd" and cfu["erros"]
    ach = vr.achados([], [poste(100, "12600D, RL")], CFG, primeiro, cod_est, codigos)
    assert any("não tem cabo MT ligado" in x["mensagem"] or "não tem cabo MT ligado" in x["explicacao"] or "12600D" in x["mensagem"] for x in ach)


def test_equipamento_nao_recebe_vinculo_de_script():
    outros = [poste(100, "12600D, RL")]
    script = {"conflitos": "avisar", "passos": [{"modo": "proximidade", "raio_m": 500, "cabos": {"operacao": ["I"]}, "estruturas": "aceitas", "pontas": 1}]}
    r = vr.executar_script(script, [cabo(0, 100)], outros, CFG, primeiro, cod_est, codigos_fn=codigos)
    assert r["vinculos"] == []


def test_valida_config_dos_campos_novos():
    ruim = {"estruturas": [{"padrao": "X", "ocupa": "sim", "herda": "ALGO", "nivel": "AT", "ordem": "X", "se_faltar": "Y", "distintos": 1}]}
    erros = vr.validar_config(ruim)
    assert len(erros) >= 5
    assert any("herda" in e for e in erros) and any("se_faltar" in e for e in erros)
    assert vr.validar_config({"estruturas": [{"padrao": "X", "ocupa": False, "herda": "POSTE", "nivel": "MT", "ordem": "MAIOR_PRIMEIRO", "distintos": True, "se_faltar": "BLOQUEAR"}]}) == []
    assert any("só vale com" in e for e in vr.validar_config({"estruturas": [{"padrao": "X", "herda": "POSTE"}]}))


def test_secao_do_cabo():
    assert vr.secao_cabo("A185P") == 185 and vr.secao_cabo("S336") == 336 and vr.secao_cabo("S04") == 21.2 and vr.secao_cabo("S1/0") == 53.5 and vr.secao_cabo("X") == 0


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
    manual = [dict(cabo(0, 100, "A35P"), vinculoEstruturas=[0], vinculoOrigem="manual"), dict(cabo(100, 200, "A185P"), vinculoEstruturas=[0], vinculoOrigem="manual")]
    casos = [
        {"vincular": True, "outros": [poste(0, "9100D, N3"), poste(100, "12600D, B4MT_RL, RL"), poste(200, "9100D, N3")], "cabos": [cabo(0, 100, "A35P"), cabo(100, 200, "A185P")]},
        {"vincular": False, "outros": [poste(100, "12600D, B4MT_RL, RL")], "cabos": manual},
        {"vincular": True, "outros": [poste(100, "12600D, CE1, CFU")], "cabos": [cabo(0, 100), cabo(100, 200)]},
        {"vincular": True, "outros": [poste(100, "12600D, CE23, CFU")], "cabos": [cabo(0, 100, "A185P"), cabo(100, 200, "A185P"), cabo(100, 150, "A35P")]},
        {"vincular": True, "outros": [poste(100, "12600D, N3, PARA_RAIO")], "cabos": [cabo(0, 100, "A35P"), cabo(100, 200, "PB35")]},
        {"vincular": True, "outros": [poste(100, "12600D, RL"), poste(300, "12700D, CFU")], "cabos": []},
    ]
    rr = subprocess.run([NODE, "-e", _JS % {"v": json.dumps(os.path.join(config.STATIC_DIR, "vinculacao_regras.js"))}], input=json.dumps({"casos": casos, "cfg": CFG}), capture_output=True, text=True, timeout=30)
    assert rr.returncode == 0, rr.stderr
    js = json.loads(rr.stdout)
    for c, got in zip(casos, js):
        cabos, outros = copy.deepcopy(c["cabos"]), copy.deepcopy(c["outros"])
        r = vr.vincular_por_poste(cabos, outros, CFG, primeiro, cod_est, codigos, aplicar=True) if c["vincular"] else None
        a = vr.analisar(cabos, outros, CFG, primeiro, cod_est, codigos)
        assert got["r"] == r, c
        assert {float(k): v for k, v in got["a"]["estruturas"].items()} == {float(k): v for k, v in a["estruturas"].items()}, c
        assert got["a"]["cabos"] == {str(k): v for k, v in a["cabos"].items()} or {int(k): v for k, v in got["a"]["cabos"].items()} == a["cabos"], c
        assert {int(k): v for k, v in got["p"].items()} == vr.parametros_por_estrutura(cabos, outros, CFG, primeiro, cod_est, codigos), c
        assert got["ach"] == vr.achados(cabos, outros, CFG, primeiro, cod_est, codigos), c
        assert got["cabos"] == cabos, c
