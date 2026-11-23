"""TASK-234: cabo M — o vínculo automático liga 1 ponta; a 2ª é sempre do usuário (limite manual `pontasMax`, padrão M = 2)."""
import copy
import json
import os
import shutil
import subprocess

import pytest

import config
from services import vinculacao_regras as vr

NODE = shutil.which("node")
CFG = {
    "versao": 2,
    "cabos": [{"padrao": "^MT\\d+$", "grupo": "MT", "valor_cabo": "P", "operacoes": ["I", "M"], "estruturas": "^(N3)$"},
              {"padrao": "^MX\\d+$", "grupo": "MT", "valor_cabo": "P", "pontas": {"M": 1}, "pontas_max": {"I": 3, "M": 1}}],
    "estruturas": [{"padrao": "^N3$", "qtd_cabos": 1, "lados": "LIVRE", "amarracao": True, "cabos": "^(MT|MX)\\d+$", "parametro": ""}],
    "ordem_maior": ["C", "P", "A"],
}


def primeiro(txt):
    return (txt or "").replace(",", " ").split()[0].upper() if (txt or "").strip() else ""


def codigos(ativo):
    return [x.strip().upper() for x in (ativo or "").split(",")][1:]


def cod_est(ativo, row=None):
    c = codigos(ativo)
    return c[0] if c else primeiro(ativo)


OUTROS = [{"entidade": "POSTE", "operacao": "I", "ativo": "9100D, N3"}, {"entidade": "POSTE", "operacao": "I", "ativo": "9100D, N3"}]


def test_limites_de_pontas_python():
    cfg = vr.normalizar_config(CFG)
    m, i = vr.info_cabo("MT1", "M", cfg), vr.info_cabo("MT1", "I", cfg)
    assert (m["pontas"], m["pontas_max"]) == (1, 2)               # automático 1, manual até 2
    assert (i["pontas"], i["pontas_max"]) == (2, 2)
    assert vr.info_cabo("MT1", "*M", cfg)["pontas_max"] == 2
    assert vr.info_cabo("MX1", "M", cfg)["pontas_max"] == 1        # `pontas_max` da Tabela de Cabos manda
    assert vr.info_cabo("MX1", "I", cfg)["pontas_max"] == 3
    assert vr.validar_config({**CFG, "cabos": [{"padrao": "X", "pontas_max": {"M": -1}}]})


def test_automatico_liga_so_uma_ponta_do_cabo_m():
    cabo = {"entidade": "CABO", "operacao": "M", "ativo": "MT1 40", "vinculoPostes": [0, 1]}
    c = [copy.deepcopy(cabo)]
    vr.vincular_por_poste(c, copy.deepcopy(OUTROS), CFG, primeiro, cod_est, codigos, aplicar=True)
    assert len(c[0]["vinculoEstruturas"]) == 1                     # a outra ponta fica para o usuário


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_js():
    js = """
const v = require(%s);
const cfg = v.normalizarConfig(%s);
const p = s => (s || '').replace(/,/g, ' ').trim().split(/\\s+/)[0].toUpperCase();
const cods = a => (a || '').split(',').slice(1).map(x => x.trim().toUpperCase());
const est = a => cods(a)[0] || p(a);
const cabos = [{entidade: 'CABO', operacao: 'M', ativo: 'MT1 40', vinculoPostes: [0, 1]}];
v.vincularPorPoste(cabos, %s, cfg, p, est, cods, undefined, true);
console.log(JSON.stringify({m: [v.infoCabo('MT1', 'M', cfg).pontas, v.infoCabo('MT1', 'M', cfg).pontasMax], i: v.infoCabo('MT1', 'I', cfg).pontasMax,
  mx: [v.infoCabo('MX1', 'M', cfg).pontasMax, v.infoCabo('MX1', 'I', cfg).pontasMax], auto: cabos[0].vinculoEstruturas.length,
  erro: v.validarConfig(Object.assign({}, cfg, {cabos: [{padrao: 'X', pontas_max: {M: -1}}]})).length > 0}));
""" % (json.dumps(os.path.join(config.STATIC_DIR, "vinculacao_regras.js")), json.dumps(CFG), json.dumps(OUTROS))
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    assert json.loads(r.stdout) == {"m": [1, 2], "i": 2, "mx": [1, 3], "auto": 1, "erro": True}


def test_tela_usa_o_limite_manual():
    ui = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    motivo = ui[ui.index("function motivoParaLigar"):ui.index("function ligar(")]
    assert "info.pontasMax" in motivo and "info.pontas)" not in motivo
    assert "{ campo: 'pontas_max'" in ui
    # o automático do mapa continua contando só as pontas automáticas
    assert "infoCabo(ic).pontas - ligacoes(ic).length" in ui
