"""TASK-247 — R e *R não geram aviso de vinculação nem de parâmetros (C1-PARAM) e não entram nas cores de parâmetros;
fictícia (F) também não avisa vínculo. Só I, *I, M e *M (e vazio = I) avisam."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import vinculacao_regras as vr
from services.validacao_planilhas import validar_base
from tests.test_validacao_planilhas import BASE_PARAM
from tests.test_vinculacao_regras import CFG, primeiro

NODE = shutil.which("node")


def _est(ativo, op):
    return {"entidade": "ESTRUTURA", "operacao": op, "ativo": ativo}


OUTROS = [_est("N1", op) for op in ("I", "*I", "M", "*M", "", "R", "*R", "F")]


def test_so_i_e_m_avisam_vinculo():
    linhas = sorted({int(a["linha_id"].split("-")[1]) for a in vr.achados([], OUTROS, CFG, primeiro, primeiro)})
    assert linhas == [0, 1, 2, 3, 4]                          # I, *I, M, *M e vazio; R, *R e F (fictícia) não
    assert [vr.avisa_vinculo(r) for r in OUTROS] == [True] * 5 + [False] * 3


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_js_igual_ao_python():
    js = """const V = require(process.argv[1]); const o = JSON.parse(process.argv[2]), c = JSON.parse(process.argv[3]);
const f = t => String(t == null ? '' : t).split(' ')[0].toUpperCase();
process.stdout.write(JSON.stringify({ x: V.achados([], o, c, f, f), a: o.map(V.avisaVinculo) }));"""
    r = subprocess.run([NODE, "-e", js, os.path.join(config.STATIC_DIR, "vinculacao_regras.js"), json.dumps(OUTROS), json.dumps(CFG)], capture_output=True, text=True, check=True)
    out = json.loads(r.stdout)
    assert out["x"] == vr.achados([], OUTROS, CFG, primeiro, primeiro) and out["a"] == [vr.avisa_vinculo(x) for x in OUTROS]


def test_c1_param_nao_avisa_removendo():
    def regras(op):
        return [x["regra_id"] for x in validar_base([], [{"id": "L1", "ativo": "1-IT_2(XX,35)", "operacao": op}], "PERNAMBUCO", [dict(r) for r in BASE_PARAM])]
    assert regras("I") == ["C1-PARAM"] and regras("M") == ["C1-PARAM"]
    assert regras("R") == [] and regras("*R") == []
    # o nome inexistente continua avisando (C1-BASE), mesmo em retirada
    assert [x["regra_id"] for x in validar_base([], [{"id": "L1", "ativo": "1-NAOEXISTE", "operacao": "R"}], "PERNAMBUCO", [dict(r) for r in BASE_PARAM])] == ["C1-BASE"]


def test_cores_pulam_removendo():
    with open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8") as f:
        js = f.read()
    corpo = js[js.index("function aplicarCoresParametros()"):js.index("function agendarCoresAposDesenhar()")]
    assert "['R', '*R'].includes" in corpo


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_nome_com_virgula_cadastrado_como_literal_existe_na_base():
    """[CAZ7,94] na base: o cabo «CAZ7,94 35 m» (sem colchetes) não fica vermelho; o que não existe continua vermelho."""
    js = """const N = require(process.argv[1]);
const ns = N.nomesDaBase([{ ativo: '[CAZ7,94]', origem: 'CABOS' }, { ativo: 'N1' }], '', 'CABOS');
process.stdout.write(JSON.stringify(N.foraDaBase(ns, ['CAZ7,94', '[CAZ7,94]', 'N1(A)', 'CAZ7', 'XYZ', 'CAZ7,5'])));"""
    r = subprocess.run([NODE, "-e", js, os.path.join(config.STATIC_DIR, "ativo_na_base.js")], capture_output=True, text=True, check=True)
    assert json.loads(r.stdout) == ["CAZ7", "XYZ", "CAZ7,5"]
