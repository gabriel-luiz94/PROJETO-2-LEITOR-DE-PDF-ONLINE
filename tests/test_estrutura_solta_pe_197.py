"""TASK-197 — Pernambuco: estrutura solta (N3, B2, 2 N3, N1(XLPE,DT)…) classifica como ESTRUTURA; poste, cabo e materiais não mudam."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
pytestmark = pytest.mark.skipif(not NODE, reason="node indisponível")
D = os.path.join(config.STATIC_DIR, "..", "data")


def _classificar(textos):
    js = ("const E = require(%s), rs = require(%s); console.log(JSON.stringify(%s.map(t => E.classificar('I', t, '', '', t, rs).entidade)));"
          % (json.dumps(os.path.join(config.STATIC_DIR, "regras_leitor_engine.js")), json.dumps(os.path.join(D, "regras_leitor_classificacao_pernambuco_seed.json")), json.dumps(textos)))
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def test_estrutura_solta_vira_estrutura_e_o_resto_continua():
    textos = ["N3", "1-N3", "2 N3", "N1(XLPE,DT)", "N1, N4", "B2", "CE1", "*2 U1", "11300D, N3", "11300D", "3xA35P(7,94)-13,8kV 25 m", "6 CON", "PR220", "FLABIT"]
    assert _classificar(textos) == ["ESTRUTURA"] * 8 + ["POSTE", "POSTE", "CABO", "0", "0", "0"]


def test_seed_novo_e_valido_e_seed_antigo_do_pacote_e_atualizado():
    from routers.regras_leitor import validar_regras, _hash_regras
    cls = json.load(open(os.path.join(D, "regras_leitor_classificacao_pernambuco_seed.json"), encoding="utf-8"))
    assert validar_regras("classificacao", cls) == [] and [r["entidade"] for r in cls] == ["POSTE", "CABO", "ESTRUTURA"]
    man = json.load(open(os.path.join(D, "pacotes", "081", "manifest.json"), encoding="utf-8"))
    antigo = cls[:2]                                                       # o seed de antes desta tarefa (só poste e cabo) é reconhecido como «intocado»
    assert _hash_regras(antigo) in man["seeds_antigos"]["classificacao"] and _hash_regras(cls) not in man["seeds_antigos"]["classificacao"]
