import shutil
import subprocess
from pathlib import Path

import pytest


@pytest.mark.skipif(not shutil.which("node"), reason="node indisponível")
def test_grafo_agrupa_postes_e_vaos_e_mantem_as_associacoes():
    script = r"""
const assert = require('node:assert/strict');
const vr = require('./static/vinculacao_regras.js');
const outros = [
  {entidade:'POSTE', ativo:'9100D,N3', _poste:[0,0]},
  {entidade:'ESTRUTURA', ativo:'N1', _poste:[0,0]},
  {entidade:'POSTE', ativo:'9200D,N3', _poste:[40,0]},
  {entidade:'POSTE', ativo:'9300D,N3'}
];
const cabos = [
  {ativo:'A 40 m', _geo:[[1,0],[20,4],[39,0]]},
  {ativo:'B 40 m', _geo:[[39,0],[20,4],[1,0]], operacao:'R'},
  {ativo:'C 40 m', vinculoPostes:[2,3]},
  {ativo:'D 40 m', _geo:[[40,0],[70,20]]},
  {ativo:'E 40 m', vinculoEstruturas:[1]},
  {ativo:'F 40 m', _geo:[null,[40,0]]}
];
const codigos = ativo => [ativo.split(',').pop()];
const grafo = vr.grafoDaRede(cabos, outros, v => codigos(v)[0], codigos);
assert.equal(grafo.nos.size, 3);
assert.deepEqual(grafo.nos.get('0,0').linhas, [0,1]);
assert.deepEqual(grafo.nos.get('0,0').estruturas, [0,1]);
assert.deepEqual([...grafo.nos.get('0,0').cabos], [0,1,4]);
assert.deepEqual(grafo.vaos[0].cabos, [0,1]);
assert.deepEqual(grafo.vaos[1].ids, ['40,0', 'linha:3']);
assert.deepEqual(grafo.vaos[2].ids, ['40,0', null]);
assert.equal(grafo.pontasSoltas, 1);
assert.equal(grafo.semVao, 2);
assert.deepEqual(cabos[4].vinculoEstruturas, [1]);
"""
    subprocess.run(["node", "-e", script], check=True, cwd=Path(__file__).resolve().parents[1])
