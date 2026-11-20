"""Comparação de preenchimento nas três entradas de obra do Resumo."""
import shutil
import subprocess
from pathlib import Path

import pytest


@pytest.mark.skipif(not shutil.which("node"), reason="node indisponível")
def test_comparacao_preserva_cabos_e_detecta_mudancas_sem_usar_ordem_das_linhas():
    script = r"""
const assert = require('node:assert/strict');
const { comparar } = require('./static/comparar_obras.js');
const poste = (ativo, placa, x) => ({ ativo, placa, entidade:'POSTE', operacao:'I', _poste:[x,0] });
const cabo = (ativo, a, b, vinculoEstruturas = []) => ({ ativo, entidade:'CABO', operacao:'I', _geo:[[a,0],[b,0]], vinculoEstruturas });
const opcoes = {
  cabo: ativo => ({ prefixo: ativo.split(' ')[0], comprimento: Number(ativo.split(' ')[1]) }),
  estruturas: linha => linha.ativo.split(',').slice(1).map(x => x.trim()),
  codigoOutro: linha => linha.ativo
};
const a = { outros: {data: [poste('12600D, N3', 'B039104', 0), poste('12400D, N2', 'B039105', 100)]},
  cabos: [cabo('A35P 42', 0, 100, [0,1]), cabo('C16 20', 0, 100), cabo('RAMA25 3', 100, 200)] };
const b = { outros: [poste('12400D, N2', 'B039105', 100), {...poste('12600D, N3', 'B039104', 0), operacao:'R'}],
  cabos: [cabo('C16 20', 100, 0), cabo('A35P 47', 100, 0, [1,0]), cabo('RAMA70 9', 100, 300)] };
const original = JSON.stringify([a,b]);
const r = comparar(a,b,opcoes);
assert.deepEqual(r.outros.map(x => [x.estado,x.campos]), [['diferente',['Operação']],['igual',[]]]);
assert.deepEqual(r.cabos.map(x => x.estado), ['diferente','igual','referencia','comparada']);
assert(r.cabos[0].campos.includes('Metragem'));
assert.equal(r.cabos[0].deltaMetros, 5);
assert(!r.cabos[0].campos.includes('Vínculos'));
assert.equal(r.cabos[0].comparada.i, 1);
assert.equal(JSON.stringify([a,b]), original);
b.cabos[1].vinculoEstruturas = [0];
assert(comparar(a,b,opcoes).cabos[0].campos.includes('Vínculos'));

const paralelosA = { cabos: [cabo('A35P 40', 0, 100), cabo('A35P 42', 0, 100)] };
const paralelosB = { cabos: [cabo('A35P 43', 0, 100), cabo('A35P 45', 0, 100)] };
const ambiguos = comparar(paralelosA,paralelosB,opcoes).cabos;
assert.equal(ambiguos.length,4);
assert(ambiguos.every(x => x.estado === 'conferir' && !(x.referencia && x.comparada)));

const u = { cabos: [{ativo:'C16 20'}, {ativo:'C16 20'}] };
const v = { cabos: [{ativo:'C16 20'}, {ativo:'C16 21'}] };
assert.deepEqual(comparar(u,v,opcoes).cabos.map(x=>x.estado), ['igual','conferir']);
"""
    subprocess.run(["node", "-e", script], check=True, cwd=Path(__file__).resolve().parents[1])
