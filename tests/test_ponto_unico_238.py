"""TASK-238: o grupo «@nome» é UM ponto no mapa, com todas as operações; a lista de estruturas do ponto mostra todas (instalando em destaque)."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
com_node = pytest.mark.skipif(not NODE, reason="node indisponível")


def _node(js):
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=60, cwd=config.BASE_DIR)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


GRAFO = """
const VR = require('./static/vinculacao_regras.js');
const cod = a => (String(a).match(/\\b(N\\d|CE\\d)\\b/) || [])[1] || '';
const cods = a => String(a).match(/\\b(N\\d|CE\\d)\\b/g) || [];
const nos = outros => [...VR.grafoDaRede([], outros, cod, cods).nos.values()].map(n => ({ id: n.id, linhas: n.linhas, poste: n.poste }));
const t = (gid, nome) => ({ entidade: '0', ativo: '@' + nome, grupoDecl: true, gid, operacao: 'I' });
"""


@com_node
def test_ponto_feito_a_mao_sem_coordenada_e_um_no_so():
    r = _node(GRAFO + """
console.log(JSON.stringify({
  mao: nos([t('g1', 'P1'), { entidade: 'POSTE', ativo: '11300D, N3', operacao: 'I', grupoDe: 'g1' }, { entidade: 'ESTRUTURA', ativo: 'CE2', operacao: 'R', grupoDe: 'g1' },
            { entidade: 'POSTE', ativo: '12600D', operacao: 'F', grupoDe: 'g1', congelada: true }, t('g2', 'P2'), { entidade: 'ESTRUTURA', ativo: 'N3', operacao: 'I', grupoDe: 'g2' }]),
  aprox: nos([t('g1', 'P1'), { entidade: 'POSTE', ativo: '11300D, N3', operacao: 'I', grupoDe: 'g1', _posteAprox: [5, 5] }, { entidade: 'ESTRUTURA', ativo: 'CE2', operacao: 'R', grupoDe: 'g1', _posteAprox: [5, 5] }]),
  coord: nos([t('g1', 'P1'), { entidade: 'ESTRUTURA', ativo: 'CE2', operacao: 'R', grupoDe: 'g1' }, { entidade: 'POSTE', ativo: '11300D, N3', operacao: 'I', grupoDe: 'g1', _poste: [1, 2] }]),
  semGrupo: nos([{ entidade: 'ESTRUTURA', ativo: 'N3', operacao: 'I' }, { entidade: 'ESTRUTURA', ativo: 'CE2', operacao: 'R' }]),
}));""")
    assert r["mao"] == [{"id": "linha:1", "linhas": [1, 2, 3], "poste": 1}, {"id": "linha:5", "linhas": [5], "poste": None}]   # poste não fictício na frente do F
    assert r["aprox"] == [{"id": "linha:1", "linhas": [1, 2], "poste": 1}]                    # plugin sem símbolo
    assert r["coord"] == [{"id": "1,2", "linhas": [1, 2], "poste": 2}]                        # a linha sem coordenada vai para a do grupo
    assert len(r["semGrupo"]) == 2                                                            # sem grupo: como antes


@com_node
def test_poste_ficticio_so_fica_como_poste_do_ponto_sem_outro():
    r = _node(GRAFO + """
console.log(JSON.stringify(nos([t('g1', 'P1'), { entidade: 'POSTE', ativo: '12600D', operacao: 'F', grupoDe: 'g1' }, { entidade: 'ESTRUTURA', ativo: 'N3', operacao: 'I', grupoDe: 'g1' }])));""")
    assert r == [{"id": "linha:1", "linhas": [1, 2], "poste": 1}]


def test_painel_lista_todas_as_estruturas_com_instalando_em_destaque():
    src = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    assert "const OPS_INSTALANDO = ['I', '*I', 'M', '*M'];" in src
    assert "!VR.participaVinculo(r)" in src and "'vz-no-estrutura removendo'" in src and "removendo · fora do vínculo" in src
    assert ".vz-no-estrutura.instalando{" in src and ".vz-no-estrutura.removendo{opacity:.55}" in src
    assert "outros()[linha].grupoDe) continue;" in src                                        # sem «Associar a poste» para linha de grupo
