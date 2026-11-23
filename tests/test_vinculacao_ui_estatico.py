"""TASK-084 — o modal de vinculação em abas: arquivos ligados na tela, sintaxe e contratos que o resumo.js usa."""
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")


def _ler(nome):
    with open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8") as f:
        return f.read()


def test_scripts_na_ordem_certa_no_resumo():
    html = _ler("resumo.html")
    ordem = [html.index(f"/static/{n}.js") for n in ("anotacao_outros", "anotacao_cabos", "vinculacao_regras", "vinculacao_ui", "resumo")]
    assert ordem == sorted(ordem)


@pytest.mark.skipif(not NODE, reason="node indisponível")
@pytest.mark.parametrize("arq", ["vinculacao_ui.js", "vinculacao_regras.js", "resumo.js"])
def test_sintaxe_dos_arquivos(arq):
    r = subprocess.run([NODE, "--check", os.path.join(config.STATIC_DIR, arq)], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr


def test_resumo_liga_o_host_e_a_totalizadora_ao_vinculo():
    js = _ler("resumo.js")
    for trecho in ["window.VinculacaoHost = {", "window.VinculacaoUI.abrir(window.VinculacaoHost)", "window.abrirModalConectores", "_parametrosVinculoPorLinha()",
                   "parametros_estrutura: parametrosVinculo[idxOutros]", "obj.cabosFicticios", "obj.vinculoOrigem", "/api/vinculacao-config?projeto_codigo=", "VinculacaoRegras.achados("]:
        assert trecho in js, trecho
    ui = _ler("vinculacao_ui.js")
    for metodo in ["host.cabos()", "host.outros()", "host.antes()", "host.desfazer()", "host.projeto()", "host.prefixoCabo", "host.codigoEstrutura", "host.postesDaLinha", "host.linhaSoPoste"]:
        assert metodo in ui, metodo
    for fn in ["cabos", "outros", "projeto", "prefixoCabo", "codigoEstrutura", "postesDaLinha", "linhaSoPoste", "antes", "desfazer", "depois"]:
        assert fn in js.split("window.VinculacaoHost = {")[1].split("window.abrirModalVinculacao")[0], fn


def test_modal_antigo_so_tem_conectores():
    html = _ler("resumo.html")
    assert "body-vinc-conectores" in html and "vinculacao-lista" not in html and "body-vinc-regras" not in html


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_edicao_do_mapa_preserva_os_vinculos_e_a_identidade_da_linha():
    resumo = _ler("resumo.js")
    trecho = resumo[resumo.index("    function validarEdicaoCaboNoMapa("):resumo.index("    function inserirAtivoNoMapa(")]
    preparar = r"""
const assert = require('node:assert/strict');
const VinculacaoRegras = require(process.argv[1]);
const poste = {ativo:'12600D, N3, N1', entidade:'POSTE', operacao:'I', _poste:[0,0]};
const cabo = {ativo:'A35P 25 M', operacao:'I', vinculoEstruturas:[.01], vinculoPostes:[0], _mapaTopologia:true};
const tableStates = {outros:{data:[poste]}, cabos:{data:[cabo]}};
const _codigosEstruturaVinculo = a => a.split(',').slice(1).map(x => x.trim());
const _codigoEstruturaVinculo = a => _codigosEstruturaVinculo(a)[0];
const _prefixoCaboVinculo = a => a.split(' ')[0];
const validarAnotacaoOutrosNoMapa = () => '';
const reclassificarEntidade = () => {};
let passos = 0, atualizou = [];
const pushHistory = () => passos++;
const atualizarLinhaDoMapa = (...args) => atualizou.push(args);
const aplicarOperacaoNaLinha = (row, v) => { row.operacao = v; };          // TASK-235 (a regra do F é testada em test_ponto_linhas_235)
"""
    testar = r"""
const cfg = {estruturas:[{padrao:'^N1$', cabos:'^A35P$'}]};
const edit = (tipo, linha, dados) => editarAtivoNoMapa(tipo, 0, linha, dados, cfg);
assert.match(edit('poste', poste, {ativo:'12600D, N1, N3', operacao:'I'}), /estrutura vinculada/);
assert.equal(edit('poste', poste, {ativo:'12400D, N3, N1', operacao:'I'}), '');
assert.equal(poste.ativo, '12400D, N3, N1');
assert.deepEqual(poste._poste, [0,0]);
assert.equal(edit('cabo', cabo, {ativo:'A35P 30 M', operacao:'I', nivel:'BT', fases:'2', tensao:''}), '');
assert.deepEqual(cabo.vinculoEstruturas, [.01]);
assert.deepEqual(cabo.vinculoPostes, [0]);
assert.equal(cabo.ativo, 'A35P 30 M');
assert.match(edit('cabo', cabo, {ativo:'A35P 30 M', operacao:'R', nivel:'BT', fases:'2', tensao:''}), /retirada/);
assert.equal(passos, 2);
assert.deepEqual(atualizou, [['outros',0,false], ['cabos',0,true]]);
"""
    subprocess.run([NODE, "-e", preparar + trecho + testar, os.path.join(config.STATIC_DIR, "vinculacao_regras.js")], check=True)
