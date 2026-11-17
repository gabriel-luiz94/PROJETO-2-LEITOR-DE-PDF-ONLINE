"""Base de Orçamento mais rápida: sync da mestre com intervalo, aviso entre telas (base_sinal.js), barra de status (base_status.js) e índice de pesquisa."""
import json
import os
import subprocess

import pytest

import config
from services import sync_service
from tests import oraculo_tela as o

node = pytest.mark.skipif(not o.NODE, reason="node não instalado")


def _js(codigo):
    r = subprocess.run([o.NODE, "-e", codigo], capture_output=True, text=True, timeout=30, cwd=config.STATIC_DIR)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout.strip().splitlines()[-1])


def test_sync_mestre_so_baixa_de_novo_depois_do_intervalo(monkeypatch):
    chamadas = []
    monkeypatch.setattr(sync_service, "sync_tabela_master", lambda: chamadas.append(1) or True)
    sync_service.invalidar_sync_master()
    assert sync_service.sync_tabela_master_se_preciso(ttl=60) is True
    assert sync_service.sync_tabela_master_se_preciso(ttl=60) is False       # dentro do intervalo: pula
    assert sync_service.sync_tabela_master_se_preciso(ttl=60) is False
    assert len(chamadas) == 1
    sync_service.invalidar_sync_master()
    assert sync_service.sync_tabela_master_se_preciso(ttl=60) is True        # invalidado: baixa
    assert len(chamadas) == 2
    sync_service.invalidar_sync_master()


def test_sync_mestre_que_falha_nao_trava_a_proxima_tentativa(monkeypatch):
    resultados = iter([False, True])
    monkeypatch.setattr(sync_service, "sync_tabela_master", lambda: next(resultados))
    sync_service.invalidar_sync_master()
    assert sync_service.sync_tabela_master_se_preciso(ttl=60) is False
    assert sync_service.sync_tabela_master_se_preciso(ttl=60) is True        # a falha não marcou o horário
    sync_service.invalidar_sync_master()


def test_get_dados_usa_o_sync_com_intervalo(monkeypatch):
    from routers import orcamento
    assert orcamento.sync_tabela_master_se_preciso is sync_service.sync_tabela_master_se_preciso
    fonte = open(os.path.join(config.BASE_DIR if hasattr(config, "BASE_DIR") else ".", "routers", "orcamento.py"), encoding="utf-8").read()
    corpo = fonte[fonte.index("def get_orcamento_dados"):fonte.index("def salvar_orcamento")]
    assert "sync_tabela_master_se_preciso()" in corpo and "sync_tabela_master()" not in corpo


@node
def test_sinal_avisa_os_ouvintes_e_cancela():
    r = _js("""
global.window = undefined; delete global.BroadcastChannel;      // (o canal mantém o Node aberto)
const S = require('./base_sinal.js'); const v = [];
const off = S.ouvir(o => v.push(o));
S.avisar(); off(); S.avisar();
console.log(JSON.stringify(v));""")
    assert r == ["local"]


@node
def test_sinal_de_outra_aba_chega_pelo_evento_storage():
    r = _js("""
const ev = {}; const mem = {};
global.localStorage = { setItem: (k, v) => { mem[k] = v; }, getItem: k => mem[k] };
const w = { localStorage: global.localStorage, addEventListener: (n, f) => { ev[n] = f; } };
global.window = w; delete global.BroadcastChannel;
const S = require('./base_sinal.js');
const v = []; S.ouvir(o => v.push(o));
ev.storage({ key: S.CHAVE, newValue: '123' }); ev.storage({ key: S.CHAVE, newValue: '123' }); ev.storage({ key: 'outra', newValue: '1' });
console.log(JSON.stringify(v));""")
    assert r == ["storage"]


@node
def test_barra_de_status_trava_botoes_e_so_a_ultima_operacao_mexe():
    r = _js("""
const botoes = [{ disabled: false, dataset: {} }, { disabled: true, dataset: {} }];
const mk = () => { const e = { style: {}, children: [], className: '', textContent: '', dataset: {}, setAttribute() {}, append(...c) { this.children.push(...c); }, appendChild(c) { this.children.push(c); } }; return e; };
global.document = { head: mk(), body: mk(), createElement: mk, querySelectorAll: () => botoes };
global.window = undefined;
const B = require('./base_status.js');
const a = B.iniciar('A'); const durante = botoes.map(b => b.disabled);
const b = B.iniciar('B'); a.ok('tarde');                 // 'a' terminou, mas 'b' é a mais recente
const txt = () => document.body.children[0].children[1].textContent;
const meio = txt(); b.ok('fim');
console.log(JSON.stringify({ durante, meio, fim: txt(), depois: botoes.map(x => x.disabled) }));""")
    assert r["durante"] == [True, True] and r["meio"] == "B" and r["fim"] == "fim"
    assert r["depois"] == [False, True]           # o que já estava desabilitado continua assim


@node
def test_pesquisa_com_indice_acha_ativo_apelido_e_codigo():
    r = _js("""
const D = require('./base_ficha_dados.js');
const L = (ativo, comp, cod, desc, proj) => ({ ativo, desc_ativo: 'd', componente: comp, projeto: proj || 'PE', mdo: 'MATERIAL', codigo: cod, desc_codigo: desc, fator_i: 1, fator_r: 0, filtro: '', lv: 'A', arred: 'NORMAL', formula_i: '', formula_r: '' });
const m = D.construir([L('CHAVE_FACA=CHAVE_FACA(a=1), SUH', '', '', ''), L('CHAVE_FACA', 'CF', '500109', 'CH SEC'), L('OUTRO', 'O', '500109', 'CH SEC')], 'PE');
const a = D.pesquisar(m, 'suh'), b = D.pesquisar(m, 'chave'), c = D.pesquisar(m, '5001');
const idx1 = m._idxPesquisa; D.pesquisar(m, 'x');
console.log(JSON.stringify({ a: a.map(x => [x.tipo, x.nome, x.via]), b: b[0].nome, c: c.filter(x => x.tipo === 'codigo').map(x => x.usos), mesmo: idx1 === m._idxPesquisa, json: 'idx' in JSON.parse(JSON.stringify(m)) }));""")
    assert r["a"][0][:2] == ["ativo", "CHAVE_FACA"] and r["a"][0][2].upper() == "SUH"
    assert r["b"] == "CHAVE_FACA" and r["c"] == [2] and r["mesmo"] is True and r["json"] is False
