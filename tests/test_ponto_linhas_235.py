"""TASK-235: o ponto do mapa = «@nome - comentário» + linhas com operação (F = fictícia); POSTE do ponto; conversão do formato antigo."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import anotacao_outros as ao
from services import obras_arquivo as oa
from services import vinculacao_regras as vr

NODE = shutil.which("node")
S = config.STATIC_DIR
com_node = pytest.mark.skipif(not NODE, reason="node indisponível")


def _node(js):
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=60, cwd=config.BASE_DIR)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def _ler(nome):
    return open(os.path.join(S, nome), encoding="utf-8").read()


@com_node
def test_titulo_com_comentario_e_conversao_do_formato_antigo():
    r = _node("""
const M = require('./static/grupos_linhas.js');
const lista = [{ativo: '@P1', grupoDecl: true, gid: 'a'}, {ativo: '# poste com telefonia', comentario: true, grupoDe: 'a'}, {ativo: '#cliente às 8h', comentario: true, grupoDe: 'a'},
  {ativo: '11300D, N3', grupoDe: 'a'}, {ativo: '@P2 - já novo', grupoDecl: true, gid: 'b'}, {ativo: 'N3', grupoDe: 'b'}];
const p = M.pontoDaLinha(lista, 3), novo = M.pontoDaLinha(lista, 5);
const conv = M.converterPontos(lista);
console.log(JSON.stringify({ sep: M.separarTitulo('@P12 - N3 solta - esquina'), sem: M.separarTitulo('@P3 RUA'), tit: M.tituloDoPonto('P7', '  # esquina '),
  nome: M.nomeLimpo('a - b'), p: [p.nome, p.comentario, p.formatoAntigo, p.membros], novo: [novo.nome, novo.comentario, novo.formatoAntigo],
  conv: [conv.n, conv.lista.map(x => x.ativo)], prox: M.proximoNome(conv.lista) }));""")
    assert r["sep"] == {"nome": "P12", "comentario": "N3 solta - esquina"} and r["sem"] == {"nome": "P3 RUA", "comentario": ""}
    assert r["tit"] == "@P7 - esquina" and r["nome"] == "A-B"
    assert r["p"] == ["P1", "poste com telefonia; cliente às 8h", True, [3]]            # opção A: junta por «; »
    assert r["novo"] == ["P2", "já novo", False]
    assert r["conv"] == [1, ["@P1 - poste com telefonia; cliente às 8h", "11300D, N3", "@P2 - já novo", "N3"]]
    assert r["prox"] == "P3"


@com_node
def test_linha_f_e_estrutura_e_cabo_f_e_ficticio_js():
    r = _node("""
const v = require('./static/vinculacao_regras.js');
const cfg = v.normalizarConfig({cabos: [{padrao: '^A35P$'}], estruturas: [{padrao: '^CEJ2$', qtd_cabos: 2, lados: 'IGUAIS', max_reais: 1}]});
const outros = [{entidade: 'ESTRUTURA', operacao: 'F', ativo: 'CEJ2', congelada: true}];
const cabos = [{entidade: 'CABO', operacao: 'I', ativo: 'A35P 30 m', vinculoEstruturas: [0]}, {entidade: 'CABO', operacao: 'F', ativo: 'A35P 30 m', congelada: true, vinculoEstruturas: [0]}];
const p = a => (a || '').split(/[ ,]/)[0].toUpperCase(), cods = a => [p(a)];
const a = v.analisar(cabos, outros, cfg, p, p, cods);
const e = v.infoEstrutura('CEJ2', cfg), real = v.infoCabo('A35P', 'I', cfg), fic = v.infoCabo('A35P', 'F', cfg);
console.log(JSON.stringify({ f: v.ehFicticia(outros[0]), links: a.links.map(l => l.ok), fic: fic.ficticio,
  iguais: v.podeComOcupantes(fic, e, [{info: real, ficticio: false}])[0], reais: v.podeComOcupantes(real, e, [{info: real, ficticio: false}])[0] }));""")
    assert r["f"] is True and r["fic"] is True
    assert r["links"] == [True, True]                     # a linha F (congelada) continua recebendo cabos
    assert r["iguais"] is True                            # cabo F não esbarra em «iguais»/«máx. reais»
    assert r["reais"] is False


def test_linha_f_e_estrutura_python_e_poste_do_ponto():
    assert vr.eh_ficticia({"operacao": "f"}) and not vr.eh_ficticia({"operacao": "I"})
    assert vr._eh_linha_estrutura({"entidade": "ESTRUTURA", "operacao": "F", "ativo": "N3", "congelada": True})
    assert not vr._eh_linha_estrutura({"entidade": "ESTRUTURA", "operacao": "I", "ativo": "N3", "congelada": True})
    cfg = vr.normalizar_config({"cabos": [], "estruturas": []})
    assert vr.info_cabo("A35P", "F", cfg)["ficticio"] is True and vr.info_cabo("A35P", "I", cfg)["ficticio"] is False
    ctx = {"assinaturas": {"N3": {"nomes": ["POSTE"]}}}
    assert ao.interpretar_linha("N3", "I", None, dict(ctx, poste_do_ponto="12600D"))["itens"][0]["ativo"] == "N3(POSTE=12600D)"
    assert ao.interpretar_linha("N3", "I", None, ctx)["itens"][0]["ativo"] == "N3"


@com_node
def test_poste_do_ponto_js_e_o_da_propria_linha_manda():
    r = _node("""
const A = require('./static/anotacao_outros.js'), ctx = {assinaturas: {N3: {nomes: ['POSTE']}}, poste_do_ponto: '12600D'};
console.log(JSON.stringify([A.interpretarLinha('N3', 'I', null, ctx).itens.map(i => i.ativo), A.interpretarLinha('11300D, N3', 'I', null, ctx).itens.map(i => i.ativo).pop()]));""")
    assert r == [["N3(POSTE=12600D)"], "N3(POSTE=11300D)"]


def test_obra_guarda_operacao_f():
    dados = {"cabos": [{"ativo": "CAA2 ABC 30 m", "entidade": "CABO", "operacao": "F", "congelada": True, "_congeladaPorF": True}],
             "outros": [{"ativo": "12600D", "entidade": "POSTE", "operacao": "F", "congelada": True, "_congeladaPorF": True}]}
    r = oa.validar_importacao(json.loads(json.dumps(oa.montar_exportacao({"nome": "F", "projeto": "081", "dados_json": dados}))), "081")["dados"]
    assert r["cabos"][0]["operacao"] == "F" and r["cabos"][0]["_congeladaPorF"] is True and r["outros"][0]["congelada"] is True


def test_tela_tem_o_novo_conceito_do_ponto():
    ui = _ler("vinculacao_ui.js")
    corpo = ui[ui.index("function abaMapa(corpo)"):]
    for trecho in ("botao('+ Ponto'", "botao('Traçar cabo'", "function abrirPonto(", "function seletorOperacao(", "function secaoLinhasDoPonto(", "function escolherCaboNo(",
                   "'Traçar cabo daqui'", "'Converter todos'", "'Converter este ponto'", "host.converterPontosNoMapa(", "host.salvarPontoNoMapa("):
        assert trecho in corpo, trecho
    barra = corpo[corpo.index("const ferramentas = el('div'"):corpo.index("const btnAgrupar")]
    assert "'+ Poste'" not in barra and "'+ Estrutura'" not in barra
    res = _ler("resumo.js")
    assert "const OPERACOES_LINHA = OPERACOES.concat(['F'])" in res and "function salvarPontoNoMapa(" in res and "function aplicarOperacaoNaLinha(" in res
    assert "salvarPontoNoMapa, converterPontosNoMapa," in res and "_ctxComPosteDoPonto(" in res
