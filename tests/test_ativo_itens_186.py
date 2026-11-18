"""TASK-186 — «Ocultar parâmetros»: o campo do ativo mostra e edita só os NOMES; os parâmetros ficam presos a cada item (apagou/renomeou = item novo);
balão interativo dos ativos/parâmetros (Ctrl+Enter, duplo clique, mouse parado, «(»); Tab Operação → Ativo → Operação da linha de baixo."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
ITENS = os.path.join(config.STATIC_DIR, "ativo_itens.js")


def _ler(nome):
    with open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8") as f:
        return f.read()


def _node(js, entrada):
    r = subprocess.run([NODE, "-e", "const A = require(%s); const e = JSON.parse(require('fs').readFileSync(0, 'utf8')); %s" % (json.dumps(ITENS), js)], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


CHEIO = "12600D, N1(XLPE,D,A), N1(EPR,R,B), N4"


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_itens_compacto_e_juntar():
    r = _node("const it = A.itens(e.t); console.log(JSON.stringify({ c: A.compacto(e.t), p: it.map(i => i.params), j: A.juntar(it), n: it.map(i => i.nome) }));", {"t": CHEIO})
    assert r["c"] == "12600D, N1, N1, N4" and r["p"] == ["", "(XLPE,D,A)", "(EPR,R,B)", ""] and r["j"] == CHEIO
    assert r["n"] == ["12600D", "N1", "N1", "N4"]


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_virgula_dentro_de_parenteses_e_colchetes_nao_divide():
    r = _node("console.log(JSON.stringify(A.itens(e.t).map(i => [i.nome, i.params])));", {"t": "PR(D,X), [CAZ7,94], CE_4(12600D,A35P)"})
    assert r == [["PR", "(D,X)"], ["[CAZ7,94]", ""], ["CE_4", "(12600D,A35P)"]]


CASOS = [
    # (novo texto compacto, cursor, esperado) — o cursor diz qual dos N1 repetidos foi apagado
    ("12600D, N1, N4", 8, "12600D, N1(EPR,R,B), N4"),                       # apagou o 1º N1: o 2º segue com os SEUS parâmetros
    ("12600D, N1, N4", 12, "12600D, N1(XLPE,D,A), N4"),                     # apagou o 2º N1
    ("12600D, N1, N1, N4, N5", 22, "12600D, N1(XLPE,D,A), N1(EPR,R,B), N4, N5"),      # item novo no fim
    ("12600D, N1, N1, N5", 22, "12600D, N1(XLPE,D,A), N1(EPR,R,B), N5"),   # N4 renomeado para N5: é novo
    ("12600D, N1, N1", 22, "12600D, N1(XLPE,D,A), N1(EPR,R,B)"),            # N4 apagado
    ("12600D, N1X, N1, N4", 16, "12600D, N1X, N1(EPR,R,B), N4"),            # 1º N1 renomeado: perde os parâmetros
    ("N1(A,B), N4", None, "N1(A,B), N4"),                                   # colou com parênteses: os parâmetros vão junto
]


@pytest.mark.skipif(not NODE, reason="node indisponível")
@pytest.mark.parametrize("novo,cursor,esperado", CASOS)
def test_parametros_ficam_presos_ao_item(novo, cursor, esperado):
    r = _node("const it = A.itens(e.t); console.log(JSON.stringify(A.aplicarEdicao(it, A.compacto(e.t), e.novo, e.cursor).texto));", {"t": CHEIO, "novo": novo, "cursor": cursor})
    assert r == esperado


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_conteudo_e_com_params():
    r = _node("const it = A.itens(e.t); console.log(JSON.stringify([A.conteudo(it[0].params), A.juntar(A.comParams(it, 0, '(cabo=X)')), A.juntar(A.comParams(it, 0, '()'))]));", {"t": "N1(D,A), N4"})
    assert r == ["D,A", "N1(cabo=X), N4", "N1, N4"]


def test_tela_campo_compacto_editavel_e_atalhos():
    js, html = _ler("resumo.js"), _ler("resumo.html")
    for trecho in ["const oculto = type === 'outros' && !!window.AtivoItens && ocultarParametrosLigado()", "window.AtivoItens.aplicarEdicao(itensDoCampo, compactoDoCampo", "inpAt._gravarCheio = novo =>",
                   "const confirmarMudanca = snapAtivo =>", "window.AtivoBalao.abrirNoCursor(inpAt)", "e.key === 'Enter' && e.ctrlKey", "window.AtivoBalao.iniciar({", "renderTable('outros'); } catch (e)",
                   "irParaCelula(type, idx + 1, 0, 1)", "(e.key === 'Enter' || e.key === 'Tab') && !e.shiftKey"]:
        assert trecho in js, trecho
    assert "ativo-compacto" not in js and "tem-params" not in js and "ativo-compacto" not in html          # acabou o rótulo sobreposto (campo transparente)
    assert "ativo_itens.js" in html and "ativo_balao.js" in html and ".ab-balao" in html
    assert "if (el.dataset.parTipo === 'outros' && window.AtivoBalao) return;" in js                           # o balão só-leitura antigo fica para Cabos/Totalizadora


def test_balao_teclado_e_mouse():
    b = _ler("ativo_balao.js")
    for trecho in ["DEMORA_HOVER = 1500", "'dblclick'", "k === 'ArrowDown' || (k === 'Tab' && !e.shiftKey)", "if (k === 'Tab') { consumir(); moverParametro", "k === 'Delete' && !noLivre", "irParaProximoComPendencia",
                   "S.modo = 'ativos'; S.livre = ''", "sugestao-ativos"]:
        assert trecho in b, trecho
    assert "AtivoDica.parametrosEditaveis" in b.replace("raiz.AtivoDica", "AtivoDica") or "AD.parametrosEditaveis" in b


def test_dica_exporta_contexto_atual():
    assert "contextoAtual" in _ler("ativo_dica.js")


@pytest.mark.parametrize("arq", ["ativo_itens.js", "ativo_balao.js", "ativo_dica.js", "resumo.js"])
@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_sintaxe(arq):
    r = subprocess.run([NODE, "--check", os.path.join(config.STATIC_DIR, arq)], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr


def test_balao_sempre_tem_a_linha_digitar_e_escrever_direto_vai_para_ela():
    """Regressão: a linha «✎ digitar outro valor…» só existia em parâmetros 'abertos'; agora está em todos, e uma letra digitada com o parâmetro marcado cai nela."""
    b = _ler("ativo_balao.js")
    assert "l.push({ custom: true }); return l; };" in b and "p.aberto || !l.length" not in b
    assert "S.livre = k; S.focoLivre = true; desenhar();" in b                       # escrever direto
    assert "function posicionar(p)" in b and "o.custom && p.atual != null" in b      # valor atual fora da lista já vem escrito na linha «digitar»


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_ativo_com_quantidade_e_procurado_pelo_nome_sem_a_quantidade():
    """Regressão: «6 CON» / «6-CON» aparecia como «ativo não encontrado na base» no balão (a quantidade ia junto na busca)."""
    r = _node("console.log(JSON.stringify([e.t.map(A.semQuantidade), A.itens('12600D, 6 CON, 6-N1(XLPE), *2-IT_2').map(i => [i.nome, i.base, i.params])]));",
              {"t": ["6 CON", "6-CON", "*1,5-IT_2", "12600D", "3XC16", "9200D", "N1", "2 N1"]})
    assert r[0] == ["CON", "CON", "IT_2", "12600D", "3XC16", "9200D", "N1", "N1"]
    assert r[1] == [["12600D", "12600D", ""], ["6 CON", "CON", ""], ["6-N1", "N1", "(XLPE)"], ["*2-IT_2", "IT_2", ""]]
    b, js = _ler("ativo_balao.js"), _ler("resumo.js")
    assert "cfg.dica(it.base || it.nome)" in b and "it.base || it.nome" in js
