"""Oráculo dos testes de paridade (TASK-031, fase A): executa no Node o JavaScript REAL da tela.

Nada é reescrito aqui: os trechos de static/resumo.js (que roda dentro de um DOMContentLoaded e não pode ser importado) são
FATIADOS do arquivo por marcadores e executados com `new Function`. Se os marcadores sumirem (a tela foi reorganizada), o
teste falha de propósito para o autor revisar o porte em services/autonomo/montagem.py.
"""
import json
import os
import shutil
import subprocess

import config

NODE = shutil.which("node")
MOTOR = os.path.join(config.STATIC_DIR, "regras_leitor_engine.js")
RESUMO = os.path.join(config.STATIC_DIR, "resumo.js")


def _fatia(texto, inicio, fim):
    i = texto.index(inicio)
    return texto[i:texto.index(fim, i)]


def _trechos():
    with open(RESUMO, encoding="utf-8") as f:
        t = f.read()
    return {
        # deepClone (cópia de campos) + autoClassifyEntidade + calcularQtdAtivos/extrairFase/isStandaloneLine/recalcAllQtdAtivos
        "funcoes": _fatia(t, "    function deepClone(obj) {\n        const c = {", "    /**\n     * Recalcula qtdAtivos apenas"),
        "computeRowLogic": _fatia(t, "    function computeRowLogic(item) {", "    /* ═══════════════════════════════════════\n       RENDER TABLE"),
        # allProcessed + separação Cabos/Outros/RAMAIS + valores padrão
        # aplicarOperacoesAjuste (usa PREFIXO_TABELA e textoOperacao ao lado)
        "aplicar": _fatia(t, "    const PREFIXO_TABELA = {", "    /** Mostra o diff e deixa escolher"),
        "sync_totalizadora": _fatia(t, "window.syncTotalizadora = async function(forceUpdate = true) {", "    showToast(\"Tabela Totalizadora atualizada!\");\n};") + "\n    showToast(\"x\");\n};",
        "payload": _fatia(t, "    async function obterPayloadCalculo() {", "    /* ═══════════════════════════════════════\n       VALIDAÇÃO DAS PLANILHAS"),
        "ramais_calculo": _fatia(t, "window.recalcularAtivosRamais = function() {", "window.adicionarAtivosRamais = function() {"),
        "ramais_adicionar": _fatia(t, "window.adicionarAtivosRamais = function() {", "window.copyRamaisTable"),
        "separacao": _fatia(t, "    const allProcessed = extractedDataCache.map(", "    buildAtivoSets();"),
    }


_PRE = """
const fs = require('fs');
const entrada = JSON.parse(fs.readFileSync(0, 'utf8'));
globalThis.window = { __regrasLeitorProcessamento: entrada.proc || [], __regrasLeitorClassificacao: entrada.cls || [] };
const Motor = require(%(motor)s);
globalThis.RegrasLeitorEngine = Motor;
let extractedDataCache = entrada.exportados || [];
const tableStates = { cabos: { data: [] }, outros: { data: [] } };
"""


def _rodar(corpo, entrada):
    if not NODE:
        raise RuntimeError("node indisponível")
    codigo = _PRE % {"motor": json.dumps(MOTOR)} + corpo
    r = subprocess.run([NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=120)
    if r.returncode != 0:
        raise RuntimeError(r.stderr)
    return json.loads(r.stdout)


def passo_leitor(itens, proc, cls):
    """script.js (updateRowLogic em cada linha) → o que 'Processar' grava para o Resumo."""
    corpo = """
const itens = entrada.itens;
const saida = itens.map((item, index) => {
  const r = Motor.processarEClassificar({ texto: item.texto, cor: item.cor || '#000000', layer: item.layer || '', pagina: item.pagina,
                                          index: index, allItems: itens }, window.__regrasLeitorProcessamento, window.__regrasLeitorClassificacao);
  return { pagina: item.pagina, texto: item.texto, cor: item.cor, layer: item.layer || '', entidade: r.entidade, operacao: r.operacao, ativo: r.ativo };
});
console.log(JSON.stringify(saida));
"""
    return _rodar(corpo, {"itens": itens, "proc": proc, "cls": cls})


def passo_resumo(exportados, proc, cls):
    """resumo.js: computeRowLogic + separação + recalcAllQtdAtivos, com o código real fatiado do arquivo."""
    tr = _trechos()
    corpo = tr["funcoes"] + "\n" + tr["computeRowLogic"] + "\n" + tr["separacao"] + """
recalcAllQtdAtivos();
console.log(JSON.stringify({ cabos: tableStates.cabos.data, outros: tableStates.outros.data,
  ramais: (window._ramaisData || []).map(r => ({ entidade: r.entidade, texto: r.texto, pagina: r.pagina })) }));
"""
    return _rodar(corpo, {"exportados": exportados, "proc": proc, "cls": cls})


def funcoes_qtd(casos):
    """calcularQtdAtivos/extrairFase/isStandaloneLine do JS real. casos: [{ativo, fase}] → [{qtd, fase, standalone}]."""
    tr = _trechos()
    corpo = tr["funcoes"] + """
const saida = entrada.casos.map(c => ({ qtd: calcularQtdAtivos(c.ativo, c.fase), fase: extrairFase(c.ativo), standalone: isStandaloneLine(c.ativo) }));
console.log(JSON.stringify(saida));
"""
    return _rodar(corpo, {"casos": casos})


def recalcular_lista(linhas):
    """recalcAllQtdAtivos do JS real sobre uma lista de {ativo, qtdAtivos?}."""
    tr = _trechos()
    corpo = tr["funcoes"] + """
tableStates.cabos.data = entrada.linhas;
recalcAllQtdAtivos();
console.log(JSON.stringify(tableStates.cabos.data));
"""
    return _rodar(corpo, {"linhas": linhas})


def aplicar_diff(cabos, outros, operacoes, escolhidas, cls):
    """aplicarOperacoesAjuste do JS real (com a parte de tela trocada por funções vazias)."""
    tr = _trechos()
    corpo = tr["funcoes"] + "\n" + """
const pushHistory = () => {}, renderTable = () => {}, refreshAllFilters = () => {}, buildAtivoSets = () => {}, buildDataLists = () => {}, atualizarResumoRedeUI = () => {};
tableStates.cabos.data = entrada.cabos; tableStates.outros.data = entrada.outros;
""" + tr["aplicar"] + """
const r = aplicarOperacoesAjuste({ operacoes: entrada.operacoes }, entrada.escolhidas === null ? new Set(entrada.operacoes.map((_, i) => i)) : new Set(entrada.escolhidas));
console.log(JSON.stringify({ cabos: tableStates.cabos.data, outros: tableStates.outros.data, aplicadas: r.aplicadas, ignoradas: r.ignoradas }));
"""
    return _rodar(corpo, {"cabos": cabos, "outros": outros, "operacoes": operacoes, "escolhidas": escolhidas, "cls": cls})


def totalizadora_e_payload(cabos, outros, regras, base, projeto_nome):
    """syncTotalizadora + obterPayloadCalculo do JS real. Devolve {totalizadora, payload} (NaN do JS vira null no JSON)."""
    tr = _trechos()
    corpo = tr["funcoes"] + "\n" + """
let orcamentoBaseData = entrada.base;
const localStorage = { getItem: k => (k === 'projeto_selecionado' ? entrada.projeto : null) };
const document = { getElementById: () => null, querySelector: () => ({ checked: true }) };
const confirm = () => true, showToast = () => {}, renderTotalizadora = () => {}, carregarRegras = () => {};
const syncTotalizadora = (...a) => window.syncTotalizadora(...a);
tableStates.regras = { data: entrada.regras }; tableStates.totalizadora = { data: [] };
tableStates.cabos.data = entrada.cabos; tableStates.outros.data = entrada.outros;
""" + tr["sync_totalizadora"] + "\n" + tr["payload"] + """
(async () => {
  const p = await obterPayloadCalculo();
  console.log(JSON.stringify({ totalizadora: tableStates.totalizadora.data, payload: p }));
})();
"""
    return _rodar(corpo, {"cabos": cabos, "outros": outros, "regras": regras, "base": base, "projeto": projeto_nome})


def ramais_para_outros(textos, outros):
    """recalcularAtivosRamais + adicionarAtivosRamais do JS real (DOM trocado por objetos mínimos). Devolve a tabela Outros resultante."""
    tr = _trechos()
    corpo = tr["funcoes"] + "\n" + """
const celulas = entrada.textos.map(t => ({ querySelectorAll: () => [{}, { textContent: t }, {}] }));
const document = { querySelector: () => null, querySelectorAll: () => celulas, getElementById: () => null };
const confirm = () => true, alert = () => {};
const pushHistory = () => {}, renderTable = () => {}, updateCounters = () => {}, updateHistoryUI = () => {}, buildAtivoSets = () => {}, buildDataLists = () => {};
tableStates.outros.data = entrada.outros;
""" + tr["ramais_calculo"] + "\n" + tr["ramais_adicionar"] + """
window.recalcularAtivosRamais();
window.adicionarAtivosRamais();
console.log(JSON.stringify(tableStates.outros.data));
"""
    return _rodar(corpo, {"textos": textos, "outros": outros})
