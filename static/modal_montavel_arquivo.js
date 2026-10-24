/*
 * modal_montavel_arquivo.js (TASK-077, etapa 6) — exportar/importar modais montáveis em arquivo .json.
 * Lógica pura (sem DOM; testada em Node: tests/test_modal_montavel_arquivo.py). Quem baixa/lê o arquivo e salva é modal_montavel_ui.js;
 * o servidor valida de novo ao salvar.
 */
(function (raiz) {
    'use strict';
    const FORMATO = 'modais_montaveis';

    function exportar(projeto, modais) {
        return { formato: FORMATO, versao: 1, projeto: projeto || '', exportado_em: new Date().toISOString(), modais: JSON.parse(JSON.stringify(modais || [])) };
    }

    /** Lê o texto do arquivo. Aceita o arquivo exportado, {versao, modais}, uma lista de modais ou um modal só. Lança Error com mensagem em português. */
    function ler(texto) {
        let d;
        try { d = JSON.parse(texto); } catch (e) { throw new Error('O arquivo não é um JSON válido.'); }
        if (Array.isArray(d)) return d;
        if (d && typeof d === 'object') {
            if (d.formato && d.formato !== FORMATO) throw new Error('Este arquivo é de outro cadastro (' + d.formato + ').');
            if (Array.isArray(d.modais)) return d.modais;
            if (d.blocos && d.nome) return [d];
        }
        throw new Error('Formato não reconhecido: esperado um arquivo exportado pelos modais montáveis.');
    }

    /** modo 'acrescentar': mantém os atuais e junta/atualiza pelo id; 'substituir': a lista fica igual à do arquivo. */
    function resumir(atuais, importados, modo) {
        atuais = atuais || []; importados = importados || [];
        const idsAtuais = new Set(atuais.map(m => m.id)), idsImp = new Set(importados.map(m => m.id));
        const novos = importados.filter(m => !idsAtuais.has(m.id)).map(m => m.nome);
        const atualizados = importados.filter(m => idsAtuais.has(m.id)).map(m => m.nome);
        let resultado, removidos = [];
        if (modo === 'substituir') { resultado = importados.slice(); removidos = atuais.filter(m => !idsImp.has(m.id)).map(m => m.nome); }
        else {
            const porId = new Map(importados.map(m => [m.id, m]));
            resultado = atuais.map(m => porId.has(m.id) ? porId.get(m.id) : m).concat(importados.filter(m => !idsAtuais.has(m.id)));
        }
        return { novos, atualizados, removidos, resultado };
    }

    const api = { exportar, ler, resumir, FORMATO };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.ModalMontavelArquivo = api;
})(typeof window !== 'undefined' ? window : globalThis);
