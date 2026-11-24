/*
 * cabo_autocompletar.js (TASK-243) — pré-preenchimento do ativo do CABO enquanto se digita (formulários de cabo do Mapa).
 *   • «3»            → «3x»
 *   • «3xA35P(»      → «3xA35P(7,94)-13,8kV»   (a continuação mais comum desse cabo na obra; sem exemplo, cabo protegido A..P → «7,94)-13,8kV»)
 *   • «3xC25-»       → «3xC25-13,8kV»          (a tensão mais comum na obra; sem exemplo, «13,8kV»)
 * A sugestão entra escrita e SELECIONADA depois do cursor: continuar digitando a substitui, Tab / → / End aceitam, Backspace apaga.
 * Parte pura (sem DOM; testada em Node: tests/test_cabo_autocompletar_243.py): sugerir(textoAteOCursor, exemplos).
 */
(function (raiz) {
    'use strict';
    const PROTEGIDO_PADRAO = '7,94)-13,8kV';
    const TENSAO_PADRAO = '13,8kV';
    const semMetragem = t => String(t == null ? '' : t).replace(/\s+\d+(?:[.,]\d+)?\s*m\.?\s*$/i, '').trim();

    /** O valor mais frequente (empate: o primeiro que apareceu). */
    function maisComum(lista) {
        const contagem = new Map();
        lista.forEach(v => contagem.set(v, (contagem.get(v) || 0) + 1));
        let melhor = '', n = 0;
        contagem.forEach((c, v) => { if (c > n) { melhor = v; n = c; } });
        return melhor;
    }

    /** O que completar depois de `antes` (o texto até o cursor, no fim do campo). '' = nada a sugerir. */
    function sugerir(antes, exemplos) {
        const t = String(antes == null ? '' : antes);
        if (!t) return '';
        const ex = (exemplos || []).map(semMetragem).filter(Boolean);
        if (/^\d$/.test(t)) return 'x';                                                       // «3» → «3x»
        const abre = /([A-Z]+\d+[A-Z]*)\($/i.exec(t.replace(/^\s*\d+\s*x/i, ''));   // «3xA35P(» → código A35P (sem o «3x»)
        if (abre) {
            const cod = abre[1].toUpperCase() + '(';
            const continuacoes = ex.map(e => { const i = e.toUpperCase().lastIndexOf(cod); return i < 0 ? '' : e.slice(i + cod.length); }).filter(Boolean);
            if (continuacoes.length) return maisComum(continuacoes);
            return /^A\d+P$/i.test(abre[1]) ? PROTEGIDO_PADRAO : '';
        }
        if (/(?:\)|[A-Z]\d+[A-Z]*)-$/i.test(t) && !/^\s*(?:\d+x)?(?:PB|RAMA)/i.test(t)) {      // «…)-» ou «…C25-» → a tensão (não em multiplex de BT)
            const tensoes = ex.map(e => { const m = /-\s*(\d+(?:[.,]\d+)?\s*kV)\s*$/i.exec(e); return m ? m[1] : ''; }).filter(Boolean);
            return tensoes.length ? maisComum(tensoes) : TENSAO_PADRAO;
        }
        return '';
    }

    /** Liga o pré-preenchimento a um campo de texto; `exemplos()` devolve os ativos de cabo já existentes (para aprender o padrão da obra). */
    function ligar(input, exemplos) {
        if (!input || input.__caboAutocompletar) return input;
        input.__caboAutocompletar = true;
        input.addEventListener('input', e => {
            if (e.inputType !== 'insertText' || !e.data) return;                               // apagar ou colar não sugere
            const v = input.value, c = input.selectionStart;
            if (c !== v.length || input.selectionEnd !== c) return;                            // só no fim do texto
            const s = sugerir(v, typeof exemplos === 'function' ? exemplos() : exemplos);
            if (!s) return;
            input.value = v + s;
            input.setSelectionRange(v.length, v.length + s.length);
        });
        input.addEventListener('keydown', e => {
            const temSugestao = input.selectionStart < input.selectionEnd && input.selectionEnd === input.value.length;
            if (!temSugestao) return;
            if ((e.key === 'Tab' && !e.shiftKey) || e.key === 'End') {
                e.preventDefault();
                input.setSelectionRange(input.value.length, input.value.length);
            }
        });
        return input;
    }

    const api = { sugerir, ligar, maisComum };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    raiz.CaboAutocompletar = api;
})(typeof window !== 'undefined' ? window : globalThis);
