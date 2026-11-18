/*
 * ativo_compacto.js — texto "compacto" do ativo em Outros: tira os parâmetros entre parênteses que vêm colados no nome (PR(D,X) → PR; CE_4(12600D,A35P) → CE_4).
 * Só aparência (a opção "Ocultar parâmetros ao sair da linha"): o texto da linha e o cálculo continuam com os parâmetros. Sem DOM; testado em Node.
 * Um "(" sem ")" correspondente (o usuário ainda está digitando) deixa o texto como está.
 */
(function (raiz) {
    'use strict';
    /** "RL, PR(D,X), ATRINT" → "RL, PR, ATRINT". Só tira o par de parênteses que começa colado a um nome (não depois de espaço, vírgula ou no início). */
    function compactar(texto) {
        const t = String(texto == null ? '' : texto);
        if (t.indexOf('(') === -1) return t;
        let saida = '';
        for (let i = 0; i < t.length; i++) {
            const c = t[i];
            if (c === '(' && i > 0 && !/[\s,]/.test(t[i - 1])) {
                let prof = 0, j = i;
                for (; j < t.length; j++) { if (t[j] === '(') prof++; else if (t[j] === ')' && --prof === 0) break; }
                if (j >= t.length) return t;                           // parêntese aberto: ainda digitando
                i = j;
                continue;
            }
            saida += c;
        }
        return saida;
    }
    const api = { compactar };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.AtivoCompacto = api;
})(typeof window !== 'undefined' ? window : globalThis);
