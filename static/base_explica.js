/*
 * base_explica.js (TASK-179) — a frase em português de cada linha da base, para a ficha: "vale quando cabo contém A35P; soma 3 de 7204 na instalação e na retirada". Sem DOM.
 * Recebe a linha como a ficha a tem (`condicao` "cabo=%A35P%,poste=D", `codigo`, `fator_i`, `fator_r`, `lv`, `mdo`, `filtro`, `itens` dos compostos).
 */
(function (raiz) {
    'use strict';
    const T = t => String(t == null ? '' : t).trim();
    const num = v => { const n = parseFloat(String(v == null ? '' : v).replace(',', '.')); return Number.isFinite(n) ? n : null; };
    const fmt = n => (Number.isInteger(n) ? String(n) : String(n).replace('.', ','));

    /** "%A35P%" → contém A35P · "A35P%" → começa com A35P · "%P" → termina com P · "C" → é C. */
    function valorEmPortugues(alt, nega) {
        const a = T(alt), ini = a.startsWith('%'), fim = a.endsWith('%'), miolo = a.replace(/^%+|%+$/g, '');
        const verbo = ini && fim ? 'contém' : (fim ? 'começa com' : (ini ? 'termina com' : 'é'));
        return nega ? (verbo === 'é' ? 'não é ' : 'não ' + verbo + ' ') + miolo : verbo + ' ' + miolo;
    }

    /** "cabo=%A35P%,poste=D" → "cabo contém A35P e poste é D". */
    function condicaoEmPortugues(cond) {
        return T(cond).split(',').map(p => {
            const nega = p.includes('!='), k = p.indexOf('=');
            if (k < 1) return '';
            const nome = T(p.slice(0, k).replace('!', '')).toLowerCase(), alts = T(p.slice(k + 1)).split('|').map(T).filter(Boolean);
            if (!alts.length) return '';
            return nome + ' ' + (alts.length > 1 ? (nega ? 'não é nenhum de ' : 'é ') + alts.join(' ou ') : valorEmPortugues(alts[0], nega));
        }).filter(Boolean).join(' e ');
    }

    function quantidade(l) {
        const fi = num(l.fator_i), fr = num(l.fator_r), cod = T(l.codigo);
        if (!cod) return '';
        const arred = { CIMA: ' (arredonda para cima a cada uso)', BAIXO: ' (arredonda para baixo a cada uso)', INTEIRO: ' (arredonda a cada uso)' }[T(l.arred).toUpperCase()] || '';
        const inst = (T(l.formula_i) ? 'pela fórmula ' + T(l.formula_i) : (fi !== null && fi > 0 ? fmt(fi) : '')), ret = (T(l.formula_r) ? 'pela fórmula ' + T(l.formula_r) : (fr !== null && fr > 0 ? fmt(fr) : ''));
        const alvo = (T(l.mdo).toUpperCase().indexOf('MAO') === 0 || T(l.mdo).toUpperCase() === 'N') ? 'mão de obra' : 'material';
        let f;
        if (inst && ret) f = inst === ret ? `soma ${inst} de ${cod} (${alvo}) na instalação e na retirada` : `soma ${inst} de ${cod} (${alvo}) na instalação e ${ret} na retirada`;
        else if (inst) f = `soma ${inst} de ${cod} (${alvo}) na instalação (não devolve na retirada)`;
        else if (ret) f = `soma ${ret} de ${cod} (${alvo}) só na retirada`;
        else f = `${cod} não soma nada (fatores zerados)`;
        return f + arred;
    }

    function explicarItem(l) {
        l = l || {};
        const partes = [];
        if (l.itens) {                                                   // composto
            const lista = l.itens.map(i => `${fmt(num(i.qtd) || 1)} × ${T(i.ativo)}`).join(' + ');
            partes.push(`composto: pedir este nome equivale a ${lista}`);
        } else if (T(l.condicao)) {
            partes.push('vale quando ' + condicaoEmPortugues(l.condicao));
        } else if (T(l.tipo) === 'declaracao') {
            partes.push('declara o ativo e seus parâmetros');
        } else {
            partes.push('vale sempre que o ativo é pedido');
        }
        const q = quantidade(l);
        if (q) partes.push(q);
        const lv = T(l.lv).toUpperCase();
        if (lv === 'LV') partes.push('só em linha viva (operação com *)');
        else if (lv === 'LM') partes.push('só em linha morta (operação sem *)');
        if (T(l.filtro)) partes.push('filtro: ' + T(l.filtro));
        return partes.join('; ') + '.';
    }

    const api = { explicarItem, condicaoEmPortugues, valorEmPortugues };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.BaseExplica = api;
})(typeof window !== 'undefined' ? window : globalThis);
