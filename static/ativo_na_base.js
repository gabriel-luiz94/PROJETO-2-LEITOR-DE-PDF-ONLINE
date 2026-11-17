/*
 * ativo_na_base.js — "o ativo existe na base?" na tela (Cabos e Outros). Sem DOM.
 * Mesma cascata do cálculo para o NOME do ativo (services/validacao_planilhas.nomes_conhecidos): ATIVO (e apelidos) → COMPONENTE → CÓDIGO, só linhas que valem no projeto
 * e na origem (CABOS/OUTROS). Os parâmetros não entram (isso é o C1-PARAM). Quem usa: a cor vermelha do interruptor "Colorir ativos por parâmetros" no Resumo.
 * O que a tela confere é o texto digitado; as regras de conversão da Totalizadora podem trocar um ativo — o Validar (C1-BASE) confere depois da conversão.
 */
(function (raiz) {
    'use strict';
    const AP = (typeof module !== 'undefined' && module.exports) ? require('./ativo_parametros.js') : raiz.AtivoParametros;
    const up = t => String(t == null ? '' : t).trim().toUpperCase();

    function valeNoProjeto(row, projeto) {
        const p = up(row && row.projeto);
        return !p || !projeto || p.split('/').map(up).includes(projeto);
    }

    /** Set (MAIÚSCULAS) dos nomes que o cálculo resolve na base para o projeto e a origem. */
    function nomesDaBase(base, projeto, origem) {
        const proj = up(projeto), org = up(origem), nomes = new Set();
        (base || []).forEach(r => {
            if (!r || !valeNoProjeto(r, proj)) return;
            const ro = up(r.origem);
            if (ro && ro !== org) return;
            const cel = String(r.ativo || '').trim();
            if (cel) {
                const a = AP.analisarCelulaBase(cel);
                if (!a.erro && a.base) { nomes.add(up(a.base)); (a.apelidos || []).forEach(x => nomes.add(up(x))); } else nomes.add(up(cel.split('#')[0]));
            }
            const comp = up(r.componente); if (comp) nomes.add(comp);
            const cod = up(r.codigo); if (cod) nomes.add(cod);
        });
        return nomes;
    }

    /** Dos `ativos` pedidos (texto, com ou sem parâmetros), os que NÃO existem em `nomes`. */
    function foraDaBase(nomes, ativos) {
        return (ativos || []).filter(a => { const n = up(AP.nomeBase(a)); return n !== '' && !nomes.has(n); });
    }

    const api = { nomesDaBase, foraDaBase, valeNoProjeto };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.AtivoNaBase = api;
})(typeof window !== 'undefined' ? window : globalThis);
