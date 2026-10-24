/*
 * modal_montavel_engine.js (TASK-077) — lógica pura dos "modais montáveis" (sem DOM; testada em Node: tests/test_modal_montavel.py).
 * Um modal é um JSON declarativo (validado no servidor, routers/modais_montaveis.py). Este arquivo só INTERPRETA a definição:
 *   - expressões aritméticas seguras (sem eval): "{1}*20 + 2"
 *   - interpretar(): linhas coladas + regras (regex) -> itens {ativo, qtd, operacao}, somados por (operação, ativo)
 *   - modelo(): "{{campo}}" e "{{buscar(tabela, colunaChave, valor, colunaRetorno)}}" -> texto
 *   - visivel(): condição simples de um bloco (visivel_se)
 * Nada aqui executa código do usuário: só regex, números e as quatro operações.
 */
(function (raiz) {
    'use strict';

    const MAX_CAP = 20;

    /** Substitui {0}..{N} pelas capturas (números → número; vazio → 0 no modo numérico, '' no modo texto). */
    function substituir(texto, caps, numerico) {
        return String(texto == null ? '' : texto).replace(/\{(\d+)\}/g, (_, n) => {
            const v = caps && caps[Number(n)];
            if (numerico) { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isNaN(x) ? '0' : '(' + x + ')'; }
            return v == null ? '' : String(v);
        });
    }

    /** Avalia só + - * / ( ) e números (descida recursiva). Retorna número ou null se a expressão for inválida. */
    function calcular(expr) {
        const s = String(expr).replace(/\s+/g, '');
        if (!s || !/^[0-9+\-*/().]+$/.test(s)) return null;
        let i = 0;
        function fator() {
            if (s[i] === '-') { i++; const v = fator(); return v === null ? null : -v; }
            if (s[i] === '+') { i++; return fator(); }
            if (s[i] === '(') { i++; const v = soma(); if (v === null || s[i] !== ')') return null; i++; return v; }
            const m = /^\d+(\.\d+)?|^\.\d+/.exec(s.slice(i));
            if (!m) return null;
            i += m[0].length;
            return parseFloat(m[0]);
        }
        function termo() {
            let v = fator();
            while (v !== null && (s[i] === '*' || s[i] === '/')) {
                const op = s[i++]; const d = fator();
                if (d === null) return null;
                v = op === '*' ? v * d : (d === 0 ? null : v / d);
            }
            return v;
        }
        function soma() {
            let v = termo();
            while (v !== null && (s[i] === '+' || s[i] === '-')) {
                const op = s[i++]; const d = termo();
                if (d === null) return null;
                v = op === '+' ? v + d : v - d;
            }
            return v;
        }
        const r = soma();
        return i === s.length && r !== null && isFinite(r) ? r : null;
    }

    /** Expressão numérica com capturas: avaliarExpressao("{1}*20", ["x","3"]) = 60. Número puro também vale. */
    function avaliarExpressao(expr, caps) {
        if (typeof expr === 'number') return expr;
        return calcular(substituir(expr, caps, true));
    }

    function regexDe(q) {
        try { return new RegExp(q.regex, q.flags == null ? 'i' : q.flags); } catch (e) { return null; }
    }

    /**
     * Aplica as regras de um bloco `interpretar` às linhas da grade (cada linha = objeto coluna→texto).
     * Cada regra: { quando: {coluna?, regex, flags?}, gera: [{ativo, qtd?, operacao?}], parar? }.
     * Sem `coluna`, a regex testa todas as colunas juntas (separadas por espaço). Retorna itens somados.
     */
    function interpretar(bloco, linhas) {
        const soma = new Map(); const ordem = [];
        const agrupar = bloco.agrupar !== false;
        const itens = [];
        (linhas || []).forEach((linha, idx) => {
            const textoTudo = Object.keys(linha || {}).map(k => linha[k]).join(' ');
            for (const regra of (bloco.regras || [])) {
                const q = regra.quando || {};
                const re = regexDe(q);
                if (!re) continue;
                const alvo = q.coluna ? String((linha || {})[q.coluna] == null ? '' : linha[q.coluna]) : textoTudo;
                const m = re.exec(alvo);
                if (!m) continue;
                const caps = Array.prototype.slice.call(m, 0, MAX_CAP);
                (regra.gera || []).forEach(g => {
                    const ativo = substituir(g.ativo, caps, false).trim();
                    const qtd = g.qtd === undefined ? 1 : avaliarExpressao(g.qtd, caps);
                    if (!ativo || qtd === null || qtd <= 0) return;
                    const operacao = g.operacao || 'I';
                    if (!agrupar) { itens.push({ ativo, qtd, operacao, linha: idx }); return; }
                    const k = operacao + '\u0000' + ativo;
                    if (!soma.has(k)) { soma.set(k, { ativo, qtd: 0, operacao }); ordem.push(k); }
                    soma.get(k).qtd += qtd;
                });
                if (regra.parar !== false) break;   // por padrão a primeira regra que casa vale (como as regras do leitor)
            }
        });
        return agrupar ? ordem.map(k => soma.get(k)) : itens;
    }

    /** Acha o valor de `nome` em: valores da linha, valores dos campos. */
    function resolver(nome, ctx) {
        const n = String(nome).trim();
        if (/^'.*'$/.test(n) || /^".*"$/.test(n)) return n.slice(1, -1);
        if (ctx.linha && Object.prototype.hasOwnProperty.call(ctx.linha, n)) return ctx.linha[n];
        if (ctx.valores && Object.prototype.hasOwnProperty.call(ctx.valores, n)) return ctx.valores[n];
        return '';
    }

    function buscar(ctx, tabelaId, colChave, valor, colRet) {
        const t = (ctx.tabelas || {})[String(tabelaId).trim()];
        if (!t) return '';
        const ci = t.colunas.indexOf(String(colChave).trim()), cr = t.colunas.indexOf(String(colRet).trim());
        if (ci < 0 || cr < 0) return '';
        const v = String(resolver(valor, ctx)).trim().toLowerCase();
        const lin = t.linhas.find(l => String(l[ci] == null ? '' : l[ci]).trim().toLowerCase() === v);
        return lin ? String(lin[cr] == null ? '' : lin[cr]) : '';
    }

    /** Preenche "{{campo}}" e "{{buscar(tabela, colunaChave, valor, colunaRetorno)}}". Desconhecido vira vazio. */
    function modelo(texto, ctx) {
        return String(texto == null ? '' : texto).replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (_, expr) => {
            const f = /^buscar\s*\(([^()]*)\)$/.exec(expr);
            if (f) {
                const a = f[1].split(',');
                return a.length === 4 ? buscar(ctx, a[0], a[1], a[2], a[3]) : '';
            }
            const v = resolver(expr, ctx);
            return v == null ? '' : String(v);
        });
    }

    /** visivel_se: { campo, op: igual|diferente|vazio|preenchido|em, valor } sobre os valores dos campos. */
    function visivel(bloco, valores) {
        const c = bloco && bloco.visivel_se;
        if (!c) return true;
        const v = String((valores || {})[c.campo] == null ? '' : valores[c.campo]).trim();
        switch (c.op) {
            case 'igual': return v === String(c.valor);
            case 'diferente': return v !== String(c.valor);
            case 'vazio': return v === '';
            case 'preenchido': return v !== '';
            case 'em': return Array.isArray(c.valor) && c.valor.map(String).includes(v);
            default: return true;
        }
    }

    /** Linhas de uma grade a partir de uma matriz colada (parseColagem de grade_tabela.js): coluna i → colunas[i].id. */
    function linhasDaMatriz(colunas, matriz) {
        return (matriz || []).map(cel => {
            const o = {};
            colunas.forEach((c, i) => { o[c.id] = (cel[i] == null ? '' : String(cel[i])).trim(); });
            return o;
        });
    }

    /** Texto final do bloco `saida_texto`: com `fonte` (grade) aplica o modelo a cada linha; sem, uma vez. */
    function gerarTexto(bloco, valores, tabelas, linhasFonte) {
        const base = { valores, tabelas };
        if (bloco.fonte) {
            return (linhasFonte || []).filter(l => Object.keys(l).some(k => String(l[k]).trim() !== ''))
                .map(l => modelo(bloco.modelo, Object.assign({ linha: l }, base))).join('\n');
        }
        return modelo(bloco.modelo, base);
    }

    const api = { calcular, avaliarExpressao, interpretar, modelo, visivel, linhasDaMatriz, gerarTexto, substituir };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.ModalMontavel = api;
})(typeof window !== 'undefined' ? window : globalThis);
