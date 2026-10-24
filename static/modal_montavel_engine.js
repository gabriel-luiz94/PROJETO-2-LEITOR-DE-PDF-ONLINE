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


    /* ───────────── expressões estilo planilha (TASK-077 E5) ─────────────
     * Gramática: comparação ( = <> < > <= >= ) > soma ( + - ) > produto ( * / ) > unário > primário.
     * Primário: número, "texto", [campo], {n} (captura da regex), FUNCAO(a; b; …), ( … ).
     * Funções: SE E OU NAO MAX MIN ABS ARRED SOMA CONCAT MAIUSCULO MINUSCULO. Sem eval, sem acesso a nada fora dos valores. */
    const FUNCOES = { SE: [3, 3], E: [1, 20], OU: [1, 20], NAO: [1, 1], MAX: [1, 20], MIN: [1, 20], ABS: [1, 1], ARRED: [1, 2], SOMA: [1, 20], CONCAT: [1, 20], MAIUSCULO: [1, 1], MINUSCULO: [1, 1] };
    const MAX_EXPR = 300, MAX_PROF = 20;

    function tokenizar(src) {
        const t = []; const re = /\s+|(\d+(?:\.\d+)?|\.\d+)|"((?:[^"]|"")*)"|\[([A-Za-z0-9_-]+)\]|\{(\d+)\}|([A-Za-zÀ-ÿ_]+)|(<=|>=|<>|[-+*\/()=<>;])/y;
        let i = 0;
        while (i < src.length) {
            re.lastIndex = i; const m = re.exec(src);
            if (!m) throw new Error('caractere inesperado em ' + i);
            i = re.lastIndex;
            if (m[1] !== undefined) t.push({ k: 'num', v: parseFloat(m[1]) });
            else if (m[2] !== undefined) t.push({ k: 'str', v: m[2].replace(/""/g, '"') });
            else if (m[3] !== undefined) t.push({ k: 'ref', v: m[3] });
            else if (m[4] !== undefined) t.push({ k: 'cap', v: parseInt(m[4], 10) });
            else if (m[5] !== undefined) t.push({ k: 'id', v: m[5].toUpperCase().replace('Ã', 'A') });
            else if (m[6] !== undefined) t.push({ k: 'op', v: m[6] });
        }
        return t;
    }

    /** Converte texto em árvore (lança Error se inválido). */
    function analisarExpressao(src) {
        src = String(src == null ? '' : src);
        if (!src.trim() || src.length > MAX_EXPR) throw new Error('expressão vazia ou longa demais');
        const t = tokenizar(src); let i = 0;
        const op = v => t[i] && t[i].k === 'op' && t[i].v === v;
        function primario(p) {
            if (p > MAX_PROF) throw new Error('aninhamento demais');
            const x = t[i++];
            if (!x) throw new Error('expressão incompleta');
            if (x.k === 'num' || x.k === 'str' || x.k === 'ref' || x.k === 'cap') return { t: x.k, v: x.v };
            if (x.k === 'op' && x.v === '(') { const e = comp(p + 1); if (!op(')')) throw new Error('falta )'); i++; return e; }
            if (x.k === 'id') {
                const ar = FUNCOES[x.v];
                if (!ar || !op('(')) throw new Error('função desconhecida: ' + x.v);
                i++; const args = [];
                if (!op(')')) { args.push(comp(p + 1)); while (op(';')) { i++; args.push(comp(p + 1)); } }
                if (!op(')')) throw new Error('falta ) em ' + x.v);
                i++;
                if (args.length < ar[0] || args.length > ar[1]) throw new Error(x.v + ' com ' + args.length + ' argumento(s)');
                return { t: 'fn', n: x.v, a: args };
            }
            throw new Error('símbolo inesperado');
        }
        function unario(p) { if (op('-')) { i++; return { t: 'neg', a: unario(p + 1) }; } if (op('+')) { i++; return unario(p + 1); } return primario(p); }
        function produto(p) { let e = unario(p); while (op('*') || op('/')) { const o = t[i++].v; e = { t: 'bin', o, a: e, b: unario(p) }; } return e; }
        function soma(p) { let e = produto(p); while (op('+') || op('-')) { const o = t[i++].v; e = { t: 'bin', o, a: e, b: produto(p) }; } return e; }
        function comp(p) {
            const e = soma(p);
            if (t[i] && t[i].k === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(t[i].v)) { const o = t[i++].v; return { t: 'bin', o, a: e, b: soma(p) }; }
            return e;
        }
        const r = comp(0);
        if (i !== t.length) throw new Error('sobrou texto no fim');
        return r;
    }

    /** Ids [campo] citados e maior {n} usado — para o servidor/editor conferirem. */
    function referencias(no, acc) {
        acc = acc || { refs: new Set(), caps: new Set() };
        if (!no) return acc;
        if (no.t === 'ref') acc.refs.add(no.v);
        if (no.t === 'cap') acc.caps.add(no.v);
        ['a', 'b'].forEach(k => { if (no[k] && !Array.isArray(no[k])) referencias(no[k], acc); });
        if (Array.isArray(no.a)) no.a.forEach(x => referencias(x, acc));
        return acc;
    }

    const ehNumTexto = v => typeof v === 'string' && /^\s*-?\d+([.,]\d+)?\s*$/.test(v);
    const paraNum = v => typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : (parseFloat(String(v == null ? '' : v).replace(',', '.')) || 0);
    const verdade = v => typeof v === 'boolean' ? v : typeof v === 'number' ? v !== 0 : String(v == null ? '' : v).trim() !== '';
    const paraTexto = v => typeof v === 'number' ? (Number.isInteger(v) ? String(v) : String(parseFloat(v.toFixed(6)))) : typeof v === 'boolean' ? (v ? 'VERDADEIRO' : 'FALSO') : String(v == null ? '' : v);
    const valorDeCampo = v => v == null ? '' : (ehNumTexto(v) ? parseFloat(String(v).replace(',', '.')) : v);

    function avaliarNo(no, ctx) {
        switch (no.t) {
            case 'num': case 'str': return no.v;
            case 'ref': return Object.prototype.hasOwnProperty.call(ctx.valores || {}, no.v) ? valorDeCampo(ctx.valores[no.v]) : '';
            case 'cap': { const c = ctx.caps && ctx.caps[no.v]; return c == null ? '' : valorDeCampo(c); }
            case 'neg': return -paraNum(avaliarNo(no.a, ctx));
            case 'bin': {
                const a = avaliarNo(no.a, ctx), b = avaliarNo(no.b, ctx);
                switch (no.o) {
                    case '+': return paraNum(a) + paraNum(b);
                    case '-': return paraNum(a) - paraNum(b);
                    case '*': return paraNum(a) * paraNum(b);
                    case '/': { const d = paraNum(b); if (d === 0) throw new Error('divisão por zero'); return paraNum(a) / d; }
                    default: {
                        const num = (typeof a === 'number' || ehNumTexto(a)) && (typeof b === 'number' || ehNumTexto(b));
                        const x = num ? paraNum(a) : paraTexto(a).trim().toLowerCase(), y = num ? paraNum(b) : paraTexto(b).trim().toLowerCase();
                        switch (no.o) { case '=': return x === y; case '<>': return x !== y; case '<': return x < y; case '>': return x > y; case '<=': return x <= y; default: return x >= y; }
                    }
                }
            }
            default: {
                const args = no.a.map(x => x);
                const v = i => avaliarNo(args[i], ctx);
                switch (no.n) {
                    case 'SE': return verdade(v(0)) ? v(1) : v(2);                 // só avalia o ramo escolhido
                    case 'E': return args.every((_, i) => verdade(v(i)));
                    case 'OU': return args.some((_, i) => verdade(v(i)));
                    case 'NAO': return !verdade(v(0));
                    case 'MAX': return Math.max.apply(null, args.map((_, i) => paraNum(v(i))));
                    case 'MIN': return Math.min.apply(null, args.map((_, i) => paraNum(v(i))));
                    case 'ABS': return Math.abs(paraNum(v(0)));
                    case 'ARRED': { const n = args.length > 1 ? Math.max(0, Math.min(6, Math.round(paraNum(v(1))))) : 0; const f = Math.pow(10, n); return Math.round(paraNum(v(0)) * f) / f; }
                    case 'SOMA': return args.reduce((s, _, i) => s + paraNum(v(i)), 0);
                    case 'CONCAT': return args.map((_, i) => paraTexto(v(i))).join('');
                    case 'MAIUSCULO': return paraTexto(v(0)).toUpperCase();
                    default: return paraTexto(v(0)).toLowerCase();
                }
            }
        }
    }

    /** Avalia uma expressão (texto) com os valores dos campos e as capturas; devolve o valor (número/texto/booleano) ou null se inválida. */
    function avaliar(src, valores, caps) {
        try { return avaliarNo(analisarExpressao(src), { valores, caps }); } catch (e) { return null; }
    }

    /** Calcula os blocos `calculo` na ordem; devolve uma cópia dos valores com cada id calculado (número ou texto). */
    function calcularCampos(blocos, valores) {
        const v = Object.assign({}, valores);
        (blocos || []).forEach(b => {
            if (b.tipo !== 'calculo') return;
            const r = avaliar(b.expressao, v, []);
            v[b.id] = r === null ? '' : (typeof r === 'boolean' ? paraTexto(r) : r);
        });
        return v;
    }
    const formatar = paraTexto;

    /** Troca [campo] pelo valor do campo (numérico: número ou 0; texto: o texto). Campo desconhecido fica como está. */
    function substituirCampos(texto, valores, numerico) {
        return String(texto == null ? '' : texto).replace(/\[([A-Za-z0-9_-]+)\]/g, (m, id) => {
            if (!valores || !Object.prototype.hasOwnProperty.call(valores, id)) return m;
            const v = valores[id];
            if (numerico) { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isNaN(x) ? '(0)' : '(' + x + ')'; }
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
    function avaliarExpressao(expr, caps, valores) {
        if (typeof expr === 'number') return expr;
        const r = avaliar(expr, valores, caps || []);
        if (r === null || r === '') return null;
        const n = typeof r === 'number' ? r : typeof r === 'boolean' ? (r ? 1 : 0) : (ehNumTexto(r) ? paraNum(r) : null);
        return n === null || !isFinite(n) ? null : n;
    }

    function regexDe(q) {
        try { return new RegExp(q.regex, q.flags == null ? 'i' : q.flags); } catch (e) { return null; }
    }

    /**
     * Aplica as regras de um bloco `interpretar` às linhas da grade (cada linha = objeto coluna→texto).
     * Cada regra: { quando: {coluna?, regex, flags?}, gera: [{ativo, qtd?, operacao?}], parar? }.
     * Sem `coluna`, a regex testa todas as colunas juntas (separadas por espaço). Retorna itens somados.
     */
    function executar(bloco, linhas, valores) {
        const soma = new Map(); const ordem = [];
        const agrupar = bloco.agrupar !== false;
        const itens = [], naoCasadas = [];
        const caixa = t => bloco.normalizar === 'maiusculo' ? t.toUpperCase() : bloco.normalizar === 'minusculo' ? t.toLowerCase() : t;
        (linhas || []).forEach((linha, idx) => {
            const textoTudo = Object.keys(linha || {}).map(k => linha[k]).join(' ');
            let casou = false;
            for (const regra of (bloco.regras || [])) {
                if (regra.se && !avaliarCondicao(regra.se, valores)) continue;   // regra só vale se a condição dos campos for verdadeira
                const q = regra.quando || {};
                const re = regexDe(q);
                if (!re) continue;
                const alvo = q.coluna ? String((linha || {})[q.coluna] == null ? '' : linha[q.coluna]) : textoTudo;
                let ocorrencias;
                if (q.global) {                                       // todas as ocorrências na linha ("P1, P2 e P3")
                    const rg = new RegExp(re.source, re.flags.indexOf('g') >= 0 ? re.flags : re.flags + 'g');
                    ocorrencias = []; let mm, guarda = 0;
                    while ((mm = rg.exec(alvo)) !== null && guarda++ < 200) { ocorrencias.push(mm); if (mm[0] === '') rg.lastIndex++; }
                } else { const m = re.exec(alvo); ocorrencias = m ? [m] : []; }
                if (!ocorrencias.length) continue;
                casou = true;
                ocorrencias.forEach(m => {
                    const caps = Array.prototype.slice.call(m, 0, MAX_CAP);
                    (regra.gera || []).forEach(g => {
                        const ativo = caixa(substituir(substituirCampos(g.ativo, valores, false), caps, false).trim());
                        const qtd = g.qtd === undefined ? 1 : avaliarExpressao(g.qtd, caps, valores);
                        if (!ativo || qtd === null || qtd <= 0) return;
                        const operacao = g.operacao || 'I';
                        if (!agrupar) { itens.push({ ativo, qtd, operacao, linha: idx }); return; }
                        const k = operacao + '\u0000' + ativo;
                        if (!soma.has(k)) { soma.set(k, { ativo, qtd: 0, operacao }); ordem.push(k); }
                        soma.get(k).qtd += qtd;
                    });
                });
                if (regra.parar !== false) break;   // por padrão a primeira regra que casa vale (como as regras do leitor)
            }
            if (!casou && textoTudo.trim() !== '') naoCasadas.push({ linha: idx, texto: textoTudo.trim() });
        });
        return { itens: agrupar ? ordem.map(k => soma.get(k)) : itens, naoCasadas };
    }
    function interpretar(bloco, linhas, valores) { return executar(bloco, linhas, valores).itens; }
    /** Mesmo que interpretar, mas também devolve as linhas não vazias que nenhuma regra reconheceu. */
    function diagnosticar(bloco, linhas, valores) { return executar(bloco, linhas, valores); }

    /**
     * Ajusta a matriz colada (TSV) ao modo da grade: bloco.colagem = { modo: 'colunas'|'linha'|'dividir', separador, pular_cabecalho }.
     * 'linha': a linha inteira vai para a 1ª coluna; 'dividir': a linha inteira é dividida pelo separador (padrão ';') nas colunas.
     */
    function ajustarColagem(bloco, matriz) {
        const c = bloco.colagem || {};
        let m = (matriz || []).map(l => l.slice());
        if (c.pular_cabecalho) m = m.slice(1);
        if (c.modo === 'linha') return m.map(l => [l.join(' ').trim()]);
        if (c.modo === 'dividir') {
            const sep = c.separador == null || c.separador === '' ? ';' : c.separador;
            return m.map(l => l.join(' ').split(sep).map(x => x.trim()));
        }
        return m;
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
            if (expr.charAt(0) === '=') {                                  // {{= expressão }}: conta/texto com [campo] e [coluna]
                const r = avaliar(expr.slice(1), Object.assign({}, ctx.valores, ctx.linha), []);
                return r === null ? '' : formatar(r);
            }
            const f = /^buscar\s*\(([^()]*)\)$/.exec(expr);
            if (f) {
                const a = f[1].split(',');
                return a.length === 4 ? buscar(ctx, a[0], a[1], a[2], a[3]) : '';
            }
            const v = resolver(expr, ctx);
            return v == null ? '' : String(v);
        });
    }

    const num = v => { const x = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isNaN(x) ? null : x; };

    /**
     * Condição: folha { campo, op, valor } (op: igual, diferente, vazio, preenchido, em, contem, maior, menor, maior_igual, menor_igual)
     * ou composta { todos: [...] } (E) / { algum: [...] } (OU) / { nao: cond }. Sem condição = verdadeira.
     */
    function avaliarCondicao(c, valores) {
        if (!c) return true;
        if (Array.isArray(c.todos)) return c.todos.every(x => avaliarCondicao(x, valores));
        if (Array.isArray(c.algum)) return c.algum.some(x => avaliarCondicao(x, valores));
        if (c.nao) return !avaliarCondicao(c.nao, valores);
        const v = String((valores || {})[c.campo] == null ? '' : valores[c.campo]).trim();
        switch (c.op) {
            case 'igual': return v === String(c.valor);
            case 'diferente': return v !== String(c.valor);
            case 'vazio': return v === '';
            case 'preenchido': return v !== '';
            case 'em': return Array.isArray(c.valor) && c.valor.map(String).includes(v);
            case 'contem': return v.toLowerCase().indexOf(String(c.valor).toLowerCase()) >= 0;
            case 'maior': case 'menor': case 'maior_igual': case 'menor_igual': {
                const a = num(v), b = num(c.valor);
                if (a === null || b === null) return false;
                return c.op === 'maior' ? a > b : c.op === 'menor' ? a < b : c.op === 'maior_igual' ? a >= b : a <= b;
            }
            default: return true;
        }
    }

    /** visivel_se de um bloco (mesma linguagem de condições). */
    function visivel(bloco, valores) { return avaliarCondicao(bloco && bloco.visivel_se, valores); }

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

    const api = { calcular, avaliar, analisarExpressao, referencias, calcularCampos, formatar, avaliarExpressao, interpretar, diagnosticar, ajustarColagem, modelo, visivel, avaliarCondicao, linhasDaMatriz, gerarTexto, substituir, substituirCampos };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.ModalMontavel = api;
})(typeof window !== 'undefined' ? window : globalThis);
