/*
 * anotacao_outros.js (TASK-081) — notação compacta na tabela Outros: "11600D, 2 N3, T3" (itens separados por ", ").
 * MESMA lógica de services/anotacao_outros.py (paridade testada em tests/test_anotacao_outros.py). Sem DOM. Depende de AtivoParametros.
 * O usuário vê o que digitou; a tradução para o formato antigo é implícita (interpretarLinha). O "classico" é o leitor do formato
 * ANTIGO de cada chamador: recebe o texto e devolve [{ativo, qtd}].
 */
(function (raiz) {
    'use strict';
    const AP = (typeof module !== 'undefined' && module.exports) ? require('./ativo_parametros.js') : raiz.AtivoParametros;
    const SUFIXO_PARA_TIPO = { D: 'DT', R: 'RC', F: 'CV', FV: 'FV' };
    const TIPO_PARA_SUFIXO = { DT: 'D', RC: 'R', CV: 'F', FV: 'FV' };
    const ALTURA_MIN = 7, ALTURA_MAX = 16;
    const RE_COMPACTO = /^(\d+)(FV|D|R|F)$/i;
    const RE_POSTE = /^(DT|RC|CV|FV)(\d+)\/(\d+)$/;
    const RE_QTD = /^\*?\d+(?:[.,]\d+)?$/;
    const RE_QTD_ANTIGA = /^\*?\d+(?:[.,]\d+)?[-xX][A-Za-z(].*$/;
    const txt = t => String(t == null ? '' : t);

    function ehPoste(nome) { return RE_POSTE.test(txt(nome).trim().toUpperCase()); }

    function posteParaCompacto(nome) {
        const m = txt(nome).trim().toUpperCase().match(RE_POSTE);
        return m ? `${m[2]}${m[3]}${TIPO_PARA_SUFIXO[m[1]]}` : null;
    }

    function compactoParaPoste(token, postes) {
        const t = txt(token).trim(), m = t.match(RE_COMPACTO);
        if (!m) return null;
        const digitos = m[1], tipo = SUFIXO_PARA_TIPO[m[2].toUpperCase()], cands = [];
        for (const hl of [2, 1]) {
            const h = digitos.slice(0, hl), e = digitos.slice(hl);
            if (e.length < 3 || e.length > 4 || h.startsWith('0') || e.startsWith('0')) continue;
            cands.push([parseInt(h, 10), `${tipo}${parseInt(h, 10)}/${parseInt(e, 10)}`]);
        }
        const r = { ativo: null, compacto: t.toUpperCase(), erro: null };
        if (postes && postes.size) {
            const naBase = cands.map(c => c[1]).filter(n => postes.has(n));
            if (naBase.length === 1) { r.ativo = naBase[0]; return r; }
        }
        const plausiveis = cands.filter(c => c[0] >= ALTURA_MIN && c[0] <= ALTURA_MAX).map(c => c[1]);
        if (plausiveis.length === 1) r.ativo = plausiveis[0];
        else r.erro = !plausiveis.length ? `não foi possível separar altura e esforço de '${t.toUpperCase()}'` : `'${t.toUpperCase()}' é ambíguo (${plausiveis.join(' ou ')})`;
        return r;
    }

    /** Vírgulas fora de parênteses: [[posição, 'separador'|'decimal'|'invalida']]. */
    function virgulas(texto) {
        const saida = []; let prof = 0;
        for (let i = 0; i < texto.length; i++) {
            const ch = texto[i];
            if (ch === '(') prof++;
            else if (ch === ')' && prof > 0) prof--;
            else if (ch === ',' && prof === 0) {
                const prox = i + 1 < texto.length ? texto[i + 1] : '', ant = i > 0 ? texto[i - 1] : '';
                if (prox === '' || /\s/.test(prox)) saida.push([i, 'separador']);
                else if (/\d/.test(ant) && /\d/.test(prox)) saida.push([i, 'decimal']);
                else saida.push([i, 'invalida']);
            }
        }
        return saida;
    }

    function ehCompacta(texto) { return virgulas(txt(texto)).some(v => v[1] === 'separador'); }

    function resolverNome(nome, ctx) {
        nome = txt(nome).trim().toUpperCase();
        const i = nome.indexOf('('), base = i !== -1 ? nome.slice(0, i) : nome, params = i !== -1 ? nome.slice(i) : '';
        ctx = ctx || {};
        if (ctx.nomes && ctx.nomes.has(base)) return [base + params, null];
        if (ctx.apelidos && Object.prototype.hasOwnProperty.call(ctx.apelidos, base)) return [ctx.apelidos[base] + params, null];
        const c = compactoParaPoste(base, ctx.postes);
        if (c !== null) return c.erro ? [base + params, c.erro] : [c.ativo + params, null];
        return [base + params, null];
    }

    function qtdDe(t) {
        t = t.trim();
        const neg = t.startsWith('*'), v = parseFloat((neg ? t.slice(1) : t).replace(',', '.'));
        return neg ? -v : v;
    }

    /** Acrescenta NOME=valor ao pedido quando a assinatura tem o parâmetro e o pedido não o informou (N1 ou N1(A)). */
    function comParam(ativo, nome, valor, assinaturas) {
        const i = ativo.indexOf('('), base = i !== -1 ? ativo.slice(0, i) : ativo, sig = assinaturas[base];
        if (!sig || !sig.nomes.includes(nome)) return ativo;
        if (i === -1) return `${ativo}(${nome}=${valor})`;
        if (!ativo.endsWith(')')) return ativo;
        const itens = ativo.slice(i + 1, -1).split(',').map(x => x.trim()).filter(x => x);
        if (itens.some(x => x.toUpperCase().startsWith(nome + '='))) return ativo;
        if (!itens.some(x => x.includes('=')) && itens.length > sig.nomes.indexOf(nome)) return ativo;
        return `${ativo.slice(0, -1)},${nome}=${valor})`;
    }

    function resolveParaPoste(token, ctx) {
        const [nome, erro] = resolverNome(token, ctx);
        return erro === null && ehPoste(nome) && !/^\*?\d+(?:[.,]\d+)?[-xX]/.test(token);
    }

    function interpretarLinha(texto, operacao, classico, ctx) {
        const t = txt(texto).trim(), op = txt(operacao).trim().toUpperCase();
        const r = { formato: 'antigo', itens: [], erros: [] };
        if (!t) return r;
        const vs = virgulas(t);
        if (vs.some(v => v[1] === 'invalida')) {
            r.erros.push({ cod: 'VIRGULA', msg: `vírgula sem espaço em '${t}': use vírgula + espaço para separar itens (decimal só entre dígitos, ex.: 0,5)` });
            r.formato = vs.some(v => v[1] === 'separador') ? 'compacto' : 'antigo';
            return r;
        }
        let brutos = [];
        if (!vs.length && AP.dividirTokens(t).length === 1 && resolveParaPoste(t, ctx)) {
            // TASK-082: poste sozinho na linha ("11300D", como o Leitor de Pernambuco entrega) = 1 poste, sem quantidade
            r.formato = 'compacto';
            brutos = [{ ativo: t, qtd: 1 }];
        } else if (vs.some(v => v[1] === 'separador')) {
            r.formato = 'compacto';
            const partes = []; let ini = 0;
            for (const [c, tipo] of vs) if (tipo === 'separador') { partes.push(t.slice(ini, c)); ini = c + 1; }
            partes.push(t.slice(ini));
            for (const parte of partes) {
                const palavras = AP.dividirTokens(parte.trim());
                if (!palavras.length) { r.erros.push({ cod: 'FORMATO', msg: `item vazio em '${t}'` }); continue; }
                if (palavras.length > 2) { r.erros.push({ cod: 'FORMATO', msg: `item '${parte.trim()}' com palavras demais (use: quantidade ativo)` }); continue; }
                let qtd, nome;
                if (palavras.length === 2) {
                    if (!RE_QTD.test(palavras[0])) { r.erros.push({ cod: 'FORMATO', msg: `quantidade '${palavras[0]}' inválida em '${parte.trim()}'` }); continue; }
                    qtd = qtdDe(palavras[0]); nome = palavras[1];
                } else { qtd = 1; nome = palavras[0]; }
                if (RE_QTD_ANTIGA.test(nome)) { r.erros.push({ cod: 'MISTURA', msg: `'${nome}' usa a quantidade do formato antigo (qtd-ativo) numa linha em vírgulas; use 'quantidade ativo'` }); continue; }
                brutos.push({ ativo: nome, qtd });
            }
        } else {
            brutos = classico(t).map(i => ({ ativo: i.ativo, qtd: i.qtd }));
        }
        if (r.erros.length) return r;
        let itens = [];
        for (const b of brutos) {
            const [nome, erro] = resolverNome(b.ativo, ctx);
            if (erro) r.erros.push({ cod: 'FORMATO', msg: erro });
            itens.push({ ativo: nome, qtd: b.qtd, operacao: op });
        }
        if (r.erros.length) return r;
        const assinaturas = (ctx && ctx.assinaturas) || {};
        const parametrosLinha = (ctx && ctx.parametros_linha) || {};
        let posteAtual = null;
        for (const it of itens) {
            if (ehPoste(it.ativo)) posteAtual = posteParaCompacto(it.ativo);
            else {
                Object.keys(parametrosLinha).forEach(nome => { it.ativo = comParam(it.ativo, nome, parametrosLinha[nome], assinaturas); });   // TASK-084
                if (posteAtual) it.ativo = comParam(it.ativo, 'POSTE', posteAtual, assinaturas);
            }
        }
        if (op === 'M' || op === '*M') {
            itens = itens.filter(i => !ehPoste(AP.nomeBase(i.ativo))).map(i => Object.assign({}, i, { operacao: op === 'M' ? 'I' : '*I' }));
        }
        r.itens = itens;
        return r;
    }

    const fmt = q => Number.isInteger(q) ? String(q) : String(q);

    function textoEquivalente(texto, classico, ctx) {
        const r = interpretarLinha(texto, '', classico, ctx);
        if (r.erros.length || !r.itens.length) return txt(texto);
        return r.itens.map(i => `${fmt(i.qtd)}-${i.ativo.replace(/[(].*$/, '')}`).join(' ');
    }

    function contextoDaBase(linhas, projeto) {
        const p = txt(projeto).trim().toUpperCase(), nomes = new Set(), postes = new Set();
        for (const l of linhas || []) {
            const lp = txt(l.projeto).trim().toUpperCase();
            if (lp && p && !lp.split('/').map(x => x.trim()).includes(p)) continue;
            const cel = txt(l.ativo).trim();
            if (!cel) continue;
            const a = AP.analisarCelulaBase(cel);
            if (a.erro || !a.base) continue;
            nomes.add(a.base);
            if (ehPoste(a.base)) postes.add(a.base);
        }
        return { apelidos: AP.apelidosDoProjeto(linhas, projeto), nomes, postes, assinaturas: AP.assinaturasDoProjeto(linhas, projeto) };
    }

    function codigoEstrutura(texto, classico, ctx) {
        const r = interpretarLinha(texto, '', classico, ctx);
        for (const i of r.itens) if (!ehPoste(AP.nomeBase(i.ativo))) return AP.nomeBase(i.ativo);
        return '';
    }

    const api = { ehPoste, posteParaCompacto, compactoParaPoste, ehCompacta, interpretarLinha, textoEquivalente, contextoDaBase, codigoEstrutura };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.AnotacaoOutros = api;
})(typeof window !== 'undefined' ? window : globalThis);
