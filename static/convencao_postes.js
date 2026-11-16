/*
 * convencao_postes.js (TASK-174, ADR-007) — convenção de postes por projeto: tipos, formatos de nome, notação compacta e faixas de altura/esforço. Sem DOM.
 * MESMA lógica de services/convencao_postes.py (paridade testada em tests/test_convencao_postes_174.py). Sem tabela, vale a de Pernambuco (PADRAO).
 */
(function (raiz) {
    'use strict';
    const PADRAO = {
        tipos: [{ codigo: 'DT', sufixo: 'D', nome: 'duplo T' }, { codigo: 'RC', sufixo: 'R', nome: 'redondo' }, { codigo: 'CV', sufixo: 'F', nome: 'concreto' }, { codigo: 'FV', sufixo: 'FV', nome: 'fibra de vidro' }],
        formatos: ['{tipo}{altura}/{esforco}', '{sufixo}{altura}/{esforco}', '{altura}{esforco}{sufixo}'], altura: [7, 16], esforco_digitos: [3, 4], compacto: '{altura}{esforco}{sufixo}',
    };
    const trim = t => String(t == null ? '' : t).trim();
    const clone = o => JSON.parse(JSON.stringify(o));
    const num2 = v => Array.isArray(v) && v.length === 2 && v.every(x => typeof x === 'number');
    const esc = s => s.replace(/[.*+?^${}()|[\]\\\/-]/g, '\\$&');
    const cache = new Map();

    function normalizar(conv) {
        if (!conv || typeof conv !== 'object' || !Array.isArray(conv.tipos) || !conv.tipos.some(t => t && typeof t === 'object' && trim(t.codigo))) return clone(PADRAO);
        const c = { tipos: conv.tipos.filter(t => t && typeof t === 'object' && trim(t.codigo)).map(t => ({ codigo: trim(t.codigo).toUpperCase(), sufixo: (trim(t.sufixo) || trim(t.codigo)).toUpperCase(), nome: trim(t.nome) })) };
        const f = Array.isArray(conv.formatos) ? conv.formatos.map(trim).filter(Boolean) : [];
        c.formatos = f.length ? f : PADRAO.formatos.slice();
        c.altura = num2(conv.altura) ? [Math.trunc(conv.altura[0]), Math.trunc(conv.altura[1])] : PADRAO.altura.slice();
        c.esforco_digitos = num2(conv.esforco_digitos) ? [Math.trunc(conv.esforco_digitos[0]), Math.trunc(conv.esforco_digitos[1])] : PADRAO.esforco_digitos.slice();
        c.compacto = trim(conv.compacto) || PADRAO.compacto;
        return c;
    }

    function validar(conv) {
        const erros = [];
        if (conv === undefined || conv === null) return erros;
        if (typeof conv !== 'object' || Array.isArray(conv)) return ["'postes' precisa ser um objeto."];
        let tipos = conv.tipos;
        if (!Array.isArray(tipos) || !tipos.length) { erros.push('Postes: \'tipos\' precisa ser uma lista com pelo menos um tipo ({"codigo": "DT", "sufixo": "D"}).'); tipos = []; }
        tipos.forEach((t, n) => { if (!t || typeof t !== 'object' || !trim(t.codigo)) erros.push(`Postes, tipo ${n + 1}: falta o 'codigo'.`); });
        const fmts = conv.formatos;
        if (fmts !== undefined && fmts !== null && (!Array.isArray(fmts) || !fmts.every(x => typeof x === 'string'))) erros.push('Postes: \'formatos\' precisa ser uma lista de textos (ex.: "{tipo}{altura}/{esforco}").');
        (Array.isArray(fmts) ? fmts : []).concat(typeof conv.compacto === 'string' ? [conv.compacto] : []).forEach(fm => {
            if (typeof fm !== 'string') return;
            ['tipo', 'sufixo', 'altura', 'esforco'].forEach(ph => { if (fm.split('{' + ph + '}').length - 1 > 1) erros.push(`Postes: o formato '${fm}' repete {${ph}}.`); });
            if (!fm.includes('{altura}') || !fm.includes('{esforco}')) erros.push(`Postes: o formato '${fm}' precisa de {altura} e {esforco}.`);
            if (!fm.includes('{tipo}') && !fm.includes('{sufixo}')) erros.push(`Postes: o formato '${fm}' precisa de {tipo} ou {sufixo}.`);
            if (fm.includes('{esforco}{altura}')) erros.push(`Postes: no formato '${fm}', {altura}{esforco} colados precisam estar nessa ordem.`);
        });
        ['altura', 'esforco_digitos'].forEach(k => { const v = conv[k]; if (v !== undefined && v !== null && !(num2(v) && v[0] <= v[1])) erros.push(`Postes: '${k}' precisa ser [mínimo, máximo].`); });
        return erros;
    }

    function compilar(conv) {
        const chave = JSON.stringify(conv);
        if (cache.has(chave)) return cache.get(chave);
        const ord = (a, b) => (b.length - a.length) || (a < b ? -1 : a > b ? 1 : 0);
        const tipos = Array.from(new Set(conv.tipos.map(t => t.codigo))).sort(ord), sufs = Array.from(new Set(conv.tipos.map(t => t.sufixo))).sort(ord), saida = [];
        for (const fm of conv.formatos) {
            let rx = '', juntos = false;
            for (const parte of fm.split(/(\{tipo\}|\{sufixo\}|\{altura\}\{esforco\}|\{altura\}|\{esforco\})/)) {
                if (parte === '{tipo}') rx += '(?<tipo>' + tipos.map(esc).join('|') + ')';
                else if (parte === '{sufixo}') rx += '(?<sufixo>' + sufs.map(esc).join('|') + ')';
                else if (parte === '{altura}{esforco}') { rx += '(?<digitos>\\d+)'; juntos = true; }
                else if (parte === '{altura}') rx += '(?<altura>\\d+)';
                else if (parte === '{esforco}') rx += '(?<esforco>\\d+)';
                else rx += esc(parte);
            }
            try { saida.push([new RegExp('^' + rx + '$', 'i'), juntos]); } catch (e) { /* formato inválido: ignorado */ }
        }
        cache.set(chave, saida);
        return saida;
    }

    function tipoDe(conv, g) {
        if (g.tipo) return g.tipo.toUpperCase();
        const suf = (g.sufixo || '').toUpperCase(), t = conv.tipos.find(x => x.sufixo === suf);
        return t ? t.codigo : null;
    }
    function sufixoDe(conv, tipo) { const t = conv.tipos.find(x => x.codigo === tipo); return t ? t.sufixo : tipo; }
    function render(modelo, conv, tipo, altura, esforco) { return modelo.split('{tipo}').join(tipo).split('{sufixo}').join(sufixoDe(conv, tipo)).split('{altura}').join(String(altura)).split('{esforco}').join(String(esforco)); }

    function explicito(nome, conv) {
        const t = trim(nome).toUpperCase();
        for (const [rx, juntos] of compilar(conv)) {
            if (juntos) continue;
            const m = rx.exec(t);
            if (m) { const tipo = tipoDe(conv, m.groups); if (tipo) return [tipo, m.groups.altura, m.groups.esforco]; }
        }
        return null;
    }
    function colado(token, conv) {
        const t = trim(token);
        for (const [rx, juntos] of compilar(conv)) {
            if (!juntos) continue;
            const m = rx.exec(t);
            if (m) { const tipo = tipoDe(conv, m.groups); if (tipo) return [tipo, m.groups.digitos]; }
        }
        return null;
    }

    function compactoParaPoste(token, postes, conv) {
        conv = normalizar(conv);
        const t = trim(token), c = colado(t, conv);
        if (!c) return null;
        const [tipo, digitos] = c, [amin, amax] = conv.altura, [dmin, dmax] = conv.esforco_digitos, cands = [];
        for (const hl of [2, 1]) {
            const h = digitos.slice(0, hl), e = digitos.slice(hl);
            if (e.length < dmin || e.length > dmax || h.startsWith('0') || e.startsWith('0') || !h) continue;
            cands.push([parseInt(h, 10), render(conv.formatos[0], conv, tipo, parseInt(h, 10), parseInt(e, 10))]);
        }
        const r = { ativo: null, compacto: t.toUpperCase(), erro: null };
        if (postes && postes.size) {
            const naBase = cands.map(x => x[1]).filter(n => postes.has(n));
            if (naBase.length === 1) { r.ativo = naBase[0]; return r; }
        }
        const plausiveis = cands.filter(x => x[0] >= amin && x[0] <= amax).map(x => x[1]);
        if (plausiveis.length === 1) r.ativo = plausiveis[0];
        else r.erro = !plausiveis.length ? `não foi possível separar altura e esforço de '${t.toUpperCase()}'` : `'${t.toUpperCase()}' é ambíguo (${plausiveis.join(' ou ')})`;
        return r;
    }

    function compactoDoPoste(nome, conv) {
        conv = normalizar(conv);
        const t = trim(nome).toUpperCase(), e = explicito(t, conv);
        if (e) return render(conv.compacto, conv, e[0], e[1], e[2]);
        if (colado(t, conv)) { const c = compactoParaPoste(t, null, conv); if (c && !c.erro) return t; }
        return null;
    }
    function ehPoste(nome, conv) { return compactoDoPoste(nome, conv) !== null; }

    function testar(nome, conv, postes) {
        conv = normalizar(conv);
        const t = trim(nome).toUpperCase(), e = explicito(t, conv);
        if (e) return { poste: true, tipo: e[0], ativo: render(conv.formatos[0], conv, e[0], e[1], e[2]), compacto: render(conv.compacto, conv, e[0], e[1], e[2]), erro: null };
        const c = compactoParaPoste(t, postes, conv);
        if (c === null) return { poste: false, tipo: null, ativo: null, compacto: null, erro: null };
        return { poste: !c.erro, tipo: colado(t, conv)[0], ativo: c.ativo, compacto: c.erro ? null : t, erro: c.erro };
    }

    const api = { PADRAO, normalizar, validar, compactoParaPoste, compactoDoPoste, ehPoste, testar };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.ConvencaoPostes = api;
})(typeof window !== 'undefined' ? window : globalThis);
