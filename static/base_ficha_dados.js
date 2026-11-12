/*
 * base_ficha_dados.js (TASK-143) — parte PURA (sem DOM) da pesquisa e da ficha de ativo da Base de Orçamento: monta o modelo a partir das linhas da base,
 * pesquisa (ativo, apelido, código, descrição), monta a ficha (itens, parâmetros e variantes, chama / é chamado por, verificações) e o uso de um código.
 * Testado em Node: tests/test_base_ficha_143.py.
 */
(function (raiz) {
    'use strict';
    const AP = typeof require !== 'undefined' && typeof module !== 'undefined' ? require('./ativo_parametros.js') : raiz.AtivoParametros;
    const U = v => String(v == null ? '' : v).trim().toUpperCase();
    const T = v => String(v == null ? '' : v).trim();
    const num = v => { const n = parseFloat(String(v == null ? '' : v).replace(/\./g, '').replace(',', '.')); return isNaN(n) ? 0 : n; };
    const fmt = v => (v === '' || v == null) ? '' : String(v);

    /** O projeto vale para a linha? (linha sem projeto vale para todos; "A/B" vale para A e B) */
    function valeNoProjeto(linhaProj, projeto) {
        const lp = U(linhaProj), p = U(projeto);
        return !lp || !p || lp.split('/').map(x => x.trim()).includes(p);
    }

    /** Texto dentro dos parênteses do ATIVO (a condição do grupo / os parâmetros), sem apelidos; '' se não houver. */
    function condicaoDe(celula) {
        const principal = AP.separarApelidos(String(celula || ''))[0];
        const comp = principal.indexOf(':=');
        const cab = comp === -1 ? principal : principal.slice(0, comp);
        const a = cab.indexOf('('), b = cab.lastIndexOf(')');
        return a !== -1 && b > a ? cab.slice(a + 1, b).trim() : '';
    }

    /** Nome (sem parâmetros) de um item de composto: "TERM(cabo)" → "TERM". */
    function nomeDoItem(texto) {
        const t = T(texto), c = t.startsWith('[') ? t.indexOf(']') + 1 : 0, a = t.indexOf('(', c);
        return AP.semColchetes ? AP.semColchetes(a === -1 ? t : t.slice(0, a)) : (a === -1 ? t : t.slice(0, a));
    }

    /**
     * Modelo de pesquisa/ficha. `linhas`: linhas da base ({ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, lv, arred, formula_i, formula_r}).
     * Só entram as linhas válidas para `projeto` (genéricas + as dele). Linha com ATIVO em branco herda o ativo da linha de cima.
     */
    function construir(linhas, projeto) {
        const ativos = new Map(), apelidos = new Map(), codigos = new Map();
        let atual = null;
        (linhas || []).forEach((r, idx) => {
            if (!valeNoProjeto(r.projeto, projeto)) return;
            const cel = T(r.ativo);
            let a = null;
            if (cel) {
                a = AP.analisarCelulaBase(cel);
                if (a.erro && !a.base) { atual = null; return; }
                atual = a;
            } else if (atual) a = { tipo: 'simples', base: atual.base, apelidos: [] };
            else return;
            const nome = a.base;
            if (!nome) return;
            let e = ativos.get(nome);
            if (!e) { e = { nome, apelidos: [], projetos: new Set(), linhas: [], codigos: new Set(), declaracao: null, descricao: '' }; ativos.set(nome, e); }
            (a.apelidos || []).forEach(ap => { if (!e.apelidos.includes(ap)) e.apelidos.push(ap); if (!apelidos.has(U(ap))) apelidos.set(U(ap), nome); });
            if (r.projeto) e.projetos.add(T(r.projeto));
            if (!e.descricao && T(r.desc_ativo)) e.descricao = T(r.desc_ativo);
            const linha = {
                idx, tipo: a.tipo, condicao: cel ? condicaoDe(cel) : '', celula: cel, mdo: T(r.mdo), codigo: T(r.codigo), desc_codigo: T(r.desc_codigo), componente: T(r.componente),
                fator_i: fmt(r.formula_i || r.fator_i), fator_r: fmt(r.formula_r || r.fator_r), lv: T(r.lv) || 'A', arred: T(r.arred) || 'NORMAL', filtro: T(r.filtro),
                formais: a.formais || null, itens: a.tipo === 'composto' ? (a.itens || []) : null
            };
            e.linhas.push(linha);
            if (a.tipo === 'declaracao') e.declaracao = { celula: cel, formais: (a.params || []).map(x => ({ nome: x.nome, padrao: x.valor })) };
            else if (a.tipo === 'composto' && a.formais && !e.declaracao) e.declaracao = { celula: cel, formais: a.formais };
            if (linha.codigo) {
                e.codigos.add(linha.codigo);
                let c = codigos.get(linha.codigo);
                if (!c) { c = { codigo: linha.codigo, descricoes: new Set(), usos: [] }; codigos.set(linha.codigo, c); }
                if (linha.desc_codigo) c.descricoes.add(linha.desc_codigo);
                c.usos.push({ ativo: nome, fator_i: linha.fator_i, fator_r: linha.fator_r, mdo: linha.mdo, condicao: linha.condicao });
            }
        });
        return { ativos, apelidos, codigos, projeto: T(projeto) };
    }

    const resolver = (m, nome) => { const n = T(nome); return m.ativos.has(n) ? n : (m.apelidos.get(U(n)) || (m.ativos.has(U(n)) ? U(n) : null) || Array.from(m.ativos.keys()).find(k => U(k) === U(n)) || null); };

    /** Pesquisa: ativos/apelidos (nome começa > contém), códigos (começa > contém) e descrição. Cada resultado: {tipo:'ativo'|'codigo', ...}. */
    function pesquisar(m, consulta, limite) {
        const q = U(consulta);
        if (!q) return [];
        const max = limite || 30, saida = [];
        const rankNome = n => { const u = U(n); return u === q ? 0 : u.startsWith(q) ? 1 : u.includes(q) ? 2 : -1; };
        const achados = [];
        m.ativos.forEach(e => {
            let r = rankNome(e.nome), via = '';
            e.apelidos.forEach(ap => { const k = rankNome(ap); if (k !== -1 && (r === -1 || k < r)) { r = k; via = ap; } });
            if (r === -1 && U(e.descricao).includes(q)) { r = 3; via = ''; }
            if (r !== -1) achados.push({ tipo: 'ativo', nome: e.nome, apelidos: e.apelidos, via, rank: r, descricao: e.descricao, projetos: Array.from(e.projetos), linhas: e.linhas.length, codigos: e.codigos.size });
        });
        achados.sort((a, b) => a.rank - b.rank || a.nome.localeCompare(b.nome, 'pt-BR', { numeric: true }));
        achados.slice(0, max).forEach(x => saida.push(x));
        const cods = [];
        m.codigos.forEach(c => {
            const desc = Array.from(c.descricoes)[0] || '';
            let r = rankNome(c.codigo);
            if (r === -1 && U(desc).includes(q)) r = 3;
            if (r !== -1) cods.push({ tipo: 'codigo', codigo: c.codigo, descricao: desc, rank: r, usos: new Set(c.usos.map(u => u.ativo)).size, em: Array.from(new Set(c.usos.map(u => u.ativo))).slice(0, 4) });
        });
        cods.sort((a, b) => a.rank - b.rank || a.codigo.localeCompare(b.codigo, 'pt-BR', { numeric: true }));
        cods.slice(0, max).forEach(x => saida.push(x));
        return saida.slice(0, max * 2);
    }

    /** Ficha do ativo `nome` (nome, apelido ou nome sem diferença de caixa). null se não existir. */
    function ficha(m, nome) {
        const base = resolver(m, nome);
        if (!base) return null;
        const e = m.ativos.get(base);
        const params = (e.declaracao && e.declaracao.formais) || [];
        const comuns = e.linhas.filter(l => l.tipo !== 'declaracao' && !l.condicao && !l.itens);
        const variantes = e.linhas.filter(l => l.tipo !== 'declaracao' && (l.condicao || l.itens) && !(l.formais && !l.condicao));
        const itensComposto = e.linhas.filter(l => l.itens).flatMap(l => l.itens.map(i => ({ qtd: i.qtd, texto: i.ativo, nome: nomeDoItem(i.ativo), condicao: l.condicao, existe: !!resolver(m, nomeDoItem(i.ativo)) })));
        const chama = itensComposto;
        const chamadoPor = [];
        m.ativos.forEach(o => {
            if (o.nome === base) return;
            o.linhas.forEach(l => (l.itens || []).forEach(i => { if (resolver(m, nomeDoItem(i.ativo)) === base) chamadoPor.push({ ativo: o.nome, qtd: i.qtd, texto: i.ativo, condicao: l.condicao }); }));
        });
        // parâmetros: padrão da declaração + máscaras que aparecem nas linhas de grupo
        const valores = {};
        e.linhas.forEach(l => { if (l.tipo !== 'declaracao' && l.condicao) l.condicao.split(',').forEach(p => { const k = p.indexOf('='); if (k > 0) { const n = U(p.slice(0, k).replace('!', '')); (valores[n] = valores[n] || new Set()).add(T(p.slice(k + 1))); } }); });
        const parametros = params.map(p => ({ nome: p.nome, padrao: p.padrao, valores: Array.from(valores[p.nome] || []) }));
        const verificacoes = [];
        const ok = (nivel, texto) => verificacoes.push({ nivel, texto });
        chama.filter(c => !c.existe).forEach(c => ok('erro', `chama "${c.nome}", que não existe na base${m.projeto ? ' do projeto' : ''}`));
        const semDesc = e.linhas.filter(l => l.codigo && !l.desc_codigo).length;
        if (semDesc) ok('aviso', `${semDesc} linha(s) com código sem descrição`);
        e.codigos.forEach(c => {
            const d = m.codigos.get(c);
            if (d && d.descricoes.size > 1) ok('aviso', `código ${c} tem ${d.descricoes.size} descrições diferentes na base`);
        });
        if (!e.linhas.some(l => l.codigo || l.itens)) ok('aviso', 'o ativo não tem itens (nenhum código nem chamada)');
        if (params.length && !variantes.length && !comuns.length) ok('aviso', 'declara parâmetros mas não tem linhas');
        if (!verificacoes.length) ok('ok', 'nada a apontar nas verificações automáticas');
        return {
            nome: base, apelidos: e.apelidos, descricao: e.descricao, projetos: Array.from(e.projetos), parametros, declaracao: e.declaracao ? e.declaracao.celula : '',
            comuns, variantes, chama, chamadoPor, verificacoes, codigos: Array.from(e.codigos), nLinhas: comuns.length + variantes.length
        };
    }

    /** Onde um código é usado: [{ativo, fator_i, fator_r, mdo, condicao}] + descrições. */
    function usoDoCodigo(m, codigo) {
        const c = m.codigos.get(T(codigo));
        return c ? { codigo: c.codigo, descricoes: Array.from(c.descricoes), usos: c.usos.slice() } : null;
    }

    const api = { valeNoProjeto, condicaoDe, nomeDoItem, construir, pesquisar, ficha, usoDoCodigo, resolver };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.BaseFichaDados = api;
})(typeof window !== 'undefined' ? window : globalThis);
