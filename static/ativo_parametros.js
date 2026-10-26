/*
 * ativo_parametros.js (TASK-079) — ativos com parâmetros por categoria: N1=N1(cabo=XLPE, poste=DT) na base, N1(XLPE,DT) em Outros/Totalizadora.
 * MESMA lógica de services/ativo_parametros.py (paridade testada em tests/test_ativo_parametros.py). Sem DOM.
 * Nomes e valores: letras, números, "_" e "." (sem hífen), sem diferença de caixa. Chaves dos objetos iguais às do Python.
 */
(function (raiz) {
    'use strict';
    const RE_ID = /^[A-Z0-9_.]+$/;
    const norm = t => String(t == null ? '' : t).replace(/\s+/g, '').toUpperCase();

    function nomeBase(texto) {
        let t = String(texto == null ? '' : texto).trim();
        const i = t.indexOf('('), e = t.indexOf('=');
        if (e !== -1 && (i === -1 || e < i)) t = t.slice(0, e);
        else if (i !== -1) t = t.slice(0, i);
        return t.trim().toUpperCase();
    }

    function analisarAtivo(texto) {
        const original = String(texto == null ? '' : texto).trim();
        const r = { texto: original, base: '', tipo: 'simples', params: [], erro: null };
        if (!original) { r.erro = 'ativo vazio'; return r; }
        const i = original.indexOf('('), e = original.indexOf('=');
        if (e !== -1 && (i === -1 || e < i)) {
            const esq = original.slice(0, e).trim().toUpperCase(), dir = original.slice(e + 1).trim();
            const sub = analisarAtivo(dir);
            r.base = esq; r.tipo = 'declaracao'; r.params = sub.params;
            if (sub.erro) r.erro = sub.erro;
            else if (sub.tipo !== 'grupo') r.erro = 'declaração precisa ter parâmetros: N1=N1(cabo=XLPE, poste=DT)';
            else if (sub.base !== esq) r.erro = `declaração com nomes diferentes (${esq} = ${sub.base})`;
            return r;
        }
        if (i === -1) { r.base = original.toUpperCase(); return r; }
        const cont = (s, c) => s.split(c).length - 1;
        if (!original.endsWith(')') || cont(original, '(') !== 1 || cont(original, ')') !== 1) {
            r.base = original.slice(0, i).trim().toUpperCase(); r.tipo = 'grupo'; r.erro = 'parênteses mal formados'; return r;
        }
        r.base = original.slice(0, i).trim().toUpperCase(); r.tipo = 'grupo';
        const corpo = original.slice(i + 1, -1);
        if (!corpo.trim()) { r.erro = 'parênteses vazios'; return r; }
        const vistos = new Set();
        for (let item of corpo.split(',')) {
            item = norm(item);
            if (!item) { r.erro = 'parâmetro vazio'; return r; }
            let par;
            const ig = item.indexOf('=');
            if (ig !== -1) par = { nome: item.slice(0, ig), valor: item.slice(ig + 1) };
            else par = { nome: null, valor: item };
            for (const campo of ['nome', 'valor']) {
                const v = par[campo];
                if (v !== null && !RE_ID.test(v)) { r.erro = `'${v}' inválido (use letras, números, _ e .)`; return r; }
            }
            if (par.nome === '') { r.erro = 'nome do parâmetro vazio'; return r; }
            if (par.valor === '' && par.nome !== null) { r.erro = `parâmetro ${par.nome} sem valor`; return r; }
            if (par.nome !== null) {
                if (vistos.has(par.nome)) { r.erro = `parâmetro ${par.nome} repetido`; return r; }
                vistos.add(par.nome);
            }
            r.params.push(par);
        }
        return r;
    }

    function analisarDeclaracao(texto) {
        const original = String(texto == null ? '' : texto).trim();
        const e = original.indexOf('='), i = original.indexOf('(');
        if (e === -1 || (i !== -1 && e > i)) return analisarAtivo(original);
        const esq = original.slice(0, e).trim().toUpperCase(), dir = original.slice(e + 1).trim();
        const r = { texto: original, base: esq, tipo: 'declaracao', params: [], erro: null };
        const j = dir.indexOf('('), cont = (s, c) => s.split(c).length - 1;
        if (j === -1 || !dir.endsWith(')') || cont(dir, '(') !== 1 || cont(dir, ')') !== 1) { r.erro = 'declaração precisa ter parâmetros: N1=N1(cabo=XLPE, poste=DT)'; return r; }
        if (dir.slice(0, j).trim().toUpperCase() !== esq) { r.erro = `declaração com nomes diferentes (${esq} = ${dir.slice(0, j).trim().toUpperCase()})`; return r; }
        const corpo = dir.slice(j + 1, -1);
        if (!corpo.trim()) { r.erro = 'parênteses vazios'; return r; }
        const vistos = new Set();
        for (let item of corpo.split(',')) {
            item = norm(item);
            if (!item) { r.erro = 'parâmetro vazio'; return r; }
            const ig = item.indexOf('=');
            const nome = ig === -1 ? item : item.slice(0, ig), valor = ig === -1 ? null : item.slice(ig + 1);
            if (!RE_ID.test(nome) || (ig !== -1 && !RE_ID.test(valor))) { r.erro = `'${item}' inválido (use letras, números, _ e .)`; return r; }
            if (vistos.has(nome)) { r.erro = `parâmetro ${nome} repetido`; return r; }
            vistos.add(nome);
            r.params.push({ nome, valor });
        }
        return r;
    }

    function analisarCelulaBase(texto) {
        const t = String(texto == null ? '' : texto), e = t.indexOf('='), i = t.indexOf('(');
        return (e !== -1 && (i === -1 || e < i)) ? analisarDeclaracao(t) : analisarAtivo(t);
    }

    function montarAssinaturas(celulas) {
        const saida = {};
        for (const cel of celulas) {
            const a = analisarCelulaBase(cel);
            if (a.erro || !a.params.length) continue;
            const s = saida[a.base] || (saida[a.base] = { nomes: [], padroes: {}, valores: {}, declarada: false });
            if (a.tipo === 'declaracao') {
                if (s.declarada) continue;
                s.declarada = true;
                const nomes = a.params.map(p => p.nome);
                s.nomes = nomes.concat(s.nomes.filter(n => !nomes.includes(n)));
                for (const p of a.params) {
                    if (p.valor !== null) {
                        s.padroes[p.nome] = p.valor;
                        (s.valores[p.nome] = s.valores[p.nome] || []);
                        if (!s.valores[p.nome].includes(p.valor)) s.valores[p.nome].push(p.valor);
                    }
                }
            } else {
                for (const p of a.params) {
                    if (p.nome === null) continue;
                    if (!s.nomes.includes(p.nome)) s.nomes.push(p.nome);
                    (s.valores[p.nome] = s.valores[p.nome] || []);
                    if (!s.valores[p.nome].includes(p.valor)) s.valores[p.nome].push(p.valor);
                }
            }
        }
        return saida;
    }

    function resolverPedido(texto, assinaturas) {
        const a = analisarAtivo(texto);
        const r = { base: a.base, texto: a.base, params: {}, faltando: [], erros: [], com_assinatura: false };
        if (a.erro) { r.erros.push(a.erro); r.texto = a.texto.toUpperCase(); return r; }
        if (a.tipo === 'declaracao') { r.erros.push('declaração não pode ser usada como pedido'); return r; }
        const sig = (assinaturas || {})[a.base];
        if (!sig) {
            if (a.params.length) { r.erros.push(`${a.base} não tem variantes cadastradas`); r.texto = a.texto.toUpperCase().replace(/ /g, ''); }
            return r;
        }
        r.com_assinatura = true;
        const nomes = sig.nomes, dados = {};
        let pos = 0;
        for (const p of a.params) {
            if (p.nome === null) {
                while (pos < nomes.length && nomes[pos] in dados) pos++;
                if (pos >= nomes.length) { r.erros.push(`parâmetros demais para ${a.base} (esperado: ${nomes.join(', ')})`); continue; }
                dados[nomes[pos]] = p.valor; pos++;
            } else if (!nomes.includes(p.nome)) r.erros.push(`parâmetro ${p.nome} desconhecido em ${a.base} (existem: ${nomes.join(', ')})`);
            else if (p.nome in dados) r.erros.push(`parâmetro ${p.nome} repetido`);
            else dados[p.nome] = p.valor;
        }
        for (const n of nomes) if (!(n in dados) && n in sig.padroes) dados[n] = sig.padroes[n];
        for (const n of nomes) {
            if (n in dados) {
                const ex = sig.valores[n] || [];
                if (ex.length && !ex.includes(dados[n])) r.erros.push(`${n}=${dados[n]} não cadastrado em ${a.base} (existem: ${ex.join(', ')})`);
            } else r.faltando.push(n);
        }
        for (const n of nomes) if (n in dados) r.params[n] = dados[n];
        if (!r.faltando.length) r.texto = `${a.base}(${nomes.map(n => dados[n]).join(',')})`;
        else if (Object.keys(dados).length) r.texto = `${a.base}(${nomes.filter(n => n in dados).map(n => n + '=' + dados[n]).join(',')})`;
        return r;
    }

    function linhaVale(paramsLinha, paramsPedido) {
        for (const nome of Object.keys(paramsLinha || {})) if ((paramsPedido || {})[nome] !== paramsLinha[nome]) return false;
        return true;
    }

    /** Divide por espaços sem quebrar o que está entre parênteses. */
    function dividirTokens(texto) {
        const saida = []; let atual = '', prof = 0;
        for (const ch of String(texto == null ? '' : texto)) {
            if (ch === '(') prof++;
            else if (ch === ')' && prof > 0) prof--;
            if (/\s/.test(ch) && prof === 0) { if (atual) { saida.push(atual); atual = ''; } }
            else atual += ch;
        }
        if (atual) saida.push(atual);
        return saida;
    }

    /** Assinaturas das linhas da base válidas para o projeto (genéricas + as do projeto), na ordem da tabela. */
    function assinaturasDoProjeto(linhas, projeto) {
        const p = String(projeto || '').trim().toUpperCase();
        const validas = (linhas || []).filter(l => {
            const lp = String(l.projeto || '').trim().toUpperCase();
            return !lp || !p || lp.split('/').map(x => x.trim()).includes(p);
        });
        return montarAssinaturas(validas.map(l => l.ativo || ''));
    }

    const api = { nomeBase, analisarAtivo, analisarDeclaracao, analisarCelulaBase, montarAssinaturas, resolverPedido, linhaVale, dividirTokens, assinaturasDoProjeto };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.AtivoParametros = api;
})(typeof window !== 'undefined' ? window : globalThis);
