/*
 * ativo_avaliar.js (TASK-123) — avalia o texto de um ativo da Totalizadora SEPARANDO o NOME dos PARÂMETROS. Sem DOM (testado em Node: tests/test_ativo_avaliar_123.py).
 *   • NOME (antes do "("): procurado em ATIVO (e apelidos) → COMPONENTE → CÓDIGO, para o projeto e a origem (a mesma cascata do cálculo; a descrição não resolve nome).
 *     Nome inexistente → naoEncontrado (linha vermelha). A descrição exibida é a da linha do nome.
 *   • PARÂMETROS: avaliados à parte, só se o nome existe: valor fora do escopo cadastrado, parâmetro não declarado, valores a mais, "ativo sem parâmetros
 *     declarados"… → erroParametros (linha âmbar) com a mensagem do motor de parâmetros.
 */
(function (raiz) {
    'use strict';
    const AP = typeof require !== 'undefined' && typeof module !== 'undefined' ? require('./ativo_parametros.js') : raiz.AtivoParametros;
    const MAIUSC = v => String(v == null ? '' : v).trim().toUpperCase();

    /** Índice da base para o projeto: nomes por tipo (ativo/componente/código) com origem e descrição, apelidos e assinaturas dos parâmetros. */
    function construirIndice(linhas, projeto) {
        const p = MAIUSC(projeto);
        const porTipo = { ativo: new Map(), componente: new Map(), codigo: new Map() };
        const juntar = (tipo, nome, info) => { if (!nome) return; const m = porTipo[tipo]; if (!m.has(nome)) m.set(nome, []); m.get(nome).push(info); };
        (linhas || []).forEach(r => {
            const lp = MAIUSC(r.projeto);
            if (lp && p && !lp.split('/').map(x => x.trim()).includes(p)) return;
            const origem = MAIUSC(r.origem), comp = MAIUSC(r.componente), cod = MAIUSC(r.codigo);
            const cel = String(r.ativo || '').trim();
            if (cel) {
                const a = AP.analisarCelulaBase(cel);
                const soDeclaracao = a.tipo === 'declaracao' && !cod && !comp;          // linha só de declaração não entra no cálculo (como no servidor)
                if (!a.erro && a.base && !soDeclaracao) {
                    juntar('ativo', a.base.toUpperCase(), { desc: r.desc_ativo || r.componente || r.desc_codigo || '', origem, grupo: a.tipo === 'grupo', composto: a.tipo === 'composto' });
                }
            }
            if (comp) juntar('componente', comp, { desc: r.desc_ativo || r.componente || '', origem, grupo: false, composto: false });
            if (cod) juntar('codigo', cod, { desc: r.desc_codigo || '', origem, grupo: false, composto: false });
        });
        const apelidos = {};
        Object.entries(AP.apelidosDoProjeto(linhas || [], projeto || '') || {}).forEach(([k, v]) => { apelidos[MAIUSC(k)] = MAIUSC(v); });
        return { porTipo, apelidos, assinaturas: AP.assinaturasDoProjeto(linhas || [], projeto || '') };
    }

    function achar(nome, origem, idx) {
        const o = MAIUSC(origem);
        for (const tipo of ['ativo', 'componente', 'codigo']) {                             // cascata: ativo → componente → código
            const lista = (idx.porTipo[tipo].get(nome) || []).filter(e => !e.origem || !o || e.origem === o);
            if (lista.length) return Object.assign({ tipo }, lista.find(e => !e.grupo) || lista[0]);
        }
        return null;
    }

    function mensagemDe(erro) {
        const m = /^(.*) não tem variantes cadastradas$/.exec(erro);
        return m ? `${m[1]}: ativo sem parâmetros declarados (os parâmetros informados não valem)` : erro;
    }

    /** → { nome, desc, naoEncontrado, erroParametros }  (erroParametros = '' quando não há problema) */
    function avaliarAtivo(texto, origem, idx) {
        const t = String(texto == null ? '' : texto).trim();
        if (!t) return { nome: '', desc: '', naoEncontrado: false, erroParametros: '' };
        const nome = AP.nomeBase(t);
        const base = idx.apelidos[nome] || nome;
        const ent = achar(base, origem, idx);
        if (!ent && !idx.apelidos[nome]) return { nome, desc: '', naoEncontrado: true, erroParametros: '' };
        const res = { nome: base, desc: ent ? ent.desc : '', naoEncontrado: false, erroParametros: '' };
        if (ent && ent.composto) return res;                                                // composto: o cálculo expande e passa os valores aos itens
        const abre = (() => { let col = false; for (let i = 0; i < t.length; i++) { const c = t[i]; if (c === '[') col = true; else if (c === ']') col = false; else if (!col && c === '(') return i; } return -1; })();
        if (abre === -1) return res;                                                         // sem parênteses: nada a conferir (o padrão da declaração vale)
        const literal = /^\[[^\]]*\]/.test(t);
        const oficial = (literal ? '[' + base + ']' : base) + t.slice(abre);                 // apelido → nome oficial, para achar a declaração
        const ped = AP.resolverPedido(oficial, idx.assinaturas);
        if (ped.erros && ped.erros.length) res.erroParametros = ped.erros.map(mensagemDe).join('; ');
        return res;
    }

    const api = { construirIndice, avaliarAtivo };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.AtivoAvaliar = api;
})(typeof window !== 'undefined' ? window : globalThis);
