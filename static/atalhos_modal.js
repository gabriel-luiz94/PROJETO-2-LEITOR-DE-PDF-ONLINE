/*
 * atalhos_modal.js (TASK-068) — atalhos e favoritos dos modais geradores (Cabos e Postes/Estruturas).
 * Três origens, mescladas sem duplicar: FÁBRICA (embutida aqui) + PROJETO (nuvem, GET/PUT /api/atalhos/{modal})
 * + COMPUTADOR (localStorage). O que é "oculto" fica só no computador. Parte pura (mesclar, rótulos, ids) testada em
 * Node: tests/test_atalhos_modal.py. O localStorage e o fetch são sempre protegidos (podem falhar/estar vazios).
 */
(function (raiz) {
    'use strict';

    const OPERACOES = ['I', '*I', 'R', '*R', 'M', '*M'];

    const GRUPOS = {
        cabos: [
            { id: 'ativos', titulo: 'ATIVOS', cor: '#58a6ff', dica: 'Altera o ativo da linha em destaque', comComp: true },
            { id: 'fases', titulo: 'FASES', cor: '#d2991e', dica: 'Altera a fase da linha em destaque' },
            { id: 'comps', titulo: 'COMPRIMENTOS (m)', cor: '#3fb950', dica: 'Altera o comprimento da linha em destaque', numerico: true }
        ],
        postes: [
            { id: 'postes', titulo: 'POSTES', cor: '#58a6ff', dica: 'Cada clique = nova linha' },
            { id: 'mt', titulo: 'ESTRUTURA MT', cor: '#d2991e', dica: 'Acumula quantidade na linha em destaque' },
            { id: 'bt', titulo: 'ESTRUTURAS BT', cor: '#3fb950', dica: 'Acumula quantidade na linha em destaque' }
        ]
    };

    const FABRICA = {
        cabos: {
            ativos: [['CAA 2', 80], ['CAA 1/0', 80], ['CAA 4/0', 80], ['CAA 336,4', 80], ['M3x1x35+35', 40], ['M3x1x70+70', 40],
                     ['M3x1x120+70', 40], ['P 50', 80], ['P 120', 80], ['P 185', 80], ['CAZ 9,5', 40]],
            fases: [['ABC'], ['A'], ['B'], ['C'], ['AC']],
            comps: [['20'], ['30'], ['40'], ['50'], ['60'], ['80'], ['100']]
        },
        postes: {
            postes: [['DT11/300'], ['DT11/600'], ['DT11/1000'], ['DT12/600'], ['DT10/300'], ['DT10/150']],
            mt: ['N1', 'N2', 'N3', 'N4', 'B1', 'B2', 'B3', 'B4', 'CE1', 'CE2', 'CE3', 'CE4', 'U1', 'U2', 'U3', 'U4'].map(x => [x]),
            bt: ['SI1', 'SI3', 'SI4', 'RA2', 'S2', 'S4'].map(x => [x])
        }
    };

    /** Comprimento padrão de um cabo quando ninguém escolheu: 40 m multiplexado/MTX, 80 m os demais (regra do gerador). */
    function compPadrao(ativo) {
        return /MULT|MTX/i.test(String(ativo || '')) ? 40 : 80;
    }

    function novoId() {
        return 'a' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    }

    function itensDeFabrica(modal) {
        const out = {};
        (GRUPOS[modal] || []).forEach(g => {
            out[g.id] = ((FABRICA[modal] || {})[g.id] || []).map(([valor, comp]) => ({
                id: 'f:' + g.id + ':' + valor, grupo: g.id, valor: String(valor), comp: comp === undefined ? undefined : comp, origem: 'fabrica'
            }));
        });
        return out;
    }

    /**
     * Junta fábrica + projeto + computador. Mesmo valor no mesmo grupo aparece uma vez (vence a primeira origem).
     * `ocultos` = ids escondidos neste computador. Devolve { grupos:{id:[item]}, favoritos:[fav], ocultosCount }.
     * `incluirOcultos` mantém os escondidos (marcados `oculto:true`) para o modo de edição.
     */
    function mesclar(modal, fontes, incluirOcultos) {
        const f = fontes || {};
        const ocultos = new Set(f.ocultos || []);
        const base = itensDeFabrica(modal);
        const grupos = {};
        let escondidos = 0;
        (GRUPOS[modal] || []).forEach(g => {
            const vistos = new Set(), lista = [];
            const empilha = it => {
                if (!it || it.grupo !== g.id || typeof it.valor !== 'string') return;
                const chave = it.valor.trim().toUpperCase();
                if (vistos.has(chave)) return;
                vistos.add(chave);
                const oculto = ocultos.has(it.id);
                if (oculto) escondidos++;
                if (oculto && !incluirOcultos) return;
                lista.push(Object.assign({}, it, oculto ? { oculto: true } : {}));
            };
            (base[g.id] || []).forEach(empilha);
            ((f.projeto && f.projeto.itens) || []).forEach(it => empilha(Object.assign({}, it, { origem: 'projeto' })));
            ((f.local && f.local.itens) || []).forEach(it => empilha(Object.assign({}, it, { origem: 'local' })));
            grupos[g.id] = lista;
        });
        const favoritos = []
            .concat((((f.projeto && f.projeto.favoritos) || []).map(x => Object.assign({}, x, { origem: 'projeto' }))))
            .concat((((f.local && f.local.favoritos) || []).map(x => Object.assign({}, x, { origem: 'local' }))))
            .filter(x => x && x.linha && !ocultos.has(x.id) || incluirOcultos);
        return { grupos, favoritos, ocultosCount: escondidos };
    }

    /** "R · CAA 2 · ABC · 80 m" / "I · DT11/300 1-N3". */
    function rotuloFavorito(modal, fav) {
        const l = (fav && fav.linha) || {};
        if (fav && fav.nome) return fav.nome;
        const partes = [l.op || 'I', l.ativo || '(vazio)'];
        if (modal === 'cabos') { if (l.fase) partes.push(l.fase); if (l.comp) partes.push(l.comp + ' m'); }
        return partes.join(' · ');
    }

    /* ── armazenamento do computador ── */
    const chaveLocal = modal => 'atalhos_' + modal + '_v1';
    const vazio = () => ({ itens: [], favoritos: [], ocultos: [] });

    function lerLocal(modal) {
        try {
            const d = JSON.parse(raiz.localStorage.getItem(chaveLocal(modal)) || 'null');
            if (d && typeof d === 'object') return { itens: Array.isArray(d.itens) ? d.itens : [], favoritos: Array.isArray(d.favoritos) ? d.favoritos : [], ocultos: Array.isArray(d.ocultos) ? d.ocultos : [] };
        } catch (e) { /* sem localStorage ou JSON ruim: começa vazio */ }
        return vazio();
    }
    function gravarLocal(modal, doc) {
        try { raiz.localStorage.setItem(chaveLocal(modal), JSON.stringify(doc)); return true; } catch (e) { return false; }
    }

    /* ── projeto (nuvem) ── */
    const chaveCache = (modal, projeto) => 'atalhos_cache_' + modal + '_' + projeto;
    function lerCacheProjeto(modal, projeto) {
        try { const d = JSON.parse(raiz.localStorage.getItem(chaveCache(modal, projeto)) || 'null'); if (d) return d; } catch (e) { /* ignora */ }
        return { itens: [], favoritos: [] };
    }
    async function carregarProjeto(modal, projeto) {
        if (!projeto) return { itens: [], favoritos: [] };
        try {
            const r = await fetch('/api/atalhos/' + encodeURIComponent(modal) + '?projeto_codigo=' + encodeURIComponent(projeto));
            if (!r.ok) throw new Error('HTTP ' + r.status);
            const d = (await r.json()).dados || { itens: [], favoritos: [] };
            try { raiz.localStorage.setItem(chaveCache(modal, projeto), JSON.stringify(d)); } catch (e) { /* cache é opcional */ }
            return d;
        } catch (e) {
            return lerCacheProjeto(modal, projeto);       // nuvem fora do ar: última leitura
        }
    }
    /** Devolve { ok, dados?, erro?, rede? }. `rede:true` = não chegou ao servidor (o chamador pode guardar no computador). */
    async function gravarProjeto(modal, projeto, doc) {
        try {
            const r = await fetch('/api/atalhos/' + encodeURIComponent(modal), {
                method: 'PUT', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ projeto_codigo: projeto, dados: { itens: doc.itens || [], favoritos: doc.favoritos || [] } })
            });
            const j = await r.json().catch(() => ({}));
            if (r.ok) {
                try { raiz.localStorage.setItem(chaveCache(modal, projeto), JSON.stringify(j.dados)); } catch (e) { /* ok */ }
                return { ok: true, dados: j.dados };
            }
            const det = j && j.detail;
            return { ok: false, erro: typeof det === 'string' ? det : (det && det.erros ? det.erros.join(' ') : 'Erro ' + r.status) };
        } catch (e) {
            return { ok: false, rede: true, erro: 'Sem conexão com o servidor.' };
        }
    }

    const api = { OPERACOES, GRUPOS, FABRICA, compPadrao, novoId, mesclar, rotuloFavorito, lerLocal, gravarLocal, lerCacheProjeto, carregarProjeto, gravarProjeto };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    raiz.AtalhosModal = api;
})(typeof window !== 'undefined' ? window : globalThis);
