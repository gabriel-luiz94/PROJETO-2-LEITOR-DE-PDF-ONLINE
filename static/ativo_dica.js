/*
 * ativo_dica.js (TASK-122) — ajuda ao digitar "(" na coluna Ativo da tabela Outros: mostra os PARÂMETROS aceitos do ativo
 * (na ordem da declaração, com padrão e valores cadastrados), ou os avisos "ativo sem parâmetros declarados" / "ativo não encontrado".
 * Parte pura (sem DOM; testada em Node: tests/test_ativo_dica_122.py): contextoDoParentese, construirContexto, dica.
 * Parte de tela: iniciar() liga os eventos (delegados) nos campos de Outros e desenha a caixa. Usa a base já carregada na tela (funciona offline).
 */
(function (raiz) {
    'use strict';
    const AP = typeof require !== 'undefined' && typeof module !== 'undefined' ? require('./ativo_parametros.js') : raiz.AtivoParametros;

    /** Texto até o cursor → o parêntese ainda aberto: {nome, indice (posição do parâmetro), nomeParam, parcial, inicioParcial} ou null. */
    function contextoDoParentese(texto, cursor) {
        const t = String(texto == null ? '' : texto).slice(0, cursor == null ? String(texto).length : cursor);
        const pilha = [];
        let colchete = false;
        for (let i = 0; i < t.length; i++) {
            const c = t[i];
            if (c === '[') colchete = true;
            else if (c === ']') colchete = false;
            else if (!colchete && c === '(') pilha.push(i);
            else if (!colchete && c === ')') pilha.pop();
        }
        if (!pilha.length) return null;
        const abre = pilha[pilha.length - 1];
        const m = /(\[[^\]]*\]|[^\s,(]+)$/.exec(t.slice(0, abre));
        if (!m) return { nome: '', indice: 0, nomeParam: null, parcial: '', inicioParcial: t.length };
        let nome = m[1].replace(/^\*?\d+(?:[.,]\d+)?-(?=\S)/, '');                 // "1-IT_2" → "IT_2"; "[CAZ7,94]" fica literal
        if (/^\[[^\]]*\]$/.test(nome)) nome = nome.slice(1, -1);
        const dentro = t.slice(abre + 1);
        const partes = dentro.split(',');
        let seg = partes[partes.length - 1];
        let nomeParam = null, parcial = seg;
        const eq = seg.indexOf('=');
        if (eq !== -1) { nomeParam = seg.slice(0, eq).trim().toUpperCase() || null; parcial = seg.slice(eq + 1); }
        const semEspaco = parcial.replace(/^\s+/, '');
        return { nome: nome.trim().toUpperCase(), indice: partes.length - 1, nomeParam, parcial: semEspaco, inicioParcial: t.length - semEspaco.length };
    }

    /** Contexto da base para a dica: assinaturas, apelidos e nomes que o cálculo resolve para o projeto e a origem OUTROS (ativo/apelido/componente/código). */
    function construirContexto(linhas, projeto) {
        const p = String(projeto || '').trim().toUpperCase();
        const nomes = new Set();
        (linhas || []).forEach(r => {
            const lp = String(r.projeto || '').trim().toUpperCase();
            if (lp && p && !lp.split('/').map(x => x.trim()).includes(p)) return;
            const ro = String(r.origem || '').trim().toUpperCase();
            if (ro && ro !== 'OUTROS') return;
            const cel = String(r.ativo || '').trim();
            if (cel) { const a = AP.analisarCelulaBase(cel); if (!a.erro && a.base) nomes.add(a.base.toUpperCase()); }
            const comp = String(r.componente || '').trim().toUpperCase(), cod = String(r.codigo || '').trim().toUpperCase();
            if (comp) nomes.add(comp);
            if (cod) nomes.add(cod);
        });
        const apelidos = {};
        Object.entries(AP.apelidosDoProjeto(linhas || [], projeto || '') || {}).forEach(([k, v]) => { apelidos[String(k).toUpperCase()] = String(v).toUpperCase(); nomes.add(String(k).toUpperCase()); });
        return { assinaturas: AP.assinaturasDoProjeto(linhas || [], projeto || ''), apelidos, nomes };
    }

    /** O que mostrar para o nome antes do "(": {tipo: 'parametros'|'sem_parametros'|'nao_encontrado', ...}. */
    function dica(nome, ctx) {
        const n = String(nome || '').trim().toUpperCase();
        const base = (ctx.apelidos && ctx.apelidos[n]) || n;
        const sig = (ctx.assinaturas || {})[base];
        if (sig && ((sig.nomes && sig.nomes.length) || (sig.livres && sig.livres.length))) {
            const params = (sig.nomes || []).map(nm => ({
                nome: nm, padrao: (sig.padroes || {})[nm] == null ? null : (sig.padroes || {})[nm],
                valores: ((sig.valores || {})[nm] || []).slice(), mascaras: ((sig.mascaras || {})[nm] || []).slice(),
                aberto: !(sig.fechados || []).includes(nm), porNome: false
            }));
            (sig.livres || []).forEach(nm => params.push({ nome: nm, padrao: null, valores: [], mascaras: [], aberto: true, porNome: true }));
            return { tipo: 'parametros', ativo: base, params };
        }
        if (ctx.nomes && (ctx.nomes.has(n) || ctx.nomes.has(base))) return { tipo: 'sem_parametros', ativo: base, mensagem: `${base}: ativo sem parâmetros declarados` };
        return { tipo: 'nao_encontrado', ativo: n, mensagem: `${n}: ativo não encontrado na base` };
    }

    const api = { contextoDoParentese, construirContexto, dica };
    if (typeof module !== 'undefined' && module.exports) { module.exports = api; return; }

    /* ───────────────────────── tela ───────────────────────── */
    const doc = raiz.document;
    let cfg = null, caixa = null, ctxCache = null, alvo = null, carregando = false;

    const CSS = '.dica-param{position:fixed;z-index:9000;min-width:260px;max-width:520px;max-height:280px;overflow:auto;background:#161b22;color:#e6edf3;border:1px solid #30363d;border-radius:8px;padding:8px 10px;font:12px Inter,sans-serif;box-shadow:0 8px 24px rgba(0,0,0,.5)}'
        + '.dica-param .dp-tit{font-weight:700;margin-bottom:4px}.dica-param .dp-aviso{color:#d29922}.dica-param .dp-erro{color:#f85149}'
        + '.dica-param .dp-lin{padding:3px 4px;border-radius:5px;display:flex;flex-wrap:wrap;gap:4px;align-items:center}.dica-param .dp-lin.atual{background:rgba(217,119,6,.16);outline:1px solid rgba(217,119,6,.5)}'
        + '.dica-param .dp-nome{font-weight:600;min-width:64px}.dica-param .dp-pad{color:#8b949e;font-size:11px}'
        + '.dica-param button.dp-val{background:#21262d;color:#e6edf3;border:1px solid #30363d;border-radius:5px;padding:1px 7px;font-size:11px;cursor:pointer}'
        + '.dica-param button.dp-val:hover{border-color:#d97706;color:#d97706}.dica-param .dp-mask{color:#8b949e;font-size:11px}.dica-param .dp-rod{color:#6e7681;font-size:10px;margin-top:4px}';

    const ehCampo = el => el && el.matches && el.matches('#body-outros .inp-ativo, #body-outros-vazio .inp-ativo-vazio');

    function contexto() {
        const linhas = (cfg.base && cfg.base()) || [];
        const proj = (cfg.projeto && cfg.projeto()) || '';
        if (ctxCache && ctxCache.linhas === linhas && ctxCache.n === linhas.length && ctxCache.proj === proj) return ctxCache.ctx;
        ctxCache = { linhas, n: linhas.length, proj, ctx: construirContexto(linhas, proj) };
        return ctxCache.ctx;
    }

    function fechar() { if (caixa) { caixa.remove(); caixa = null; } alvo = null; }

    function inserir(input, c, valor) {
        input.focus();
        input.setSelectionRange(c.inicioParcial, input.selectionStart);
        let ok = false;
        try { ok = doc.execCommand('insertText', false, valor); } catch (e) { ok = false; }
        if (!ok) {
            const fim = input.selectionStart;
            input.value = input.value.slice(0, c.inicioParcial) + valor + input.value.slice(fim);
            input.setSelectionRange(c.inicioParcial + valor.length, c.inicioParcial + valor.length);
            input.dispatchEvent(new Event('input', { bubbles: true }));
        }
        atualizar(input);
    }

    function desenhar(input, c, d) {
        if (!caixa) {
            if (!doc.getElementById('dica-param-css')) { const st = doc.createElement('style'); st.id = 'dica-param-css'; st.textContent = CSS; doc.head.appendChild(st); }
            caixa = doc.createElement('div'); caixa.className = 'dica-param';
            caixa.addEventListener('mousedown', e => e.preventDefault());          // clicar na caixa não tira o foco do campo
            doc.body.appendChild(caixa);
        }
        caixa.innerHTML = '';
        const tit = doc.createElement('div'); tit.className = 'dp-tit';
        if (d.tipo !== 'parametros') {
            tit.className += d.tipo === 'nao_encontrado' ? ' dp-erro' : ' dp-aviso';
            tit.textContent = d.mensagem; caixa.appendChild(tit);
        } else {
            tit.textContent = `${d.ativo} — parâmetros`; caixa.appendChild(tit);
            d.params.forEach((p, i) => {
                const lin = doc.createElement('div');
                const atual = c.nomeParam ? p.nome.toUpperCase() === c.nomeParam : (!p.porNome && i === c.indice);
                lin.className = 'dp-lin' + (atual ? ' atual' : '');
                const nm = doc.createElement('span'); nm.className = 'dp-nome'; nm.textContent = `${p.porNome ? '·' : i + 1 + '.'} ${p.nome.toLowerCase()}`; lin.appendChild(nm);
                const pad = doc.createElement('span'); pad.className = 'dp-pad';
                pad.textContent = p.porNome ? '(só na fórmula: informe por nome, ex.: ' + p.nome.toLowerCase() + '=…)' : (p.padrao != null ? `padrão ${p.padrao}` : 'sem padrão') + (p.aberto ? ' · aceita qualquer valor' : '');
                lin.appendChild(pad);
                p.valores.forEach(v => {
                    if (!atual) { const t = doc.createElement('span'); t.className = 'dp-mask'; t.textContent = v; lin.appendChild(t); return; }      // só o parâmetro da posição atual insere valor
                    const b = doc.createElement('button'); b.type = 'button'; b.className = 'dp-val'; b.textContent = v;
                    b.addEventListener('click', () => inserir(input, c, v)); lin.appendChild(b);
                });
                if (p.mascaras.length) { const m = doc.createElement('span'); m.className = 'dp-mask'; m.title = '% = qualquer texto'; m.textContent = 'também: ' + p.mascaras.join('  '); lin.appendChild(m); }
                caixa.appendChild(lin);
            });
            const rod = doc.createElement('div'); rod.className = 'dp-rod'; rod.textContent = 'Clique num valor do parâmetro destacado para inserir · valores na ordem, separados por vírgula · Esc fecha'; caixa.appendChild(rod);
        }
        const r = input.getBoundingClientRect();
        caixa.style.left = Math.max(4, Math.min(r.left, raiz.innerWidth - 300)) + 'px';
        caixa.style.top = (r.bottom + 3) + 'px';
    }

    function atualizar(input) {
        if (!cfg || !ehCampo(input)) return;
        const c = contextoDoParentese(input.value, input.selectionStart);
        if (!c || !c.nome) { fechar(); return; }
        const linhas = (cfg.base && cfg.base()) || [];
        if (!linhas.length && cfg.carregarBase && !carregando) {                  // a base ainda não veio: carrega uma vez e redesenha
            carregando = true;
            Promise.resolve(cfg.carregarBase()).finally(() => { carregando = false; if (doc.activeElement === input) atualizar(input); });
            return;
        }
        alvo = input;
        desenhar(input, c, dica(c.nome, contexto()));
    }

    function iniciar(opcoes) {
        cfg = opcoes || {};
        doc.addEventListener('input', e => { if (ehCampo(e.target)) atualizar(e.target); });
        doc.addEventListener('keyup', e => { if (ehCampo(e.target) && e.key !== 'Escape') atualizar(e.target); });
        doc.addEventListener('click', e => { if (ehCampo(e.target)) atualizar(e.target); });
        doc.addEventListener('focusout', e => { if (ehCampo(e.target)) setTimeout(() => { if (doc.activeElement !== alvo) fechar(); }, 120); });
        doc.addEventListener('keydown', e => { if (e.key === 'Escape' && caixa) { fechar(); } }, true);
        raiz.addEventListener('scroll', e => { if (caixa && e.target && caixa.contains && caixa.contains(e.target)) return; fechar(); }, true);
        raiz.addEventListener('resize', fechar);
    }

    raiz.AtivoDica = Object.assign({ iniciar }, api);
})(typeof window !== 'undefined' ? window : globalThis);
