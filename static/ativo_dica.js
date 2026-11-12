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
            const params = (sig.nomes || []).map(nm => {
                const valores = ((sig.valores || {})[nm] || []).slice(), mascaras = ((sig.mascaras || {})[nm] || []).slice();
                mascaras.forEach(m => {                                                // TASK-149: máscara "%S336%" também vira o valor selecionável "S336" (um miolo só, sem outro % nem !)
                    const miolo = /^!/.test(m) ? null : String(m).replace(/^%+|%+$/g, '');
                    if (miolo && miolo.indexOf('%') === -1 && miolo.indexOf('|') === -1 && !valores.includes(miolo)) valores.push(miolo);
                });
                return { nome: nm, padrao: (sig.padroes || {})[nm] == null ? null : (sig.padroes || {})[nm], valores, mascaras, aberto: !(sig.fechados || []).includes(nm), porNome: false };
            });
            (sig.livres || []).forEach(nm => params.push({ nome: nm, padrao: null, valores: [], mascaras: [], aberto: true, porNome: true }));
            return { tipo: 'parametros', ativo: base, params };
        }
        if (ctx.nomes && (ctx.nomes.has(n) || ctx.nomes.has(base))) return { tipo: 'sem_parametros', ativo: base, mensagem: `${base}: ativo sem parâmetros declarados` };
        return { tipo: 'nao_encontrado', ativo: n, mensagem: `${n}: ativo não encontrado na base` };
    }

    /** TASK-129: fecha os parênteses que faltam (N1(D, X → N1(D, X)); vírgula/espaço pendentes no fim saem antes; [nome literal] não conta; colchete aberto = texto malformado, não mexe. */
    function completarParenteses(texto) {
        const t = String(texto == null ? '' : texto);
        let abertos = 0, colchete = false;
        for (const c of t) {
            if (colchete) { if (c === ']') colchete = false; continue; }
            if (c === '[') colchete = true;
            else if (c === '(') abertos++;
            else if (c === ')' && abertos > 0) abertos--;
        }
        if (!abertos || colchete) return t;
        return t.replace(/[,\s]+$/, '') + ')'.repeat(abertos);
    }

    /** TASK-149: os parâmetros POSICIONAIS (na ordem) e o índice do parâmetro da vez; `atual` = o parâmetro (por posição ou pelo nome já digitado) ou null. */
    function parametroAtual(c, d) {
        if (!d || d.tipo !== 'parametros') return { pos: [], i: -1, atual: null };
        const pos = d.params.filter(p => !p.porNome);
        let i = c.nomeParam ? pos.findIndex(p => p.nome.toUpperCase() === c.nomeParam) : c.indice;
        if (i < 0 || i >= pos.length) i = -1;
        return { pos, i, atual: i === -1 ? null : pos[i] };
    }

    /** Valores selecionáveis (cadastrados) do parâmetro da vez. */
    function valoresDoAtual(c, d) { const a = parametroAtual(c, d).atual; return a ? a.valores.slice() : []; }

    /**
     * Inserir um valor do parâmetro da vez e AVANÇAR: troca o texto do parâmetro (do início do valor até a próxima "," ou ")") por `valor` e acrescenta ", " — ou ")" se for o
     * último parâmetro. Se a "," / ")" já existe logo depois, só pula por cima dela. Devolve {inicio, fim, inserir, cursor}: substituir [inicio, fim) por `inserir` e pôr o cursor em `cursor`.
     */
    function planoDeInsercao(texto, cursor, c, d, valor) {
        const t = String(texto == null ? '' : texto), { pos, i } = parametroAtual(c, d);
        const ult = i === -1 || i >= pos.length - 1;
        let fim = cursor, prof = 0;
        for (let k = cursor; k < t.length; k++) {                       // fim do parâmetro: próxima "," ou ")" do mesmo nível
            const ch = t[k];
            if (ch === '(') prof++;
            else if (ch === ')') { if (prof === 0) { fim = k; break; } prof--; }
            else if (ch === ',' && prof === 0) { fim = k; break; }
            fim = k + 1;
        }
        const resto = t.slice(fim);
        const base = { inicio: c.inicioParcial, fim, inserir: String(valor) };
        if (!ult) {
            if (resto[0] === ',') { const extra = resto[1] === ' ' ? 2 : 1; return Object.assign(base, { cursor: c.inicioParcial + base.inserir.length + extra }); }
            return Object.assign(base, { inserir: base.inserir + ', ', cursor: c.inicioParcial + base.inserir.length + 2 });
        }
        if (resto[0] === ')') return Object.assign(base, { cursor: c.inicioParcial + base.inserir.length + 1 });
        return Object.assign(base, { inserir: base.inserir + ')', cursor: c.inicioParcial + base.inserir.length + 1 });
    }

    const api = { contextoDoParentese, construirContexto, dica, completarParenteses, parametroAtual, valoresDoAtual, planoDeInsercao };
    if (typeof module !== 'undefined' && module.exports) { module.exports = api; return; }

    /* ───────────────────────── tela ───────────────────────── */
    const doc = raiz.document;
    let cfg = null, caixa = null, ctxCache = null, alvo = null, carregando = false, selecionado = -1, chaveParam = '', valoresVez = [], ctxVez = null, dicaVez = null;

    const CSS = '.dica-param{position:fixed;z-index:9000;min-width:260px;max-width:520px;max-height:280px;overflow:auto;background:#161b22;color:#e6edf3;border:1px solid #30363d;border-radius:8px;padding:8px 10px;font:12px Inter,sans-serif;box-shadow:0 8px 24px rgba(0,0,0,.5)}'
        + '.dica-param .dp-tit{font-weight:700;margin-bottom:4px}.dica-param .dp-aviso{color:#d29922}.dica-param .dp-erro{color:#f85149}'
        + '.dica-param .dp-lin{padding:3px 4px;border-radius:5px;display:flex;flex-wrap:wrap;gap:4px;align-items:center}.dica-param .dp-lin.atual{background:rgba(217,119,6,.16);outline:1px solid rgba(217,119,6,.5)}'
        + '.dica-param .dp-nome{font-weight:600;min-width:64px}.dica-param .dp-pad{color:#8b949e;font-size:11px}'
        + '.dica-param button.dp-val{background:#21262d;color:#e6edf3;border:1px solid #30363d;border-radius:5px;padding:1px 7px;font-size:11px;cursor:pointer}'
        + '.dica-param button.dp-val:hover,.dica-param button.dp-val.sel{border-color:#d97706;color:#d97706}.dica-param button.dp-val.sel{background:rgba(217,119,6,.22)}.dica-param .dp-mask{color:#8b949e;font-size:11px}.dica-param .dp-rod{color:#6e7681;font-size:10px;margin-top:4px}';

    const ehCampo = el => el && el.matches && el.matches('#body-outros .inp-ativo, #body-outros-vazio .inp-ativo-vazio');

    function contexto() {
        const linhas = (cfg.base && cfg.base()) || [];
        const proj = (cfg.projeto && cfg.projeto()) || '';
        if (ctxCache && ctxCache.linhas === linhas && ctxCache.n === linhas.length && ctxCache.proj === proj) return ctxCache.ctx;
        ctxCache = { linhas, n: linhas.length, proj, ctx: construirContexto(linhas, proj) };
        return ctxCache.ctx;
    }

    function fechar() { if (caixa) { caixa.remove(); caixa = null; } alvo = null; }

    /** TASK-149: insere o valor e avança para o próximo parâmetro (", ") ou fecha o ")" no último. */
    function inserir(input, c, valor, d) {
        const plano = planoDeInsercao(input.value, input.selectionStart, c, d, valor);
        input.focus();
        input.setSelectionRange(plano.inicio, plano.fim);
        let ok = false;
        try { ok = doc.execCommand('insertText', false, plano.inserir); } catch (e) { ok = false; }
        if (!ok) {
            input.value = input.value.slice(0, plano.inicio) + plano.inserir + input.value.slice(plano.fim);
            input.dispatchEvent(new Event('input', { bubbles: true }));
        }
        input.setSelectionRange(plano.cursor, plano.cursor);
        selecionado = -1;
        atualizar(input);
    }
    function marcarSelecionado() {
        if (!caixa) return;
        caixa.querySelectorAll('button.dp-val').forEach((b, i) => { b.classList.toggle('sel', i === selecionado); if (i === selecionado && b.scrollIntoView) b.scrollIntoView({ block: 'nearest' }); });
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
                    b.addEventListener('click', () => inserir(input, c, v, d)); lin.appendChild(b);
                });
                const outras = p.mascaras.filter(m => !p.valores.includes(String(m).replace(/^%+|%+$/g, '')));            // as máscaras simples já viraram botões
                if (outras.length) { const m = doc.createElement('span'); m.className = 'dp-mask'; m.title = '% = qualquer texto'; m.textContent = 'também: ' + outras.join('  '); lin.appendChild(m); }
                caixa.appendChild(lin);
            });
            const rod = doc.createElement('div'); rod.className = 'dp-rod'; rod.textContent = '↑↓ escolher · Enter/Tab insere e vai ao próximo parâmetro · ou clique no valor · Esc fecha'; caixa.appendChild(rod);
        }
        ctxVez = c; dicaVez = d; valoresVez = valoresDoAtual(c, d);
        const chave = d.ativo + '|' + parametroAtual(c, d).i + '|' + (c.nomeParam || '');
        if (chave !== chaveParam) { chaveParam = chave; selecionado = -1; }                // trocou de parâmetro: sem valor realçado
        marcarSelecionado();
        const r = input.getBoundingClientRect();
        caixa.style.left = Math.max(4, Math.min(r.left, raiz.innerWidth - 300)) + 'px';
        caixa.style.top = (r.bottom + 3) + 'px';
    }

    function atualizar(input) {
        if (!cfg || !ehCampo(input)) return;
        const c = contextoDoParentese(input.value, input.selectionStart);
        if (!c || !c.nome) { fechar(); return; }
        const linhas = (cfg.base && cfg.base()) || [];
        if (!linhas.length) {                                                      // TASK-128: a base ainda não veio — não desenha nada (antes mostrava "ativo não encontrado" enquanto carregava)
            fechar();
            if (cfg.carregarBase && !carregando) {                                 // carrega uma vez e redesenha quando chegar
                carregando = true;
                Promise.resolve(cfg.carregarBase()).finally(() => { carregando = false; if (doc.activeElement === input) atualizar(input); });
            }
            return;
        }
        alvo = input;
        desenhar(input, c, dica(c.nome, contexto()));
    }

    function iniciar(opcoes) {
        cfg = opcoes || {};
        // TASK-149: teclado dentro dos parênteses (em captura e ANTES dos demais tratadores, que o consumido não pode alcançar)
        doc.addEventListener('keydown', e => {
            if (!caixa || !ehCampo(e.target) || e.isComposing || !dicaVez || dicaVez.tipo !== 'parametros' || !valoresVez.length) return;
            if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
                e.preventDefault(); e.stopImmediatePropagation();
                selecionado = e.key === 'ArrowDown' ? Math.min(selecionado + 1, valoresVez.length - 1) : Math.max(selecionado - 1, 0);
                marcarSelecionado();
            } else if ((e.key === 'Enter' || e.key === 'Tab') && selecionado >= 0 && valoresVez[selecionado] != null && !e.shiftKey) {
                e.preventDefault(); e.stopImmediatePropagation();
                inserir(e.target, ctxVez, valoresVez[selecionado], dicaVez);
            }
        }, true);
        doc.addEventListener('input', e => { if (ehCampo(e.target)) atualizar(e.target); });
        doc.addEventListener('keyup', e => { if (ehCampo(e.target) && e.key !== 'Escape') atualizar(e.target); });
        doc.addEventListener('click', e => { if (ehCampo(e.target)) atualizar(e.target); });
        doc.addEventListener('focusout', e => { if (ehCampo(e.target)) setTimeout(() => { if (doc.activeElement !== alvo) fechar(); }, 120); });
        doc.addEventListener('keydown', e => { if (e.key === 'Escape' && caixa) { fechar(); } }, true);
        // TASK-129: Enter ou sair do campo completa o ")" que falta (em captura: roda antes dos tratadores da própria tela, que leem o valor)
        const completar = input => {
            const novo = completarParenteses(input.value);
            if (novo === input.value) return;
            input.value = novo;
            input.dispatchEvent(new Event('input', { bubbles: true }));          // a tela guarda o valor no 'input'
            fechar();
        };
        doc.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.isComposing && ehCampo(e.target)) completar(e.target); }, true);
        doc.addEventListener('blur', e => { if (ehCampo(e.target)) completar(e.target); }, true);
        raiz.addEventListener('scroll', e => { if (caixa && e.target && caixa.contains && caixa.contains(e.target)) return; fechar(); }, true);
        raiz.addEventListener('resize', fechar);
    }

    raiz.AtivoDica = Object.assign({ iniciar }, api);
})(typeof window !== 'undefined' ? window : globalThis);
