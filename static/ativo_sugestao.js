/*
 * ativo_sugestao.js (TASK-136) — sugestão de ATIVOS ao digitar, como nas IDEs: "SU" → SUH, SUPORTE_X…
 *   • só nomes de ATIVO e APELIDOS que existem na Base de Orçamento para o PROJETO atual (linhas do projeto + genéricas) e para a origem (OUTROS);
 *     componentes e códigos não entram; ativo de outro projeto nunca aparece.
 *   • a sugestão vale para o item que está sendo digitado (depois de "qtd-", de ", " ou de um espaço); dentro de parênteses quem trabalha é a dica de parâmetros.
 *   • ↑/↓ navegam, Tab aceita (o destacado ou o 1º), Enter aceita só o destacado (sem destaque o Enter segue como sempre), Esc fecha, clique aceita.
 *   • botão "💡 Sugestões" (chip da barra do Resumo) liga/desliga; guardado em localStorage (`sugestao_ativos`, padrão ligado).
 * Parte pura (sem DOM; testada em Node: tests/test_ativo_sugestao_136.py): indiceDeNomes, tokenAtual, sugerir, aplicar.
 */
(function (raiz) {
    'use strict';
    const AP = typeof require !== 'undefined' && typeof module !== 'undefined' ? require('./ativo_parametros.js') : raiz.AtivoParametros;
    const MAIUSC = v => String(v == null ? '' : v).trim().toUpperCase();

    /** Nomes (ativos e apelidos) do projeto/origem: [{ nome, texto (o que é inserido), tipo: 'ativo'|'apelido', de }] sem repetição, em ordem alfabética. */
    function indiceDeNomes(linhas, projeto, origem) {
        const p = MAIUSC(projeto), o = MAIUSC(origem);
        const mapa = new Map();
        const juntar = (nome, texto, tipo, de) => { if (nome && !mapa.has(nome)) mapa.set(nome, { nome, texto, tipo, de }); };
        (linhas || []).forEach(r => {
            const lp = MAIUSC(r.projeto);
            if (lp && p && !lp.split('/').map(x => x.trim()).includes(p)) return;                       // linha de OUTRO projeto: não conta
            if (lp && !p) return;                                                                         // sem projeto escolhido só valem as genéricas
            const ro = MAIUSC(r.origem);
            if (ro && o && ro !== o) return;
            const cel = String(r.ativo || '').trim();
            if (!cel) return;
            const a = AP.analisarCelulaBase(cel);
            if (a.erro || !a.base) return;
            const literal = /^\[/.test(AP.separarApelidos(cel)[0]);
            const texto = nome => (literal && nome === a.base ? '[' + nome + ']' : nome);
            juntar(a.base, texto(a.base), 'ativo', null);
            (a.apelidos || []).forEach(ap => juntar(ap, ap, 'apelido', a.base));
        });
        return Array.from(mapa.values()).sort((x, y) => x.nome.localeCompare(y.nome, 'pt-BR', { numeric: true }));
    }

    /** O item que está sendo digitado até o cursor: { inicio, fim, query } ou null (dentro de parênteses/colchetes, sem texto ou só quantidade). */
    function tokenAtual(texto, cursor) {
        const t = String(texto == null ? '' : texto);
        const ate = t.slice(0, cursor == null ? t.length : cursor);
        let prof = 0, colchete = false, ini = 0;
        for (let i = 0; i < ate.length; i++) {
            const c = ate[i];
            if (colchete) { if (c === ']') colchete = false; continue; }
            if (c === '[') colchete = true;
            else if (c === '(') prof++;
            else if (c === ')' && prof > 0) prof--;
            else if (prof === 0 && (c === ',' || /\s/.test(c))) ini = i + 1;
        }
        if (prof > 0 || colchete) return null;
        let seg = ate.slice(ini);
        const q = /^\*?\d+(?:[.,]\d+)?[-xX]/.exec(seg);                                                  // "2-N1", "*1-N1": a quantidade não é o nome
        const deslocar = q ? q[0].length : 0;
        seg = seg.slice(deslocar);
        if (!seg || /^\*?(?:\d{1,2}|\d+[.,]\d+)$/.test(seg)) return null;                                     // vazio ou número curto (quantidade): sem sugestão; 3+ dígitos pode ser o início de um poste (11300D)
        let fim = ate.length;                                                                              // o item vai até o fim da palavra que o cursor está tocando
        while (fim < t.length && !/[\s,()\[\]]/.test(t[fim])) fim++;
        return { inicio: ini + deslocar, fim, query: seg };
    }

    /** Até `limite` nomes: começam com a consulta primeiro, depois os que a contêm. Nome idêntico à consulta, sozinho, não vira sugestão. */
    function sugerir(query, indice, limite) {
        const q = MAIUSC(query).replace(/^\[/, '');
        if (!q) return [];
        const soNumero = /^\d+$/.test(q);                                                                // só dígitos: só quem COMEÇA com eles (evita sugerir por um pedaço no meio do nome)
        const come = [], contem = [];
        for (const e of indice || []) {
            const n = e.nome.toUpperCase();
            if (n.startsWith(q)) come.push(e); else if (!soNumero && n.includes(q)) contem.push(e);
        }
        const todos = come.concat(contem);
        if (todos.length === 1 && todos[0].nome.toUpperCase() === q) return [];
        return todos.slice(0, limite || 12);
    }

    /** Troca o item sob o cursor pelo escolhido: { texto, cursor }. */
    function aplicar(texto, token, escolha) {
        const t = String(texto == null ? '' : texto);
        const novo = t.slice(0, token.inicio) + escolha.texto + t.slice(token.fim);
        return { texto: novo, cursor: token.inicio + escolha.texto.length };
    }

    const api = { indiceDeNomes, tokenAtual, sugerir, aplicar };
    if (typeof module !== 'undefined' && module.exports) { module.exports = api; return; }

    /* ───────────────────────── tela ───────────────────────── */
    const doc = raiz.document;
    const CHAVE = 'sugestao_ativos';
    let cfg = null, caixa = null, itens = [], destaque = -1, alvo = null, token = null, cache = null, carregando = false;

    function ligado() { try { return raiz.localStorage.getItem(CHAVE) !== '0'; } catch (e) { return true; } }
    function definir(v) { try { raiz.localStorage.setItem(CHAVE, v ? '1' : '0'); } catch (e) { /* sem localStorage: vale só até recarregar */ } if (!v) fechar(); atualizarBotao(); }
    function atualizarBotao() {
        const b = doc.getElementById('chip-sugestoes');
        if (!b) return;
        const on = ligado();
        b.classList.toggle('on', on); b.setAttribute('aria-pressed', String(on));
        b.title = on ? 'Sugestões de ativos ao digitar: LIGADAS (clique para desligar)' : 'Sugestões de ativos ao digitar: DESLIGADAS (clique para ligar)';
        b.textContent = on ? '💡 Sugestões' : '💡 Sugestões (off)';
    }

    const ehCampo = el => el && el.matches && el.matches('#body-outros .inp-ativo, #body-outros-vazio .inp-ativo-vazio, #body-totalizadora .tot-ativo-input');
    const origemDe = el => (el.matches('#body-totalizadora .tot-ativo-input') ? '' : 'OUTROS');

    function indice(origem) {
        const linhas = (cfg.base && cfg.base()) || [], proj = (cfg.projeto && cfg.projeto()) || '';
        if (cache && cache.linhas === linhas && cache.n === linhas.length && cache.proj === proj && cache.origem === origem) return cache.lista;
        cache = { linhas, n: linhas.length, proj, origem, lista: indiceDeNomes(linhas, proj, origem) };
        return cache.lista;
    }

    function fechar() { if (caixa) { caixa.remove(); caixa = null; } itens = []; destaque = -1; alvo = null; token = null; }

    function marcar(i) {
        destaque = i;
        if (!caixa) return;
        Array.from(caixa.children).forEach((el, k) => { el.classList.toggle('ativa', k === i); if (k === i && el.scrollIntoView) el.scrollIntoView({ block: 'nearest' }); });
    }

    function aceitar(i) {
        if (!alvo || !token || !itens[i]) return false;
        const input = alvo, r = aplicar(input.value, token, itens[i]);
        fechar();
        input.value = r.texto;
        try { input.setSelectionRange(r.cursor, r.cursor); } catch (e) { /* campo sem seleção */ }
        input.dispatchEvent(new Event('input', { bubbles: true }));                                    // a tela guarda o valor no 'input'
        fechar();
        input.focus();
        return true;
    }

    function desenhar(input) {
        const r = input.getBoundingClientRect();
        if (!caixa) {
            if (!doc.getElementById('sugestao-ativos-css')) {
                const st = doc.createElement('style'); st.id = 'sugestao-ativos-css';
                st.textContent = '.sugestao-ativos{position:fixed;z-index:9100;min-width:200px;max-width:420px;max-height:260px;overflow:auto;background:#161b22;color:#e6edf3;border:1px solid #30363d;border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.5);font:12px/1.4 system-ui,sans-serif;padding:3px}'
                    + '.sugestao-ativos .sg{padding:4px 8px;border-radius:5px;cursor:pointer;display:flex;gap:8px;align-items:baseline;white-space:nowrap}.sugestao-ativos .sg.ativa,.sugestao-ativos .sg:hover{background:rgba(217,119,6,.22)}'
                    + '.sugestao-ativos .sg b{font-weight:700;color:#f0f6fc}.sugestao-ativos .sg i{font-style:normal;color:#8b949e;font-size:11px}.sugestao-ativos .sg-rod{padding:3px 8px 1px;color:#6e7681;font-size:10px}';
                doc.head.appendChild(st);
            }
            caixa = doc.createElement('div'); caixa.className = 'sugestao-ativos'; doc.body.appendChild(caixa);
        }
        caixa.innerHTML = '';
        const q = token.query.toUpperCase();
        itens.forEach((e, i) => {
            const linha = doc.createElement('div'); linha.className = 'sg';
            const nome = doc.createElement('b');
            const k = e.nome.toUpperCase().indexOf(q);
            if (k >= 0) { nome.append(e.nome.slice(0, k)); const m = doc.createElement('span'); m.style.color = '#d97706'; m.textContent = e.nome.slice(k, k + q.length); nome.append(m, e.nome.slice(k + q.length)); }
            else nome.textContent = e.nome;
            linha.appendChild(nome);
            if (e.tipo === 'apelido') { const d = doc.createElement('i'); d.textContent = 'apelido de ' + e.de; linha.appendChild(d); }
            linha.addEventListener('mousedown', ev => { ev.preventDefault(); aceitar(i); });
            linha.addEventListener('mouseover', () => marcar(i));
            caixa.appendChild(linha);
        });
        const rod = doc.createElement('div'); rod.className = 'sg-rod'; rod.textContent = '↑↓ escolhe · Tab aceita · Esc fecha'; caixa.appendChild(rod);
        caixa.style.left = Math.max(4, Math.min(r.left, raiz.innerWidth - 300)) + 'px';
        caixa.style.top = (r.bottom + 3) + 'px';
        marcar(-1);
    }

    function atualizar(input) {
        if (!cfg || !ligado() || !ehCampo(input)) { fechar(); return; }
        const t = tokenAtual(input.value, input.selectionStart);
        if (!t) { fechar(); return; }
        const linhas = (cfg.base && cfg.base()) || [];
        if (!linhas.length) {                                                                              // a base ainda não veio: pede uma vez e redesenha
            fechar();
            if (cfg.carregarBase && !carregando) { carregando = true; Promise.resolve(cfg.carregarBase()).finally(() => { carregando = false; if (doc.activeElement === input) atualizar(input); }); }
            return;
        }
        const lista = sugerir(t.query, indice(origemDe(input)), 12);
        if (!lista.length) { fechar(); return; }
        alvo = input; token = t; itens = lista;
        desenhar(input);
    }

    function iniciar(opcoes) {
        cfg = opcoes || {};
        doc.addEventListener('input', e => { if (ehCampo(e.target)) atualizar(e.target); });
        doc.addEventListener('click', e => { if (ehCampo(e.target)) atualizar(e.target); });
        doc.addEventListener('focusout', e => { if (ehCampo(e.target)) setTimeout(() => { if (doc.activeElement !== alvo) fechar(); }, 120); });
        doc.addEventListener('keydown', e => {
            if (!caixa || !alvo || e.target !== alvo || e.isComposing) return;
            const consumir = () => { e.preventDefault(); e.stopImmediatePropagation(); };
            if (e.key === 'ArrowDown') { consumir(); marcar(Math.min(destaque + 1, itens.length - 1)); }
            else if (e.key === 'ArrowUp') { consumir(); marcar(Math.max(destaque - 1, 0)); }
            else if (e.key === 'Tab' && !e.shiftKey) { consumir(); aceitar(destaque >= 0 ? destaque : 0); }
            else if (e.key === 'Enter' && destaque >= 0) { consumir(); aceitar(destaque); }
            else if (e.key === 'Escape') { consumir(); fechar(); }
        }, true);
        raiz.addEventListener('scroll', e => { if (caixa && e.target && caixa.contains && caixa.contains(e.target)) return; fechar(); }, true);
        raiz.addEventListener('resize', fechar);
        const b = doc.getElementById('chip-sugestoes');
        if (b) b.addEventListener('click', () => definir(!ligado()));
        atualizarBotao();
    }

    raiz.AtivoSugestao = Object.assign({ iniciar, ligado, definir }, api);
})(typeof window !== 'undefined' ? window : globalThis);
