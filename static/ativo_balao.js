/*
 * ativo_balao.js (TASK-186) — balão INTERATIVO do ativo em Outros: lista os ativos da linha (com o que já está preenchido e o que falta) e edita os parâmetros de cada um.
 *   · abre com Ctrl+Enter no campo, duplo clique, mouse parado ~1,5 s na célula, ou ao digitar "(" com «Ocultar parâmetros» ligado (abre direto nos parâmetros do ativo);
 *   · lista de ativos: ↑↓ e Tab/Shift+Tab navegam, Enter (ou clique) entra nos parâmetros do ativo, Esc fecha e devolve o foco ao campo;
 *   · parâmetros: Tab/Shift+Tab andam entre os parâmetros, ↑↓ escolhem o valor, Enter grava e vai ao próximo (no último, ao próximo ativo com pendência; sem pendência, fecha), Esc volta à lista.
 * A tela (resumo.js) fornece as pontas em iniciar(cfg): itens(input), dica(nome), origens(input, nome), gravar(input, i, nomeParam, valor), carregarBase(), mensagem(texto).
 */
(function (raiz) {
    'use strict';
    const doc = raiz.document;
    let cfg = null, balao = null, S = null, timerHover = 0, timerFecha = 0;
    const esc = t => String(t == null ? '' : t).replace(/[&<>"]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[m]));
    const DEMORA_HOVER = 1500;

    function pendente(it) { return !!(it.status && (!it.status.completo || it.status.naoExiste)); }

    /** Os parâmetros editáveis do item i: {d, params:[{nome, padrao, valores, aberto, atual, rotulo}]} | {msg} | {carregando}. */
    function paramsDe(i) {
        const itens = cfg.itens(S.input), it = itens[i];
        if (!it) return { msg: 'Ativo não encontrado na linha.' };
        const d = cfg.dica(it.nome);
        if (!d) return { carregando: true };
        if (d.tipo !== 'parametros') return { msg: d.mensagem || (it.nome + ': ativo sem parâmetros declarados') };
        const AD = raiz.AtivoDica, AI = raiz.AtivoItens;
        const editaveis = AD.parametrosEditaveis(d);
        const escritos = AD.lerPedido(AI.conteudo(it.params), d.params.map(p => p.nome)).valores;
        const orig = (cfg.origens && cfg.origens(S.input, String(d.ativo || it.nome).toUpperCase())) || {};
        return {
            d, it, params: editaveis.map(p => {
                const N = p.nome.toUpperCase(), escrito = escritos[N], o = orig[N];
                return { nome: p.nome, padrao: p.padrao, valores: p.valores.slice(), aberto: p.aberto !== false, atual: escrito != null && escrito !== '' ? escrito : (o ? o.valor : null), manual: escrito != null && escrito !== '', rotulo: escrito ? 'manual' : (o ? o.rotulo : '') };
            })
        };
    }
    const opcoes = p => { const l = p.valores.map(v => ({ v, padrao: p.padrao != null && String(p.padrao).toUpperCase() === String(v).toUpperCase() })); if (p.aberto || !l.length) l.push({ custom: true }); return l; };
    function valorInicial(p) {
        const l = opcoes(p), atual = p.atual == null ? null : String(p.atual).toUpperCase();
        let k = atual == null ? -1 : l.findIndex(o => !o.custom && String(o.v).toUpperCase() === atual);
        if (k < 0 && atual != null) k = l.findIndex(o => o.custom);
        if (k < 0) k = l.findIndex(o => o.padrao);
        return Math.max(0, k);
    }

    function resumoDoItem(it) {
        const s = it.status;
        if (!s) return '<span class="m">sem parâmetros</span>';
        if (s.naoExiste) return '<span class="f">ativo não existe na base</span>';
        return s.params.map(p => (p.falta || p.invalido ? `<u class="f">${esc(p.nome)} ?</u>` : `${esc(p.nome)} = ${esc(p.valor)}`)).join(' · ') || '<span class="m">sem parâmetros</span>';
    }

    function desenhar() {
        if (!S) return;
        if (S.foco) { const sg = doc.querySelector('.sugestao-ativos'); if (sg) sg.style.display = 'none'; }
        if (!balao) {
            balao = doc.createElement('div'); balao.className = 'ab-balao';
            balao.addEventListener('mousedown', e => { if (!(e.target.closest && e.target.closest('input'))) e.preventDefault(); });          // clicar no balão não tira o foco do campo
            balao.addEventListener('click', aoClicar);
            balao.addEventListener('mouseenter', () => { clearTimeout(timerFecha); });
            balao.addEventListener('mouseleave', () => { if (S && S.hover) agendarFechar(); });
            doc.body.appendChild(balao);
        }
        const itens = cfg.itens(S.input);
        let h = '';
        if (S.modo === 'ativos') {
            if (S.i >= itens.length) S.i = Math.max(0, itens.length - 1);
            h += `<div class="ab-h">Linha ${esc(String(Number(S.input.dataset.parIdx) + 1))} · ${itens.length} ativo(s)</div>`;
            if (!itens.length) h += '<div class="ab-vazio">Escreva um ativo na linha.</div>';
            itens.forEach((it, i) => {
                const s = it.status, ic = !s ? '<span class="m">–</span>' : (s.naoExiste ? '<span class="f">✖</span>' : (s.completo ? '<span class="v">✓</span>' : '<span class="f">⚠</span>'));
                h += `<div class="ab-it${i === S.i ? ' on' : ''}" data-i="${i}">${ic}<b>${esc(it.nome)}</b><span class="ab-r">${resumoDoItem(it)}</span></div>`;
            });
            h += '<div class="ab-ft"><kbd>↑↓</kbd>/<kbd>Tab</kbd> navega · <kbd>Enter</kbd> parâmetros · <kbd>Esc</kbd> fecha</div>';
        } else {
            const it = itens[S.i], pd = paramsDe(S.i);
            h += `<div class="ab-h">Linha ${esc(String(Number(S.input.dataset.parIdx) + 1))} › <b>${esc(it ? it.nome : '')}</b> · parâmetros</div>`;
            if (pd.carregando) { h += '<div class="ab-vazio">Carregando a base…</div>'; if (cfg.carregarBase) Promise.resolve(cfg.carregarBase()).then(() => { if (S && S.modo === 'params') desenhar(); }); }
            else if (pd.msg) h += `<div class="ab-vazio">${esc(pd.msg)}</div>`;
            else {
                if (S.p >= pd.params.length) S.p = Math.max(0, pd.params.length - 1);
                pd.params.forEach((p, k) => {
                    const val = p.atual == null ? '<span class="f">⚠ escolha…</span>' : `<span class="v">✓ ${esc(p.atual)}</span>`;
                    h += `<div class="ab-p${k === S.p ? ' on' : ''}" data-p="${k}"><b>${esc(p.nome)}</b>${val}<span class="ab-r m">${esc(p.rotulo || (p.atual == null ? 'a preencher' : ''))}</span></div>`;
                    if (k === S.p) {
                        h += '<div class="ab-vals">';
                        opcoes(p).forEach((o, j) => {
                            if (o.custom) h += `<div class="ab-v${j === S.v ? ' on' : ''}" data-v="${j}">✎ <input class="ab-in" type="text" placeholder="digitar outro valor…" value="${esc(S.livre || '')}" autocomplete="off" spellcheck="false"></div>`;
                            else h += `<div class="ab-v${j === S.v ? ' on' : ''}" data-v="${j}">${esc(o.v)}${o.padrao ? '<span class="m"> padrão</span>' : ''}</div>`;
                        });
                        h += '</div>';
                    }
                });
                if (!pd.params.length) h += '<div class="ab-vazio">Este ativo não tem parâmetros editáveis.</div>';
            }
            h += '<div class="ab-ft"><kbd>Tab</kbd> parâmetro · <kbd>↑↓</kbd> valor · <kbd>Enter</kbd> escolhe e vai ao próximo · <kbd>Del</kbd> limpa · <kbd>Esc</kbd> volta</div>';
        }
        balao.innerHTML = h;
        balao.style.display = 'block';
        const r = S.input.getBoundingClientRect(), w = Math.min(520, raiz.innerWidth - 16);
        balao.style.width = w + 'px';
        balao.style.left = Math.max(8, Math.min(r.left, raiz.innerWidth - w - 8)) + 'px';
        const alto = balao.offsetHeight;
        balao.style.top = (r.bottom + alto + 8 > raiz.innerHeight && r.top - alto - 6 > 8 ? r.top - alto - 6 : r.bottom + 3) + 'px';
        const on = balao.querySelector('.on'); if (on && on.scrollIntoView) on.scrollIntoView({ block: 'nearest' });
        const inp = balao.querySelector('.ab-v.on .ab-in');
        if (inp && S.foco && S.focoLivre) { inp.focus(); S.focoLivre = false; }
    }

    function abrir(input, opcoesAbrir) {
        const o = opcoesAbrir || {};
        clearTimeout(timerHover); clearTimeout(timerFecha);
        if (!cfg || !input) return;
        const itens = cfg.itens(input);
        if (!itens.length && !o.modo) { if (cfg.mensagem) cfg.mensagem('Escreva um ativo na linha primeiro.'); return; }
        const sg = doc.querySelector('.sugestao-ativos'); if (sg) sg.style.display = 'none';                 // a lista de sugestões de ativos não fica por cima do balão
        S = { input, modo: o.modo || 'ativos', i: o.i != null ? o.i : Math.max(0, itens.findIndex(pendente)), p: 0, v: 0, foco: !!o.foco, hover: !o.foco, livre: '' };
        if (S.modo === 'params') iniciarParams();
        desenhar();
    }
    function iniciarParams() {
        const pd = paramsDe(S.i);
        S.p = 0; S.livre = '';
        if (pd.params && pd.params.length) { const primeiro = Math.max(0, pd.params.findIndex(p => p.atual == null)); S.p = primeiro; S.v = valorInicial(pd.params[S.p]); }
    }
    function fechar(devolver) {
        clearTimeout(timerHover); clearTimeout(timerFecha);
        const in_ = S && S.input; S = null;
        if (balao) balao.style.display = 'none';
        if (devolver && in_) { in_.focus(); }
    }
    function agendarFechar() { clearTimeout(timerFecha); timerFecha = setTimeout(() => { if (S && S.hover) fechar(false); }, 350); }

    /** "(" no campo compacto: abre nos parâmetros do ativo sob (ou logo antes de) o cursor. */
    function abrirNoCursor(input) {
        if (!cfg) return false;
        const AI = raiz.AtivoItens, t = AI.tokens(input.value), c = input.selectionStart == null ? input.value.length : input.selectionStart;
        if (!t.length) { if (cfg.mensagem) cfg.mensagem('Escreva o nome do ativo antes de abrir os parâmetros.'); return true; }
        let i = t.findIndex(k => c >= k.ini && c <= k.fim);
        if (i < 0) { i = -1; t.forEach((k, n) => { if (k.fim <= c) i = n; }); if (i < 0) i = 0; }
        abrir(input, { foco: true, modo: 'params', i });
        return true;
    }

    function irParaProximoComPendencia(aPartirDe) {
        const itens = cfg.itens(S.input), n = itens.length;
        for (let k = 1; k <= n; k++) { const j = (aPartirDe + k) % n; if (j !== aPartirDe && pendente(itens[j])) return j; }
        return -1;
    }
    function aceitarValor() {
        const pd = paramsDe(S.i);
        if (!pd.params || !pd.params.length) { S.modo = 'ativos'; desenhar(); return; }
        const p = pd.params[S.p], op = opcoes(p)[S.v];
        let valor = null;
        if (op && op.custom) { const inp = balao && balao.querySelector('.ab-in'); valor = inp ? inp.value.trim() : ''; if (valor === '') valor = null; }
        else if (op) valor = op.v;
        if (valor != null && (String(valor) !== String(p.atual == null ? '' : p.atual) || !p.manual)) cfg.gravar(S.input, S.i, p.nome, valor);
        S.livre = '';
        const novo = paramsDe(S.i);
        if (novo.params && S.p < novo.params.length - 1) { S.p++; S.v = valorInicial(novo.params[S.p]); S.focoLivre = false; desenhar(); return; }
        const prox = irParaProximoComPendencia(S.i);                                                // acabou os parâmetros: próximo ativo com pendência; sem pendência, fecha
        if (prox >= 0) { S.i = prox; iniciarParams(); desenhar(); } else fechar(true);
    }
    function limparValor() {
        const pd = paramsDe(S.i);
        if (!pd.params || !pd.params.length) return;
        cfg.gravar(S.input, S.i, pd.params[S.p].nome, '');
        desenhar();
    }
    function moverValor(passo) {
        const pd = paramsDe(S.i);
        if (!pd.params || !pd.params.length) return;
        const n = opcoes(pd.params[S.p]).length;
        S.v = (S.v + passo + n) % n;
        S.focoLivre = !!opcoes(pd.params[S.p])[S.v].custom;
        desenhar();
        if (!S.focoLivre && S.input) S.input.focus();
    }
    function moverParametro(passo) {
        const pd = paramsDe(S.i);
        if (!pd.params || !pd.params.length) return;
        S.p = (S.p + passo + pd.params.length) % pd.params.length; S.v = valorInicial(pd.params[S.p]); S.livre = ''; S.focoLivre = false;
        desenhar();
        if (S.input) S.input.focus();
    }

    function aoClicar(e) {
        if (!S) return;
        const t = e.target;
        const v = t.closest && t.closest('[data-v]'); if (v && S.modo === 'params') { S.v = Number(v.getAttribute('data-v')); const pd = paramsDe(S.i), o = pd.params && opcoes(pd.params[S.p])[S.v]; if (o && o.custom) { S.focoLivre = true; desenhar(); } else aceitarValor(); return; }
        const p = t.closest && t.closest('[data-p]'); if (p && S.modo === 'params') { const pd = paramsDe(S.i); S.p = Number(p.getAttribute('data-p')); S.v = valorInicial(pd.params[S.p]); S.livre = ''; desenhar(); return; }
        const it = t.closest && t.closest('[data-i]'); if (it && S.modo === 'ativos') { S.i = Number(it.getAttribute('data-i')); S.modo = 'params'; S.hover = false; S.foco = true; iniciarParams(); desenhar(); S.input.focus(); }
    }

    function aoTeclar(e) {
        if (!S || e.isComposing) return;
        const alvoOk = e.target === S.input || (balao && balao.contains(e.target));
        if (!alvoOk || !S.foco) return;
        const k = e.key, consumir = () => { e.preventDefault(); e.stopImmediatePropagation(); };
        if (e.ctrlKey || e.metaKey || e.altKey) { if (k === 'Enter' && e.ctrlKey) { consumir(); } return; }
        if (S.modo === 'ativos') {
            const n = cfg.itens(S.input).length;
            if (k === 'ArrowDown' || (k === 'Tab' && !e.shiftKey)) { consumir(); if (n) { S.i = (S.i + 1) % n; desenhar(); } }
            else if (k === 'ArrowUp' || (k === 'Tab' && e.shiftKey)) { consumir(); if (n) { S.i = (S.i - 1 + n) % n; desenhar(); } }
            else if (k === 'Enter') { consumir(); if (n) { S.modo = 'params'; iniciarParams(); desenhar(); } }
            else if (k === 'Escape') { consumir(); fechar(true); }
            else if (k.length === 1 || k === 'Backspace' || k === 'Delete') fechar(false);                  // voltou a digitar: o balão sai da frente
            return;
        }
        const noLivre = e.target.classList && e.target.classList.contains('ab-in');
        if (k === 'Escape') { consumir(); S.modo = 'ativos'; S.livre = ''; desenhar(); S.input.focus(); return; }
        if (k === 'Tab') { consumir(); moverParametro(e.shiftKey ? -1 : 1); return; }
        if (k === 'Enter') { consumir(); aceitarValor(); return; }
        if (k === 'ArrowDown') { consumir(); moverValor(1); return; }
        if (k === 'ArrowUp') { consumir(); moverValor(-1); return; }
        if (k === 'Delete' && !noLivre) { consumir(); limparValor(); return; }
        if (noLivre) { S.livre = e.target.value; setTimeout(() => { if (S && balao) { const inp = balao.querySelector('.ab-in'); if (inp) S.livre = inp.value; } }, 0); return; }
        if (k.length === 1) { consumir(); }                                                                // letras soltas não vão para o texto da célula enquanto edita parâmetros
    }

    function iniciar(opcoesIniciar) {
        cfg = opcoesIniciar || {};
        doc.addEventListener('keydown', aoTeclar, true);
        doc.addEventListener('mousedown', e => { if (S && !(balao && balao.contains(e.target)) && e.target !== S.input) fechar(false); }, true);
        doc.addEventListener('focusout', e => { if (S && S.foco && e.target === S.input) setTimeout(() => { if (S && doc.activeElement !== S.input && !(balao && balao.contains(doc.activeElement))) fechar(false); }, 120); });
        doc.addEventListener('dblclick', e => { const el = e.target && e.target.closest ? e.target.closest('input[data-par-tipo="outros"]') : null; if (el) { e.preventDefault(); abrir(el, { foco: true }); } });
        doc.addEventListener('mouseover', e => {
            const el = e.target && e.target.closest ? e.target.closest('[data-par-tipo="outros"]') : null;
            if (!el || (S && S.input === el)) return;
            clearTimeout(timerHover);
            if (S && S.hover) fechar(false);
            if (doc.activeElement === el && S) return;
            timerHover = setTimeout(() => { if (!S || S.hover) abrir(el, { foco: false }); }, DEMORA_HOVER);
        });
        doc.addEventListener('mouseout', e => {
            const el = e.target && e.target.closest ? e.target.closest('[data-par-tipo="outros"]') : null;
            if (!el) return;
            clearTimeout(timerHover);
            if (S && S.hover && S.input === el && !(e.relatedTarget && balao && balao.contains(e.relatedTarget))) agendarFechar();
        });
        raiz.addEventListener('scroll', e => { if (S && !(balao && e.target && balao.contains && balao.contains(e.target))) fechar(false); }, true);
        raiz.addEventListener('resize', () => fechar(false));
    }

    raiz.AtivoBalao = { iniciar, abrir, abrirNoCursor, fechar, aberto: () => !!S };
})(typeof window !== 'undefined' ? window : globalThis);
