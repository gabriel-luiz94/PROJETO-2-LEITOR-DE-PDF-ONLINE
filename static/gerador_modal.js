/*
 * gerador_modal.js (TASK-068) — interface dos modais geradores de Cabos e de Postes/Estruturas:
 * painéis de atalhos (com "＋ Novo", computador/projeto), favoritos ★ e linhas estilo Excel (selecionar, copiar, colar).
 * Lógica pura em atalhos_modal.js e grade_tabela.js (testadas em Node); aqui só a ligação com a tela.
 * Quem abre/insere na tabela é o resumo.js (usa ctl.linhas).
 */
(function (raiz) {
    'use strict';
    const AM = raiz.AtalhosModal, GT = raiz.GradeTabela;
    const doc = raiz.document;

    function el(tag, props, filhos) {
        const e = doc.createElement(tag);
        Object.keys(props || {}).forEach(k => {
            if (k === 'style') e.style.cssText = props[k];
            else if (k === 'text') e.textContent = props[k];
            else if (k.indexOf('on') === 0) e.addEventListener(k.slice(2), props[k]);
            else e.setAttribute(k, props[k]);
        });
        (filhos || []).forEach(f => f && e.appendChild(typeof f === 'string' ? doc.createTextNode(f) : f));
        return e;
    }
    const toast = m => { if (raiz.showToast) raiz.showToast(m); };
    const projetoAtual = () => { try { return raiz.localStorage.getItem('projeto_selecionado_codigo') || ''; } catch (e) { return ''; } };

    /** Janela simples: campos + botões. Resolve { botao, valores } ou null (Esc/fora). */
    function dialogo(cfg) {
        return new Promise(resolve => {
            const fundo = el('div', { class: 'gm-dialogo', style: 'position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:10001;display:flex;align-items:center;justify-content:center;' });
            const caixa = el('div', { style: 'background:#0d1117;color:#e6edf3;border:1px solid #30363d;border-radius:10px;padding:18px;max-width:420px;width:92%;font-family:Inter,sans-serif;' });
            caixa.appendChild(el('h3', { text: cfg.titulo, style: 'margin:0 0 8px;font-size:1rem;' }));
            if (cfg.texto) caixa.appendChild(el('p', { text: cfg.texto, style: 'margin:0 0 10px;font-size:.8rem;color:#8b949e;line-height:1.4;' }));
            const campos = {};
            (cfg.campos || []).forEach(c => {
                const linha = el('div', { style: 'margin-bottom:10px;' });
                linha.appendChild(el('label', { text: c.rotulo, style: 'display:block;font-size:.72rem;color:#8b949e;margin-bottom:3px;' }));
                if (c.tipo === 'escopo') {
                    const grupo = el('div', { style: 'display:flex;gap:14px;font-size:.82rem;' });
                    const rs = [['local', 'Só neste computador'], ['projeto', 'Projeto (nuvem)']].map(([v, t]) => {
                        const r = el('input', { type: 'radio', name: 'gm-escopo', value: v });
                        if (v === (c.valor || 'local')) r.checked = true;
                        if (v === 'projeto' && !projetoAtual()) { r.disabled = true; t += ' — selecione um projeto'; }
                        if (c.travado) r.disabled = true;
                        grupo.appendChild(el('label', { style: 'display:flex;gap:5px;align-items:center;cursor:pointer;' }, [r, t]));
                        return r;
                    });
                    campos[c.nome] = () => (rs.find(r => r.checked) || rs[0]).value;
                    linha.appendChild(grupo);
                } else {
                    const inp = el('input', { type: c.tipo || 'text', value: c.valor == null ? '' : c.valor, placeholder: c.placeholder || '',
                        style: 'width:100%;padding:6px 8px;background:#161b22;color:#e6edf3;border:1px solid #30363d;border-radius:6px;box-sizing:border-box;' });
                    campos[c.nome] = () => inp.value;
                    linha.appendChild(inp);
                    if (!campos.__primeiro) campos.__primeiro = inp;
                }
                caixa.appendChild(linha);
            });
            const botoes = el('div', { style: 'display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;margin-top:6px;' });
            const fim = v => { fundo.remove(); doc.removeEventListener('keydown', tecla, true); resolve(v); };
            const tecla = e => {
                if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); fim(null); }
                if (e.key === 'Enter' && e.target.tagName === 'INPUT' && e.target.type !== 'radio') { e.preventDefault(); e.stopPropagation(); const p = (cfg.botoes || []).find(b => b.padrao); if (p) clicar(p); }
            };
            const clicar = b => {
                const valores = {};
                Object.keys(campos).forEach(k => { if (k !== '__primeiro') valores[k] = campos[k](); });
                fim({ botao: b.valor, valores });
            };
            (cfg.botoes || []).forEach(b => botoes.appendChild(el('button', { class: b.perigo ? 'btn-secondary' : (b.padrao ? 'btn-primary' : 'btn-secondary'), text: b.rotulo,
                style: b.perigo ? 'color:#f85149;border-color:rgba(248,81,73,.5);' : '', onclick: () => clicar(b) })));
            caixa.appendChild(botoes);
            fundo.appendChild(caixa);
            doc.body.appendChild(fundo);
            doc.addEventListener('keydown', tecla, true);
            setTimeout(() => campos.__primeiro && campos.__primeiro.focus(), 30);
        });
    }

    /* ───────────── comportamentos de cada modal ───────────── */
    const cacheDesc = {};
    async function buscarDescricao(ativo) {
        const q = String(ativo || '').trim().replace(/\s+/g, '');
        if (!q) return '';
        if (cacheDesc[q] !== undefined) return cacheDesc[q];
        let d;
        try {
            const res = await fetch('/api/orcamento/search?q=' + encodeURIComponent(q) + '&col=ativo');
            const data = await res.json();
            d = data.resultados && data.resultados.length ? (data.resultados[0].desc_ativo || 'Ativo encontrado') : 'Ativo não encontrado';
        } catch (e) { return 'Erro na busca'; }
        cacheDesc[q] = d;
        return d;
    }

    const CONFIGS = {
        cabos: {
            colunas: [
                { campo: 'op', tipo: 'op', rotulo: 'OP', larg: '9%' },
                { campo: 'ativo', tipo: 'ativo', rotulo: 'ATIVO (ex: CAA 2)', larg: '30%' },
                { campo: 'fase', tipo: 'fase', rotulo: 'FASE', larg: '10%' },
                { campo: 'comp', tipo: 'comp', rotulo: 'COMP. (m)', larg: '10%' },
                { campo: 'qtd', tipo: 'qtd', rotulo: 'QTD', larg: '8%' }
            ],
            nova(prev, op) { return { op: op || (prev ? prev.op : 'I'), ativo: '', fase: prev ? prev.fase : 'ABC', comp: prev ? prev.comp : 80, qtd: 1, desc: '', compManual: false }; },
            inicial() { return this.nova(null, 'I'); },
            clique(ctl, grupo, item) {
                let alvos = ctl.alvos();
                if (!alvos.length) { ctl.linhas.push(this.nova(null, 'I')); alvos = [ctl.linhas.length - 1]; ctl.selecionarLinhas(alvos); }
                alvos.forEach(r => {
                    const l = ctl.linhas[r];
                    if (grupo.id === 'ativos') { l.ativo = item.valor; if (!l.compManual) l.comp = item.comp != null ? item.comp : AM.compPadrao(item.valor); ctl.pedirDesc(r); }
                    else if (grupo.id === 'fases') l.fase = item.valor;
                    else if (grupo.id === 'comps') { l.comp = parseFloat(String(item.valor).replace(',', '.')) || l.comp; l.compManual = true; }
                });
            },
            linhaDeFavorito(f) { const l = f.linha; return { op: l.op || 'I', ativo: l.ativo || '', fase: l.fase || 'ABC', comp: l.comp || 80, qtd: 1, desc: '', compManual: true }; },
            favoritoDe(l) { return { op: l.op, ativo: (l.ativo || '').trim(), fase: l.fase, comp: l.comp }; }
        },
        postes: {
            colunas: [
                { campo: 'op', tipo: 'op', rotulo: 'OP', larg: '9%' },
                { campo: 'ativo', tipo: 'ativo', rotulo: 'ATIVO (ex: DT11/300 1-N3)', larg: '50%' },
                { campo: 'qtd', tipo: 'qtd', rotulo: 'QTD', larg: '9%' }
            ],
            nova(prev, op) { return { op: op || (prev ? prev.op : 'I'), ativo: '', qtd: 1, desc: '' }; },
            inicial() { return this.nova(null, 'I'); },
            clique(ctl, grupo, item) {
                if (grupo.id === 'postes') {          // poste = linha nova (ou preenche a última se estiver em branco)
                    const last = ctl.linhas[ctl.linhas.length - 1];
                    if (!last || last.ativo.trim() !== '') ctl.linhas.push(this.nova(last, null));
                    const r = ctl.linhas.length - 1;
                    ctl.linhas[r].ativo = item.valor;
                    ctl.selecionarLinhas([r]);
                    ctl.pedirDesc(r);
                    return;
                }
                let alvos = ctl.alvos();                 // estruturas acumulam na(s) linha(s) em destaque
                if (!alvos.length) { ctl.linhas.push(this.nova(null, 'I')); alvos = [0]; ctl.selecionarLinhas(alvos); }
                alvos.forEach(r => {
                    const l = ctl.linhas[r], partes = l.ativo.trim() ? l.ativo.trim().split(/\s+/) : [];
                    const rx = new RegExp('^(\\d+)-' + item.valor.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$');
                    let achou = false;
                    for (let i = 0; i < partes.length; i++) { const m = partes[i].match(rx); if (m) { partes[i] = (parseInt(m[1], 10) + 1) + '-' + item.valor; achou = true; break; } }
                    if (!achou) partes.push('1-' + item.valor);
                    l.ativo = partes.join(' ');
                    ctl.pedirDesc(r);
                });
            },
            linhaDeFavorito(f) { const l = f.linha; return { op: l.op || 'I', ativo: l.ativo || '', qtd: 1, desc: '' }; },
            favoritoDe(l) { return { op: l.op, ativo: (l.ativo || '').trim() }; }
        }
    };

    /* ───────────── controlador de um modal ───────────── */
    function criar(modal, ids) {
        const cfg = CONFIGS[modal], grupos = AM.GRUPOS[modal];
        const overlay = doc.getElementById(ids.overlay), paineis = doc.getElementById(ids.paineis), tbody = doc.getElementById(ids.tbody);
        const ctl = { modal, linhas: [], sel: null, edicao: false, undo: [], local: AM.lerLocal(modal), projetoDoc: { itens: [], favoritos: [] } };
        const colunas = cfg.colunas;
        const aberto = () => overlay && !overlay.classList.contains('hidden') && !doc.querySelector('.gm-dialogo');

        /* seleção / alvos */
        ctl.alvos = function () {
            if (ctl.sel) {
                const r1 = Math.min(ctl.sel.a.r, ctl.sel.f.r), r2 = Math.max(ctl.sel.a.r, ctl.sel.f.r), out = [];
                for (let r = r1; r <= r2; r++) if (ctl.linhas[r]) out.push(r);
                if (out.length) return out;
            }
            return ctl.linhas.length ? [ctl.linhas.length - 1] : [];       // sem seleção: a última
        };
        ctl.selecionarLinhas = function (rs) {
            if (!rs.length) { ctl.sel = null; return; }
            ctl.sel = { a: { r: Math.min.apply(null, rs), c: 0 }, f: { r: Math.max.apply(null, rs), c: colunas.length - 1 }, linhas: true };
        };
        function marcar() {
            const alvos = new Set(ctl.alvos());
            const iv = ctl.sel ? { r1: Math.min(ctl.sel.a.r, ctl.sel.f.r), r2: Math.max(ctl.sel.a.r, ctl.sel.f.r), c1: Math.min(ctl.sel.a.c, ctl.sel.f.c), c2: Math.max(ctl.sel.a.c, ctl.sel.f.c) } : null;
            const multi = iv && (iv.r2 > iv.r1 || iv.c2 > iv.c1);
            Array.from(tbody.children).forEach(tr => {
                const r = parseInt(tr.dataset.r);
                tr.classList.toggle('gm-alvo', alvos.has(r));
                tr.querySelectorAll('td[data-col]').forEach(td => {
                    const c = parseInt(td.dataset.col);
                    td.classList.toggle('gm-sel', !!(multi && r >= iv.r1 && r <= iv.r2 && c >= iv.c1 && c <= iv.c2));
                });
            });
        }

        /* histórico local do modal */
        const guardar = () => { ctl.undo.push(JSON.stringify(ctl.linhas)); if (ctl.undo.length > 40) ctl.undo.shift(); };
        ctl.desfazer = function () {
            if (!ctl.undo.length) { toast('Nada para desfazer neste modal.'); return; }
            ctl.linhas = JSON.parse(ctl.undo.pop());
            if (ctl.sel && !ctl.linhas[ctl.sel.f.r]) ctl.sel = null;
            ctl.render();
        };

        /* descrição automática */
        ctl.pedirDesc = function (r) {
            const l = ctl.linhas[r];
            if (!l) return;
            const ativo = l.ativo;
            buscarDescricao(ativo).then(d => {
                if (ctl.linhas[r] !== l || l.ativo !== ativo) return;
                l.desc = d;
                const td = tbody.querySelector('tr[data-r="' + r + '"] .gm-desc');
                if (td) td.textContent = d;
            });
        };

        /* tabela */
        ctl.render = function () {
            tbody.innerHTML = '';
            const frag = doc.createDocumentFragment();
            ctl.linhas.forEach((l, r) => {
                const tr = el('tr', { 'data-r': r });
                tr.appendChild(el('td', { class: 'gm-num', text: String(r + 1), title: 'Clique para selecionar a linha (Shift estende)' }));
                colunas.forEach((c, ci) => {
                    const td = el('td', { 'data-col': ci });
                    let campo;
                    if (c.tipo === 'op' || c.tipo === 'fase') {
                        campo = el('select', { class: 'gm-ctl' });
                        (c.tipo === 'op' ? AM.OPERACOES : GT.FASES).forEach(v => { const o = el('option', { value: v, text: v }); if (v === l[c.campo]) o.selected = true; campo.appendChild(o); });
                        campo.addEventListener('change', () => { guardar(); l[c.campo] = campo.value; });
                    } else {
                        campo = el('input', { class: 'gm-ctl', type: c.tipo === 'ativo' ? 'text' : 'number', value: l[c.campo] == null ? '' : l[c.campo], autocomplete: 'off' });
                        if (c.tipo === 'qtd') campo.min = '1';
                        let antes = null;
                        campo.addEventListener('focus', () => { antes = JSON.stringify(ctl.linhas); });
                        campo.addEventListener('input', () => {
                            if (c.tipo === 'ativo') { l.ativo = campo.value; }
                            else if (c.tipo === 'comp') { l.comp = parseFloat(campo.value) || 0; l.compManual = true; }
                            else if (c.tipo === 'qtd') { l.qtd = parseInt(campo.value, 10) || 1; }
                        });
                        campo.addEventListener('change', () => {
                            if (antes !== null && antes !== JSON.stringify(ctl.linhas)) { ctl.undo.push(antes); if (ctl.undo.length > 40) ctl.undo.shift(); }
                            if (c.tipo === 'ativo') {
                                if (modal === 'cabos' && !l.compManual && l.ativo.trim()) {
                                    l.comp = AM.compPadrao(l.ativo);
                                    const ic = tr.querySelector('td[data-col="3"] input'); if (ic) ic.value = l.comp;
                                }
                                ctl.pedirDesc(r);
                            }
                        });
                        if (c.tipo === 'ativo') campo.addEventListener('keydown', e => {
                            if (e.key === 'Enter') { e.preventDefault(); if (!focar(r + 1, ci)) { guardar(); ctl.linhas.push(cfg.nova(ctl.linhas[r], null)); ctl.render(); focar(r + 1, ci); } }
                            else if (e.key === 'ArrowDown' && !e.shiftKey) { e.preventDefault(); focar(r + 1, ci); }
                            else if (e.key === 'ArrowUp' && !e.shiftKey) { e.preventDefault(); focar(r - 1, ci); }
                        });
                    }
                    campo.addEventListener('focus', () => { ctl.sel = { a: { r, c: ci }, f: { r, c: ci } }; marcar(); });
                    td.appendChild(campo);
                    tr.appendChild(td);
                });
                tr.appendChild(el('td', { class: 'gm-desc', text: l.desc || '' }));
                tr.appendChild(el('td', { class: 'gm-acoes' }, [
                    el('button', { class: 'gm-btn-fav', title: 'Salvar esta linha como favorito', text: '★', onclick: () => ctl.salvarFavorito(r) }),
                    el('button', { class: 'gm-btn-del', title: 'Excluir linha', text: '✖', onclick: () => { guardar(); ctl.linhas.splice(r, 1); ctl.sel = null; ctl.render(); } })
                ]));
                frag.appendChild(tr);
            });
            tbody.appendChild(frag);
            marcar();
        };
        function focar(r, ci) {
            const tr = tbody.querySelector('tr[data-r="' + r + '"]');
            const c = tr && tr.querySelector('td[data-col="' + ci + '"] .gm-ctl');
            if (!c) return false;
            c.focus(); if (c.select && c.tagName === 'INPUT') c.select();
            return true;
        }

        ctl.adicionarLinha = function (op) {
            guardar();
            ctl.linhas.push(cfg.nova(ctl.linhas[ctl.linhas.length - 1], op || null));
            ctl.selecionarLinhas([ctl.linhas.length - 1]);
            ctl.render();
        };
        ctl.limparEdicao = function () { ctl.render(); };

        /* ───────────── painéis ───────────── */
        function dadosMesclados() {
            return AM.mesclar(modal, { projeto: ctl.projetoDoc, local: ctl.local, ocultos: ctl.local.ocultos }, ctl.edicao);
        }
        ctl.renderPaineis = function () {
            const m = dadosMesclados();
            paineis.innerHTML = '';
            const barra = el('div', { class: 'gm-barra' }, [
                el('span', { class: 'gm-hint', text: 'Atalhos: a linha em destaque (amarela) é a que ativo/fase/comprimento alteram; operação cria linha nova.' }),
                el('span', { style: 'flex:1' }),
                el('button', { class: 'gm-mini' + (ctl.edicao ? ' ativo' : ''), text: ctl.edicao ? '✔ Concluir edição' : '✎ Editar atalhos', onclick: () => { ctl.edicao = !ctl.edicao; ctl.renderPaineis(); } }),
                el('button', { class: 'gm-mini', text: '↶ Desfazer', title: 'Desfaz a última alteração das linhas deste modal (Ctrl+Z)', onclick: () => ctl.desfazer() })
            ]);
            paineis.appendChild(barra);
            const grade = el('div', { class: 'gm-paineis' });
            // operação: cria linha nova
            const pOp = painel('OPERAÇÃO', '#a371f7', 'Cada clique = nova linha');
            AM.OPERACOES.forEach(op => pOp.corpo.appendChild(el('button', { class: 'btn-cable-preset gm-op', text: op, style: 'border-color:rgba(163,113,247,.5);color:#a371f7;', onclick: () => ctl.adicionarLinha(op) })));
            grade.appendChild(pOp.raiz);
            grupos.forEach(g => {
                const p = painel(g.titulo, g.cor, g.dica, () => ctl.novoAtalho(g));
                (m.grupos[g.id] || []).forEach(it => p.corpo.appendChild(botaoAtalho(g, it)));
                grade.appendChild(p.raiz);
            });
            paineis.appendChild(grade);
            // favoritos
            const pf = painel('FAVORITOS ★', '#f0883e', 'Clique para inserir a linha pronta', null, () => ctl.salvarFavorito(ctl.alvos()[0]));
            if (!m.favoritos.length) pf.corpo.appendChild(el('span', { class: 'gm-vazio', text: 'Nenhum favorito ainda. Use a ★ de uma linha ou "★ Salvar linha em destaque".' }));
            m.favoritos.forEach(f => pf.corpo.appendChild(botaoFavorito(f)));
            const faixa = el('div', { class: 'gm-paineis', style: 'grid-template-columns:1fr;margin-top:8px;' }, [pf.raiz]);
            paineis.appendChild(faixa);
            if (ctl.edicao && m.ocultosCount) paineis.appendChild(el('div', { class: 'gm-hint', style: 'margin-top:4px;', text: m.ocultosCount + ' atalho(s) oculto(s) neste computador aparecem esmaecidos; clique para mostrar de novo.' }));
        };
        function painel(titulo, cor, dica, aoNovo, aoSalvar) {
            const cab = el('div', { class: 'gm-painel-cab' }, [el('span', { text: titulo, style: 'color:' + cor + ';font-weight:600;font-size:.72rem;letter-spacing:.05em;' })]);
            if (aoNovo) cab.appendChild(el('button', { class: 'gm-mini', text: '＋ Novo', onclick: aoNovo }));
            if (aoSalvar) cab.appendChild(el('button', { class: 'gm-mini', text: '★ Salvar linha em destaque', onclick: aoSalvar }));
            const corpo = el('div', { class: 'gm-painel-corpo' });
            const raiz = el('div', { class: 'gm-painel', style: 'border-color:' + cor + '55;background:' + cor + '0d;' }, [cab, el('div', { class: 'gm-dica', text: dica }), corpo]);
            return { raiz, corpo };
        }
        const marcaOrigem = o => o === 'projeto' ? ' ☁' : o === 'local' ? ' ⌂' : '';
        function botaoAtalho(g, it) {
            const t = it.valor + (g.numerico ? '' : '') + marcaOrigem(it.origem);
            const b = el('button', { class: 'btn-cable-preset' + (it.oculto ? ' gm-oculto' : ''), text: t,
                title: (it.origem === 'projeto' ? 'Atalho do projeto' : it.origem === 'local' ? 'Atalho deste computador' : 'Atalho de fábrica') + (it.comp ? ' · ' + it.comp + ' m' : ''),
                style: 'border-color:' + g.cor + '66;color:' + g.cor + ';' });
            b.addEventListener('click', () => {
                if (ctl.edicao) { ctl.editarAtalho(g, it); return; }
                guardar();
                cfg.clique(ctl, g, it);
                ctl.render();
            });
            return b;
        }
        function botaoFavorito(f) {
            const b = el('button', { class: 'btn-cable-preset', text: '★ ' + AM.rotuloFavorito(modal, f) + marcaOrigem(f.origem),
                title: 'Inserir esta linha (' + (f.origem === 'projeto' ? 'projeto' : 'computador') + ')', style: 'border-color:rgba(240,136,62,.5);color:#f0883e;' });
            b.addEventListener('click', () => {
                if (ctl.edicao) { ctl.editarFavorito(f); return; }
                guardar();
                ctl.linhas.push(cfg.linhaDeFavorito(f));
                ctl.selecionarLinhas([ctl.linhas.length - 1]);
                ctl.pedirDesc(ctl.linhas.length - 1);
                ctl.render();
            });
            return b;
        }

        /* ───────────── persistência dos atalhos ───────────── */
        async function persistir(escopo, mutar) {
            if (escopo === 'projeto') {
                const projeto = projetoAtual();
                const copia = JSON.parse(JSON.stringify(ctl.projetoDoc || { itens: [], favoritos: [] }));
                mutar(copia);
                const r = await AM.gravarProjeto(modal, projeto, copia);
                if (r.ok) { ctl.projetoDoc = r.dados; return true; }
                if (r.rede) {
                    toast('Sem conexão: o item foi guardado só neste computador.');
                    return persistir('local', mutar);
                }
                toast(r.erro);
                return false;
            }
            mutar(ctl.local);
            if (!AM.gravarLocal(modal, ctl.local)) { toast('Não foi possível gravar neste navegador.'); return false; }
            return true;
        }
        const camposEscopo = (valor, travado) => ({ nome: 'escopo', rotulo: 'Onde guardar', tipo: 'escopo', valor, travado });

        ctl.novoAtalho = async function (g) {
            const campos = [{ nome: 'valor', rotulo: 'Valor do atalho', placeholder: g.numerico ? 'ex: 45' : 'ex: ' + (g.id === 'ativos' ? 'CAA 2' : '...') }];
            if (g.comComp) campos.push({ nome: 'comp', rotulo: 'Comprimento padrão (m) — opcional', tipo: 'number', placeholder: 'vazio = 40 (multiplex) ou 80' });
            campos.push(camposEscopo('local'));
            const r = await dialogo({ titulo: 'Novo atalho — ' + g.titulo, campos, botoes: [{ rotulo: 'Cancelar', valor: 'x' }, { rotulo: 'Criar', valor: 'ok', padrao: true }] });
            if (!r || r.botao !== 'ok') return;
            const valor = String(r.valores.valor || '').trim();
            if (!valor || valor.length > 40 || /[<>]/.test(valor)) { toast('Valor inválido (até 40 caracteres, sem < >).'); return; }
            if (g.numerico && !(parseFloat(valor.replace(',', '.')) > 0)) { toast('Informe um número maior que zero.'); return; }
            const comp = parseFloat(String(r.valores.comp || '').replace(',', '.'));
            const item = { id: AM.novoId(), grupo: g.id, valor };
            if (g.comComp && comp > 0) item.comp = comp;
            if (await persistir(r.valores.escopo, d => { d.itens = (d.itens || []).concat([item]); })) { toast('Atalho criado.'); ctl.renderPaineis(); }
        };
        ctl.editarAtalho = async function (g, it) {
            if (it.origem === 'fabrica') {
                const r = await dialogo({ titulo: it.valor, texto: 'Atalho de fábrica: não pode ser apagado, mas pode ficar oculto neste computador.',
                    botoes: [{ rotulo: 'Cancelar', valor: 'x' }, { rotulo: it.oculto ? 'Mostrar de novo' : 'Ocultar', valor: 'ocultar', padrao: true }] });
                if (r && r.botao === 'ocultar') alternarOculto(it.id);
                return;
            }
            const campos = [{ nome: 'valor', rotulo: 'Valor do atalho', valor: it.valor }];
            if (g.comComp) campos.push({ nome: 'comp', rotulo: 'Comprimento padrão (m)', tipo: 'number', valor: it.comp || '' });
            const r = await dialogo({ titulo: 'Editar atalho (' + (it.origem === 'projeto' ? 'projeto' : 'este computador') + ')', campos,
                botoes: [{ rotulo: 'Excluir', valor: 'del', perigo: true }, { rotulo: it.oculto ? 'Mostrar' : 'Ocultar aqui', valor: 'ocultar' }, { rotulo: 'Cancelar', valor: 'x' }, { rotulo: 'Salvar', valor: 'ok', padrao: true }] });
            if (!r) return;
            if (r.botao === 'ocultar') { alternarOculto(it.id); return; }
            const escopo = it.origem;
            if (r.botao === 'del') {
                if (await persistir(escopo, d => { d.itens = (d.itens || []).filter(x => x.id !== it.id); })) { toast('Atalho excluído.'); ctl.renderPaineis(); }
            } else if (r.botao === 'ok') {
                const valor = String(r.valores.valor || '').trim();
                if (!valor || valor.length > 40 || /[<>]/.test(valor)) { toast('Valor inválido.'); return; }
                const comp = parseFloat(String(r.valores.comp || '').replace(',', '.'));
                if (await persistir(escopo, d => { d.itens = (d.itens || []).map(x => { if (x.id !== it.id) return x; const n = Object.assign({}, x, { valor }); if (g.comComp) { if (comp > 0) n.comp = comp; else delete n.comp; } return n; }); })) ctl.renderPaineis();
            }
        };
        function alternarOculto(id) {
            const s = new Set(ctl.local.ocultos || []);
            if (s.has(id)) s.delete(id); else s.add(id);
            ctl.local.ocultos = Array.from(s);
            AM.gravarLocal(modal, ctl.local);
            ctl.renderPaineis();
        }
        ctl.salvarFavorito = async function (r) {
            const l = ctl.linhas[r];
            if (!l || !String(l.ativo || '').trim()) { toast('Escolha uma linha com ativo preenchido para salvar como favorito.'); return; }
            const res = await dialogo({ titulo: 'Salvar favorito ★', texto: AM.rotuloFavorito(modal, { linha: cfg.favoritoDe(l) }),
                campos: [{ nome: 'nome', rotulo: 'Nome (opcional)', placeholder: 'em branco = mostra a linha' }, camposEscopo('local')],
                botoes: [{ rotulo: 'Cancelar', valor: 'x' }, { rotulo: 'Salvar', valor: 'ok', padrao: true }] });
            if (!res || res.botao !== 'ok') return;
            const fav = { id: AM.novoId(), linha: cfg.favoritoDe(l) };
            const nome = String(res.valores.nome || '').trim();
            if (nome) fav.nome = nome.slice(0, 60);
            if (await persistir(res.valores.escopo, d => { d.favoritos = (d.favoritos || []).concat([fav]); })) { toast('Favorito salvo.'); ctl.renderPaineis(); }
        };
        ctl.editarFavorito = async function (f) {
            const res = await dialogo({ titulo: 'Editar favorito', texto: AM.rotuloFavorito(modal, f), campos: [{ nome: 'nome', rotulo: 'Nome', valor: f.nome || '' }],
                botoes: [{ rotulo: 'Excluir', valor: 'del', perigo: true }, { rotulo: 'Cancelar', valor: 'x' }, { rotulo: 'Salvar', valor: 'ok', padrao: true }] });
            if (!res) return;
            if (res.botao === 'del') { if (await persistir(f.origem, d => { d.favoritos = (d.favoritos || []).filter(x => x.id !== f.id); })) ctl.renderPaineis(); }
            else if (res.botao === 'ok') {
                const nome = String(res.valores.nome || '').trim().slice(0, 60);
                if (await persistir(f.origem, d => { d.favoritos = (d.favoritos || []).map(x => { if (x.id !== f.id) return x; const n = Object.assign({}, x); if (nome) n.nome = nome; else delete n.nome; return n; }); })) ctl.renderPaineis();
            }
        };

        /* ───────────── abrir ───────────── */
        ctl.abrir = function () {
            ctl.linhas = [cfg.inicial()];
            ctl.sel = null; ctl.undo = []; ctl.edicao = false;
            ctl.local = AM.lerLocal(modal);
            const projeto = projetoAtual();
            ctl.projetoDoc = projeto ? AM.lerCacheProjeto(modal, projeto) : { itens: [], favoritos: [] };
            ctl.renderPaineis();
            ctl.render();
            overlay.classList.remove('hidden');
            if (projeto) AM.carregarProjeto(modal, projeto).then(d => { ctl.projetoDoc = d; if (aberto()) ctl.renderPaineis(); });
        };

        /* ───────────── seleção com o mouse, copiar, colar, teclado ───────────── */
        let arrastando = false;
        const celulaDe = alvo => {
            const td = alvo && alvo.closest ? alvo.closest('td[data-col]') : null;
            if (!td || !tbody.contains(td)) return null;
            return { r: parseInt(td.parentElement.dataset.r), c: parseInt(td.dataset.col) };
        };
        tbody.addEventListener('mousedown', e => {
            if (e.button !== 0) return;
            const num = e.target.closest ? e.target.closest('td.gm-num') : null;
            if (num) {
                e.preventDefault();
                const r = parseInt(num.parentElement.dataset.r);
                if (e.shiftKey && ctl.sel) ctl.selecionarLinhas([ctl.sel.a.r, r]); else ctl.selecionarLinhas([r]);
                if (doc.activeElement && doc.activeElement.blur) doc.activeElement.blur();
                marcar();
                return;
            }
            const cel = celulaDe(e.target);
            if (!cel) return;
            if (e.shiftKey && ctl.sel) { e.preventDefault(); ctl.sel.f = cel; delete ctl.sel.linhas; if (doc.activeElement && doc.activeElement.blur) doc.activeElement.blur(); marcar(); return; }
            ctl.sel = { a: cel, f: cel };
            arrastando = true;
            marcar();
        });
        tbody.addEventListener('mousemove', e => {
            if (!arrastando || !(e.buttons & 1)) { arrastando = false; return; }
            const cel = celulaDe(e.target);
            if (!cel || (cel.r === ctl.sel.f.r && cel.c === ctl.sel.f.c)) return;
            ctl.sel.f = cel;
            e.preventDefault();
            const s = raiz.getSelection && raiz.getSelection(); if (s) s.removeAllRanges();
            if (doc.activeElement && tbody.contains(doc.activeElement)) doc.activeElement.blur();
            marcar();
        });
        doc.addEventListener('mouseup', () => { arrastando = false; });

        const campoEmEdicao = () => { const a = doc.activeElement; return !!a && tbody.contains(a) && a.tagName === 'INPUT'; };
        function intervalo() {
            if (!ctl.sel) return null;
            return { r1: Math.min(ctl.sel.a.r, ctl.sel.f.r), r2: Math.max(ctl.sel.a.r, ctl.sel.f.r), c1: Math.min(ctl.sel.a.c, ctl.sel.f.c), c2: Math.max(ctl.sel.a.c, ctl.sel.f.c) };
        }
        const multi = iv => !!iv && (iv.r2 > iv.r1 || iv.c2 > iv.c1);

        doc.addEventListener('copy', e => {
            if (!aberto()) return;
            const iv = intervalo();
            if (!iv || (!multi(iv) && campoEmEdicao())) return;
            e.preventDefault();
            e.clipboardData.setData('text/plain', GT.textoDoIntervaloModal(ctl.linhas, colunas, iv.r1, iv.c1, iv.r2, iv.c2));
            toast('Copiado (' + (iv.r2 - iv.r1 + 1) + ' linha(s)).');
        });
        doc.addEventListener('paste', e => {
            if (!aberto()) return;
            const alvo = e.target, naGrade = tbody.contains(alvo), corpo = alvo === doc.body || alvo === doc.documentElement || overlay.contains(alvo) && !naGrade && alvo.tagName !== 'INPUT';
            if (!naGrade && !(corpo && ctl.sel)) return;
            const texto = (e.clipboardData || raiz.clipboardData).getData('text');
            if (!texto) return;
            const iv = intervalo();
            if (!GT.ehColagemMultipla(texto) && !(corpo && iv)) return;       // valor único num campo: colagem nativa
            const cel = naGrade ? celulaDe(alvo) : null;
            let linha, coluna;
            if (iv && (multi(iv) || !cel)) { linha = iv.r1; coluna = iv.c1; } else if (cel) { linha = cel.r; coluna = cel.c; } else return;
            e.preventDefault();
            guardar();
            const r = GT.aplicarColagemModal(ctl.linhas, texto, { colunas, linha, coluna, nova: p => cfg.nova(p, null) });
            if (!r.linhasColadas) { ctl.undo.pop(); toast('Nada para colar.'); return; }
            r.alteradas.forEach(i => {
                const l = ctl.linhas[i];
                if (modal === 'cabos' && !l.compManual && l.ativo.trim()) l.comp = AM.compPadrao(l.ativo);
                l.desc = '';
            });
            ctl.sel = { a: { r: r.alteradas[0], c: 0 }, f: { r: r.alteradas[r.alteradas.length - 1], c: colunas.length - 1 }, linhas: true };
            ctl.render();
            r.alteradas.slice(0, 300).forEach(i => ctl.pedirDesc(i));
            let msg = r.linhasColadas + ' linha(s) coladas' + (r.criadas ? ', ' + r.criadas + ' nova(s)' : '') + '.';
            if (r.invalidos) msg += ' ' + r.invalidos + ' valor(es) inválido(s) ignorado(s).';
            if (r.truncado) msg += ' Limite de linhas atingido.';
            if (r.alteradas.length > 300) msg += ' Descrições buscadas só nas 300 primeiras.';
            toast(msg);
        });
        doc.addEventListener('keydown', e => {
            if (!aberto()) return;
            const mod = e.ctrlKey || e.metaKey, iv = intervalo();
            if (mod && !e.shiftKey && (e.key === 'z' || e.key === 'Z')) {
                if (campoEmEdicao() && doc.activeElement.value !== undefined && ctl.undo.length === 0) return;
                if (campoEmEdicao()) return;                       // dentro do campo: desfazer nativo do texto
                e.preventDefault(); e.stopPropagation(); ctl.desfazer(); return;
            }
            if (mod && (e.key === 'a' || e.key === 'A') && !campoEmEdicao() && ctl.linhas.length) {
                e.preventDefault();
                ctl.sel = { a: { r: 0, c: 0 }, f: { r: ctl.linhas.length - 1, c: colunas.length - 1 }, linhas: true };
                marcar(); return;
            }
            if ((e.key === 'Delete' || e.key === 'Backspace') && iv && (multi(iv) || !campoEmEdicao()) && !(doc.activeElement && doc.activeElement.tagName === 'SELECT')) {
                e.preventDefault();
                guardar();
                if (ctl.sel.linhas) {
                    ctl.linhas.splice(iv.r1, iv.r2 - iv.r1 + 1);
                    if (!ctl.linhas.length) ctl.linhas.push(cfg.inicial());
                    ctl.sel = null;
                } else {
                    const ia = colunas.findIndex(c => c.tipo === 'ativo');
                    for (let r = iv.r1; r <= iv.r2; r++) if (ia >= iv.c1 && ia <= iv.c2) { ctl.linhas[r].ativo = ''; ctl.linhas[r].desc = ''; }
                }
                ctl.render();
            }
        }, true);

        // fechar pelo X ou clicando fora não limpa nada; o estado é refeito em abrir()
        return ctl;
    }

    raiz.GeradorModal = { criar, dialogo };
})(typeof window !== 'undefined' ? window : globalThis);
