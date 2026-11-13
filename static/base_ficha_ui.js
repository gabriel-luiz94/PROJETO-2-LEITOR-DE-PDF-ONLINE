/*
 * base_ficha_ui.js (TASK-146) — barra de pesquisa fixa no topo (Ctrl+K foca nela) e FICHA DO ATIVO na tela da Base de Orçamento: mostra tudo do ativo (itens, parâmetros e variantes,
 * quem ele chama / quem o chama, verificações, simulador) e navega entre ativos (Ctrl+clique no ATIVO/CÓDIGO da tabela, links dentro da ficha).
 * Lê a base inteira do servidor e troca a fatia visível na tela pelas linhas da tela (edições ainda não salvas aparecem). Só leitura: editar continua na tabela.
 */
(function (raiz) {
    'use strict';
    const doc = raiz.document, D = raiz.BaseFichaDados;
    let cfg = { tbody: null, projeto: () => '', visao: () => 'projeto', admin: () => false, recarregar: async () => {} };
    let baseCache = null, baseEm = 0, pilha = [], tabAtual = 'itens', modeloAtual = null;
    const U = v => String(v == null ? '' : v).trim().toUpperCase();

    const css = `
    .bf-ov{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:9000;display:flex;align-items:flex-start;justify-content:center}
    .bf-md{background:var(--bg-surface,#12171f);color:var(--text-primary,#e6edf3);border:1px solid var(--border-default,#262e3d);border-radius:10px;box-shadow:0 20px 60px #000a;display:flex;flex-direction:column;overflow:hidden;font-size:13px}
    .bf-hd{display:flex;align-items:center;gap:8px;padding:10px 14px;border-bottom:1px solid var(--border-default,#262e3d);flex-wrap:wrap}.bf-hd h2{margin:0;font-size:16px}
    .bf-chip{display:inline-block;padding:2px 8px;border-radius:99px;font-size:11px;background:var(--bg-subtle,#1a202c);border:1px solid var(--border-default,#262e3d);color:var(--text-secondary,#919bb0)}
    .bf-chip.o{color:#ffb77a;border-color:#7a3d10;background:#2a1a0e}.bf-chip.k{color:#7ee787;border-color:#1f5a2b;background:#0f2415}.bf-chip.w{color:#e3b341;border-color:#5c4a12;background:#241d0a}.bf-chip.r{color:#ff938a;border-color:#6b1f1b;background:#2a1010}
    .bf-sp{flex:1}.bf-b{background:var(--btn-neutral-bg,#161c26);border:1px solid var(--border-default,#262e3d);color:var(--text-primary,#e6edf3);border-radius:6px;padding:4px 10px;font-size:12px;cursor:pointer}.bf-b:hover{border-color:#d97706}
    .bf-tabs{display:flex;gap:2px;padding:0 10px;border-bottom:1px solid var(--border-default,#262e3d)}.bf-tab{padding:8px 12px;color:var(--text-secondary,#919bb0);border-bottom:2px solid transparent;cursor:pointer}.bf-tab.a{color:var(--text-primary,#e6edf3);border-color:#d97706}
    .bf-body{display:flex;flex:1;min-height:0}.bf-main{flex:1;overflow:auto;padding:10px}.bf-side{background:#0f141c;border-left:1px solid var(--border-default,#262e3d);padding:12px;width:290px;overflow:auto}
    .bf-side h4{margin:0 0 6px;font-size:11px;text-transform:uppercase;color:var(--text-secondary,#919bb0);letter-spacing:.05em}
    .bf-tb{border-collapse:collapse;width:100%}.bf-tb th{background:var(--bg-subtle,#1a202c);color:var(--text-secondary,#919bb0);text-align:left;padding:5px 8px;border:1px solid var(--border-default,#262e3d);font-size:11px;text-transform:uppercase}
    .bf-tb td{padding:4px 8px;border:1px solid var(--border-default,#262e3d);white-space:nowrap}.bf-tb td.n{text-align:right;font-variant-numeric:tabular-nums}.bf-tb tr.g td{background:#161d29;color:#58a6ff;font-weight:600}
    .bf-tb tr.s td{background:#1d2a44}.bf-mono{font-family:Consolas,monospace}.bf-link{color:#58a6ff;cursor:pointer;text-decoration:underline}.bf-mut{color:var(--text-secondary,#919bb0)}
    .bf-kv{display:flex;justify-content:space-between;padding:3px 0;border-bottom:1px dashed var(--border-default,#262e3d);gap:8px}
    .bf-in{background:var(--bg-canvas,#090d12);border:1px solid var(--border-default,#262e3d);color:var(--text-primary,#e6edf3);border-radius:6px;padding:8px 12px;font:inherit;width:100%;box-sizing:border-box}
    .bf-res{padding:7px 12px;display:flex;gap:8px;align-items:center;cursor:pointer}.bf-res.a,.bf-res:hover{background:var(--bg-hover,#242c3d)}
    .bf-ok{color:#7ee787}.bf-av{color:#e3b341}.bf-er{color:#ff938a}
    .bf-bar{position:sticky;top:0;z-index:8500;background:var(--bg-surface,#12171f);border:1px solid var(--border-default,#262e3d);border-radius:8px;box-shadow:0 4px 14px #0006}
    .bf-bar:focus-within{border-color:#d97706}.bf-bar-in{display:flex;align-items:center;gap:10px;padding:5px 12px}.bf-bar-ic{opacity:.75;font-size:14px}
    .bf-bar-inp{flex:1;min-width:0;background:transparent;border:0;outline:0;color:var(--text-primary,#e6edf3);font:inherit;font-size:14px;padding:7px 0}.bf-bar-inp::placeholder{color:var(--text-secondary,#919bb0)}
    .bf-bar-k{font:11px Consolas,monospace;color:var(--text-secondary,#919bb0);border:1px solid var(--border-default,#262e3d);border-radius:4px;padding:1px 6px;background:var(--bg-subtle,#1a202c)}
    .bf-bar-x{background:transparent;border:0;color:var(--text-secondary,#919bb0);cursor:pointer;font-size:13px;padding:2px 6px;border-radius:4px}.bf-bar-x:hover{color:var(--text-primary,#e6edf3);background:var(--bg-hover,#242c3d)}
    .bf-drop{position:absolute;left:0;right:0;top:100%;margin-top:4px;background:var(--bg-surface,#12171f);border:1px solid var(--border-default,#262e3d);border-radius:8px;max-height:min(60vh,520px);overflow:auto;box-shadow:0 14px 40px #000a}
    .bf-flash td{animation:bfpisca 1.6s ease-out}@keyframes bfpisca{0%,40%{background:#3b2a0a}100%{background:transparent}}`;
    function estilo() { if (doc.getElementById('bf-css')) return; const s = doc.createElement('style'); s.id = 'bf-css'; s.textContent = css; doc.head.appendChild(s); }

    function el(tag, props, filhos) {
        const e = doc.createElement(tag);
        Object.keys(props || {}).forEach(k => {
            if (k === 'class') e.className = props[k]; else if (k === 'text') e.textContent = props[k]; else if (k === 'style') e.style.cssText = props[k];
            else if (k.indexOf('on') === 0) e.addEventListener(k.slice(2), props[k]); else e.setAttribute(k, props[k]);
        });
        (filhos || []).forEach(f => f != null && f !== false && e.appendChild(typeof f === 'string' ? doc.createTextNode(f) : f));
        return e;
    }
    const chip = (t, c) => el('span', { class: 'bf-chip' + (c ? ' ' + c : ''), text: t });

    /* ───────── dados ───────── */
    function linhasDaTela() {
        if (!cfg.tbody) return [];
        return Array.from(cfg.tbody.querySelectorAll('tr')).map(tr => {
            const i = tr.querySelectorAll('.cell-input');
            if (i.length < 13) return null;
            return { ativo: i[0].value, desc_ativo: i[1].value, componente: i[2].value, projeto: i[3].value, mdo: i[4].value, origem: i[5].value, codigo: i[6].value, desc_codigo: i[7].value,
                fator_i: i[8].value, fator_r: i[9].value, filtro: i[10].value, lv: i[11].value, arred: i[12].value, formula_i: '', formula_r: '' };
        }).filter(Boolean);
    }
    async function modelo(forcar) {
        if (!baseCache || forcar || Date.now() - baseEm > 45000) { try { baseCache = (await (await raiz.fetch('/api/orcamento/dados')).json()).dados || []; baseEm = Date.now(); } catch (e) { baseCache = baseCache || []; } }
        const proj = cfg.projeto(), visao = cfg.visao();
        let linhas = baseCache;
        if (visao === 'projeto' && proj) linhas = baseCache.filter(r => U(r.projeto) !== U(proj)).concat(linhasDaTela());
        else if (visao === 'genericas') linhas = baseCache.filter(r => U(r.projeto) !== '').concat(linhasDaTela());
        modeloAtual = D.construir(linhas, visao === 'projeto' ? proj : '');
        return modeloAtual;
    }
    function invalidar() { baseCache = null; }

    /* ───────── janela ───────── */
    function limparJanelas() { doc.querySelectorAll('.bf-ov').forEach(o => o.remove()); doc.removeEventListener('keydown', aoTeclar, true); }
    function fecharTudo() { limparJanelas(); pilha = []; }
    function aoTeclar(e) { if (e.key === 'Escape') { e.stopPropagation(); fecharTudo(); } }
    function janela(largura, topo) {
        estilo(); limparJanelas();
        const ov = el('div', { class: 'bf-ov' }), md = el('div', { class: 'bf-md', style: `width:${largura};margin-top:${topo}` });
        ov.appendChild(md); ov.addEventListener('mousedown', e => { if (e.target === ov) fecharTudo(); });
        doc.body.appendChild(ov); doc.addEventListener('keydown', aoTeclar, true);
        return md;
    }

    /* ───────── barra de pesquisa FIXA no topo da tela (Ctrl+K leva o cursor para ela) ───────── */
    let barra = null;
    function montarBarra(host) {
        estilo();
        host.className = 'bf-bar'; host.innerHTML = '';
        const inp = el('input', { class: 'bf-bar-inp', type: 'text', placeholder: 'Pesquisar ativo, apelido, código ou descrição…', autocomplete: 'off', spellcheck: 'false', 'aria-label': 'Pesquisar na Base de Orçamento' });
        const limpar = el('button', { class: 'bf-bar-x', type: 'button', title: 'Limpar a pesquisa', text: '✕', style: 'display:none' });
        const drop = el('div', { class: 'bf-drop', style: 'display:none' });
        host.append(el('div', { class: 'bf-bar-in' }, [el('span', { class: 'bf-bar-ic', text: '🔍' }), inp, limpar, el('kbd', { class: 'bf-bar-k', text: 'Ctrl+K' })]), drop);
        let m = null, itens = [], sel = 0, seq = 0;
        const fechar = () => { drop.style.display = 'none'; };
        const atualizar = () => { limpar.style.display = inp.value ? '' : 'none'; };
        async function carregar() { const meu = ++seq; const novo = await modelo(); if (meu === seq) m = novo; return m; }      // base e linhas da tela de agora (edições ainda não salvas entram)
        const abrir = r => { if (!r) return; fechar(); if (r.tipo === 'ativo') abrirFicha(r.nome); else abrirCodigo(r.codigo); };
        function desenhar() {
            drop.innerHTML = ''; drop.style.display = '';
            if (!m) { drop.appendChild(el('div', { class: 'bf-mut', style: 'padding:12px', text: 'Carregando a base…' })); return; }
            if (!inp.value.trim()) { drop.appendChild(el('div', { class: 'bf-mut', style: 'padding:12px', text: 'Digite para pesquisar na base' + (m.projeto ? ' do projeto ' + m.projeto : '') + ' (e nas linhas comuns). ↑↓ navegam, Enter abre, Esc fecha.' })); itens = []; return; }
            itens = D.pesquisar(m, inp.value, 14);
            if (!itens.length) { drop.appendChild(el('div', { class: 'bf-er', style: 'padding:12px', text: 'Nada encontrado.' })); return; }
            sel = Math.min(sel, itens.length - 1);
            itens.forEach((r, i) => {
                const linha = r.tipo === 'ativo'
                    ? el('div', { class: 'bf-res' + (i === sel ? ' a' : '') }, [el('b', { text: r.nome }), ...r.projetos.slice(0, 1).map(p => chip(p, 'o')), r.via ? chip('apelido: ' + r.via) : null, el('span', { class: 'bf-mut', text: r.descricao }), el('span', { class: 'bf-sp' }), el('span', { class: 'bf-mut', text: `${r.linhas} linhas · ${r.codigos} códigos` })])
                    : el('div', { class: 'bf-res' + (i === sel ? ' a' : '') }, [el('span', { class: 'bf-mut', text: 'código' }), el('b', { class: 'bf-mono', text: r.codigo }), el('span', { text: r.descricao }), el('span', { class: 'bf-sp' }), el('span', { class: 'bf-mut', text: `usado em ${r.em.join(', ')}${r.usos > r.em.length ? '…' : ''}` })]);
                linha.addEventListener('mousedown', e => { e.preventDefault(); abrir(r); });
                drop.appendChild(linha);
            });
            const ativa = drop.querySelector('.bf-res.a'); if (ativa && ativa.scrollIntoView) ativa.scrollIntoView({ block: 'nearest' });
        }
        inp.addEventListener('focus', () => { desenhar(); carregar().then(desenhar); });
        inp.addEventListener('input', () => { sel = 0; atualizar(); if (m) desenhar(); else { desenhar(); carregar().then(desenhar); } });
        inp.addEventListener('keydown', e => {
            if (e.key === 'ArrowDown') { e.preventDefault(); if (drop.style.display === 'none') desenhar(); sel = Math.min(sel + 1, Math.max(itens.length - 1, 0)); desenhar(); }
            else if (e.key === 'ArrowUp') { e.preventDefault(); sel = Math.max(sel - 1, 0); desenhar(); }
            else if (e.key === 'Enter') { e.preventDefault(); abrir(itens[sel]); }
            else if (e.key === 'Escape') { e.stopPropagation(); if (drop.style.display !== 'none') fechar(); else inp.blur(); }
        });
        limpar.addEventListener('mousedown', e => { e.preventDefault(); inp.value = ''; sel = 0; atualizar(); inp.focus(); desenhar(); });
        doc.addEventListener('mousedown', e => { if (!host.contains(e.target)) fechar(); });
        barra = { focar(consulta) { if (consulta != null) { inp.value = consulta; sel = 0; atualizar(); } inp.focus(); inp.select(); } };
    }
    /** Ctrl+K: leva o cursor para a barra do topo (com `consulta`, já preenche o texto). */
    function abrirPesquisa(consulta) { if (barra) barra.focar(consulta); }

    /* ───────── código ───────── */
    async function abrirCodigo(codigo) {
        const m = modeloAtual || await modelo(), u = D.usoDoCodigo(m, codigo);
        const md = janela('min(640px,94vw)', '80px');
        md.appendChild(el('div', { class: 'bf-hd' }, [el('h2', { class: 'bf-mono', text: codigo }), ...(u ? u.descricoes.map(d => chip(d)) : []), el('span', { class: 'bf-sp' }), el('button', { class: 'bf-b', text: '✕', onclick: fecharTudo })]));
        const corpo = el('div', { class: 'bf-main', style: 'max-height:60vh' });
        if (!u) corpo.appendChild(el('div', { class: 'bf-er', text: 'Código não encontrado na base.' }));
        else {
            if (u.descricoes.length > 1) corpo.appendChild(el('div', { class: 'bf-av', style: 'margin-bottom:8px', text: `⚠ ${u.descricoes.length} descrições diferentes para este código.` }));
            const tb = el('table', { class: 'bf-tb' }, [el('tr', {}, ['Ativo', 'Condição', 'MDO', 'Fator I', 'Fator R'].map(h => el('th', { text: h })))]);
            u.usos.forEach(x => tb.appendChild(el('tr', {}, [el('td', {}, [el('span', { class: 'bf-link', text: x.ativo, onclick: () => abrirFicha(x.ativo) })]), el('td', { class: 'bf-mono', text: x.condicao || '—' }), el('td', { text: x.mdo }), el('td', { class: 'n', text: x.fator_i }), el('td', { class: 'n', text: x.fator_r })])));
            corpo.appendChild(tb);
        }
        md.appendChild(corpo);
    }

    /* ───────── ficha ───────── */
    async function abrirFicha(nome, voltando) {
        const m = modeloAtual || await modelo();
        const f = D.ficha(m, nome);
        if (!f) { raiz.showToast && raiz.showToast(`Ativo "${nome}" não encontrado na base${m.projeto ? ' deste projeto' : ''}.`); return; }
        if (!voltando) { if (!pilha.length || pilha[pilha.length - 1] !== f.nome) pilha.push(f.nome); tabAtual = 'itens'; }
        desenharFicha(m, f);
    }
    function desenharFicha(m, f) {
        const md = janela('min(1200px,96vw)', '28px'); md.style.height = 'min(780px,92vh)';
        const erros = f.verificacoes.filter(v => v.nivel === 'erro').length, avisos = f.verificacoes.filter(v => v.nivel === 'aviso').length;
        md.appendChild(el('div', { class: 'bf-hd' }, [
            pilha.length > 1 ? el('button', { class: 'bf-b', text: '← ' + pilha[pilha.length - 2], onclick: () => { pilha.pop(); abrirFicha(pilha[pilha.length - 1], true); } }) : null,
            el('h2', { text: f.nome }), ...f.projetos.slice(0, 2).map(p => chip(p, 'o')), ...f.apelidos.map(a => chip('apelido ' + a)),
            f.parametros.length ? chip('parâmetros: ' + f.parametros.map(p => p.nome.toLowerCase() + (p.padrao ? '=' + p.padrao : '')).join(' · ')) : null,
            chip(erros ? '✖ ' + erros + ' erro(s)' : avisos ? '⚠ ' + avisos + ' aviso(s)' : '✔ sem problemas', erros ? 'r' : avisos ? 'w' : 'k'),
            el('span', { class: 'bf-sp' }),
            cfg.admin() && cfg.projeto() && raiz.BaseCadastro ? el('button', { class: 'bf-b', style: 'background:#d97706;border-color:#d97706;color:#111;font-weight:600', text: 'Editar', title: 'Abre o cadastro rápido com a definição atual deste ativo (admin)', onclick: () => { fecharTudo(); raiz.BaseCadastro.abrir({ editar: f.nome }); } }) : null,
            cfg.admin() && cfg.projeto() ? el('button', { class: 'bf-b', id: 'bf-desfazer', style: 'display:none', text: '↶ Desfazer última publicação', onclick: () => desfazerDaFicha(f.nome) }) : null,
            el('button', { class: 'bf-b', text: 'Ir para a tabela', title: 'Fecha e destaca a primeira linha deste ativo na tabela', onclick: () => { fecharTudo(); irParaLinha(f.nome); } }),
            el('button', { class: 'bf-b', text: '✕', onclick: fecharTudo })]));
        if (cfg.admin() && cfg.projeto()) raiz.fetch(`/api/admin/ativo/versao?projeto=${encodeURIComponent(cfg.projeto())}&nome=${encodeURIComponent(f.nome)}`, { headers: { Authorization: 'Bearer ' + (raiz.localStorage.getItem('auth_token') || '') } }).then(r => r.json()).then(v => { const b = doc.getElementById('bf-desfazer'); if (b && v.tem_anterior) { b.style.display = ''; b.title = 'Volta ao que o ativo era antes da última publicação' + (v.anterior_por ? ' (por ' + v.anterior_por + ')' : ''); } }).catch(() => {});
        const abas = [['itens', `Itens (${f.nLinhas})`], ['param', 'Parâmetros e variantes'], ['rel', `Chama / É chamado por (${f.chama.length + f.chamadoPor.length})`], ['sim', 'Simulador']];
        const tabs = el('div', { class: 'bf-tabs' }), corpo = el('div', { class: 'bf-body' });
        abas.forEach(([id, rot]) => tabs.appendChild(el('span', { class: 'bf-tab' + (id === tabAtual ? ' a' : ''), text: rot, onclick: () => { tabAtual = id; desenharFicha(m, f); } })));
        md.append(tabs, corpo);
        const main = el('div', { class: 'bf-main' }), lado = el('div', { class: 'bf-side' });
        corpo.append(main, lado);
        if (tabAtual === 'itens') abaItens(m, f, main, lado); else if (tabAtual === 'param') abaParametros(f, main, lado); else if (tabAtual === 'rel') abaRelacoes(f, main, lado); else abaSimulador(f, main, lado);
        if (tabAtual !== 'itens') lateralVerificacoes(f, lado);
    }
    function linhaItem(m, l, ref) {
        const tr = el('tr', { 'data-cod': l.codigo });
        const compost = l.itens ? el('td', { colspan: 7 }, [el('span', { style: 'color:#58a6ff', text: ':= ' }), ...l.itens.flatMap((i, k) => [k ? ' + ' : '', `${i.qtd} × `, el('span', { class: 'bf-link', text: D.nomeDoItem(i.ativo), onclick: () => abrirFicha(D.nomeDoItem(i.ativo)) }), '(' + (i.ativo.indexOf('(') !== -1 ? i.ativo.slice(i.ativo.indexOf('(') + 1) : ')')])]) : null;
        tr.append(el('td', { class: 'bf-mono bf-mut', text: l.condicao || '—' }));
        if (compost) tr.appendChild(compost);
        else tr.append(el('td', { text: l.mdo.indexOf('MAO') === 0 || U(l.mdo) === 'N' ? 'N' : 'M' }), el('td', { class: 'bf-mono' }, [el('span', { class: 'bf-link', text: l.codigo, onclick: () => ref(l.codigo) })]), el('td', { text: l.desc_codigo }), el('td', { class: 'n', text: l.fator_i }), el('td', { class: 'n', text: l.fator_r }), el('td', { text: l.lv }), el('td', { text: l.filtro }));
        return tr;
    }
    function abaItens(m, f, main, lado) {
        const cab = ['Condição', 'MDO', 'Código', 'Descrição', 'Fator I', 'Fator R', 'LV', 'Filtro'];
        const tb = el('table', { class: 'bf-tb' }, [el('tr', {}, cab.map(h => el('th', { text: h })))]);
        const mostrar = cod => { desenharLateralCodigo(m, f, cod, lado); tb.querySelectorAll('tr.s').forEach(t => t.classList.remove('s')); tb.querySelectorAll('tr[data-cod="' + cod + '"]').forEach(t => t.classList.add('s')); };
        const grupo = (rot, ls) => { if (!ls.length) return; tb.appendChild(el('tr', { class: 'g' }, [el('td', { colspan: 8, text: rot })])); ls.forEach(l => tb.appendChild(linhaItem(m, l, mostrar))); };
        grupo('Itens comuns — valem sempre', f.comuns);
        grupo('Por parâmetro / chamadas a outros ativos', f.variantes);
        main.appendChild(tb);
        main.appendChild(el('div', { class: 'bf-mut', style: 'margin-top:10px;font-size:12px', text: 'Clique num código para ver onde ele é usado; clique no nome de um ativo (:=) para abrir a ficha dele. A edição é feita na tabela (botão "Ir para a tabela").' }));
        const primeiro = f.codigos[0];
        if (primeiro) mostrar(primeiro); else lateralVerificacoes(f, lado);
    }
    function desenharLateralCodigo(m, f, cod, lado) {
        lado.innerHTML = '';
        const u = D.usoDoCodigo(m, cod);
        lado.appendChild(el('h4', { text: 'Código selecionado' }));
        lado.appendChild(el('div', { class: 'bf-mono', style: 'font-size:15px;font-weight:600', text: cod }));
        lado.appendChild(el('div', { class: 'bf-mut', style: 'margin-bottom:8px', text: u ? u.descricoes[0] || '' : '' }));
        if (u) {
            const ativos = Array.from(new Set(u.usos.map(x => x.ativo)));
            lado.appendChild(el('div', { class: 'bf-kv' }, [el('span', { text: 'Usado em' }), el('b', { text: ativos.length + (ativos.length === 1 ? ' ativo' : ' ativos') })]));
            u.usos.slice(0, 12).forEach(x => lado.appendChild(el('div', { class: 'bf-kv' }, [el('span', { class: x.ativo === f.nome ? '' : 'bf-link', text: x.ativo, onclick: x.ativo === f.nome ? null : () => abrirFicha(x.ativo) }), el('span', { text: x.fator_i })])));
            if (u.descricoes.length > 1) lado.appendChild(el('div', { class: 'bf-av', style: 'margin-top:6px', text: `⚠ ${u.descricoes.length} descrições diferentes na base` }));
        }
        lateralVerificacoes(f, lado);
    }
    function lateralVerificacoes(f, lado) {
        lado.appendChild(el('h4', { style: 'margin-top:14px', text: 'Verificações' }));
        f.verificacoes.forEach(v => lado.appendChild(el('div', { class: v.nivel === 'erro' ? 'bf-er' : v.nivel === 'aviso' ? 'bf-av' : 'bf-ok', text: (v.nivel === 'erro' ? '✖ ' : v.nivel === 'aviso' ? '⚠ ' : '✔ ') + v.texto })));
        if (f.chamadoPor.length) {
            lado.appendChild(el('h4', { style: 'margin-top:14px', text: 'Quem chama este ativo' }));
            f.chamadoPor.slice(0, 8).forEach(c => lado.appendChild(el('div', {}, [el('span', { class: 'bf-link', text: c.ativo, onclick: () => abrirFicha(c.ativo) }), ' ×' + c.qtd])));
        }
    }
    function abaParametros(f, main) {
        if (!f.parametros.length) main.appendChild(el('div', { class: 'bf-mut', text: 'Este ativo não declara parâmetros.' }));
        else {
            const tb = el('table', { class: 'bf-tb', style: 'margin-bottom:14px' }, [el('tr', {}, ['Parâmetro', 'Padrão', 'Valores / máscaras usados nas linhas'].map(h => el('th', { text: h })))]);
            f.parametros.forEach(p => tb.appendChild(el('tr', {}, [el('td', { text: p.nome.toLowerCase() }), el('td', { class: 'bf-mono', text: p.padrao || '—' }), el('td', {}, p.valores.map(v => chip(v, 'k')))])));
            main.appendChild(tb);
        }
        const vs = f.variantes;
        main.appendChild(el('h4', { class: 'bf-mut', style: 'font-size:11px;text-transform:uppercase', text: `Variantes (${vs.length})` }));
        const tb2 = el('table', { class: 'bf-tb' }, [el('tr', {}, ['Condição', 'O que traz'].map(h => el('th', { text: h })))]);
        const por = new Map();
        vs.forEach(l => { const k = l.condicao || '—'; if (!por.has(k)) por.set(k, []); por.get(k).push(l); });
        por.forEach((ls, k) => tb2.appendChild(el('tr', {}, [el('td', { class: 'bf-mono', text: k }), el('td', { text: ls.map(l => l.itens ? 'chama ' + l.itens.map(i => i.qtd + '× ' + D.nomeDoItem(i.ativo)).join(' + ') : `${l.codigo} (${l.fator_i})`).join(' · ') })])));
        main.appendChild(tb2);
    }
    function abaRelacoes(f, main) {
        const bloco = (tit, ls, mk) => {
            main.appendChild(el('h4', { class: 'bf-mut', style: 'font-size:11px;text-transform:uppercase;margin:12px 0 6px', text: tit }));
            if (!ls.length) { main.appendChild(el('div', { class: 'bf-mut', text: '— nenhum —' })); return; }
            ls.forEach(x => main.appendChild(mk(x)));
        };
        bloco(`Este ativo chama (${f.chama.length})`, f.chama, c => el('div', { style: 'padding:3px 0' }, [c.qtd + ' × ', c.existe ? el('span', { class: 'bf-link', text: c.nome, onclick: () => abrirFicha(c.nome) }) : el('span', { class: 'bf-er', text: c.nome + ' (não existe)' }), el('span', { class: 'bf-mut bf-mono', text: '   ' + c.texto + (c.condicao ? '  quando ' + c.condicao : '') })]));
        bloco(`É chamado por (${f.chamadoPor.length})`, f.chamadoPor, c => el('div', { style: 'padding:3px 0' }, [el('span', { class: 'bf-link', text: c.ativo, onclick: () => abrirFicha(c.ativo) }), el('span', { class: 'bf-mut bf-mono', text: `   ${c.qtd} × ${c.texto}` + (c.condicao ? '  quando ' + c.condicao : '') })]));
    }
    function abaSimulador(f, main) {
        const padrao = f.parametros.length ? f.nome + '(' + f.parametros.map(p => p.padrao || '').join(',') + ')' : f.nome;
        main.appendChild(el('div', { class: 'bf-mut', style: 'margin-bottom:6px', text: 'Digite o ativo como no Resumo (ex.: ' + padrao + ') e veja os códigos e quantidades. O cálculo usa a base já SALVA.' }));
        const inp = el('input', { class: 'bf-in bf-mono', value: padrao }), qtd = el('input', { class: 'bf-in', value: '1', style: 'width:70px' });
        const saida = el('div', { style: 'margin-top:10px' });
        main.appendChild(el('div', { style: 'display:flex;gap:8px' }, [qtd, inp, el('button', { class: 'bf-b', text: 'Calcular', onclick: calcular })]));
        main.appendChild(saida);
        let t = null;
        async function calcular() {
            saida.textContent = 'Calculando…';
            try {
                const q = parseFloat(String(qtd.value).replace(',', '.')) || 1;
                const r = await (await raiz.fetch('/api/orcamento/calcular', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ cabos: [], outros: [{ ativo: (q === 1 ? '' : q + '-') + inp.value.trim(), operacao: 'I', entidade: '0' }], projeto: cfg.projeto() || '' }) })).json();
                saida.innerHTML = '';
                const tb = el('table', { class: 'bf-tb' }, [el('tr', {}, ['MDO', 'Código', 'Descrição', 'Quantidade', 'Filtro'].map(h => el('th', { text: h })))]);
                (r.resultado || []).forEach(x => tb.appendChild(el('tr', {}, [el('td', { text: U(x.mdo).indexOf('MAO') === 0 ? 'N' : 'M' }), el('td', { class: 'bf-mono', text: x.codigo }), el('td', { text: x.desc_codigo }), el('td', { class: 'n', text: String(x.total).replace('.', ',') }), el('td', { text: x.filtro })])));
                saida.appendChild((r.resultado || []).length ? tb : el('div', { class: 'bf-mut', text: 'Nenhum código gerado.' }));
                (r.nao_encontrados || []).forEach(n => saida.appendChild(el('div', { class: 'bf-er', text: '✖ ' + n })));
                (r.avisos || []).forEach(n => saida.appendChild(el('div', { class: 'bf-av', text: '⚠ ' + n })));
            } catch (e) { saida.innerHTML = ''; saida.appendChild(el('div', { class: 'bf-er', text: 'Não foi possível calcular: ' + e.message })); }
        }
        [inp, qtd].forEach(i => i.addEventListener('input', () => { clearTimeout(t); t = setTimeout(calcular, 450); }));
        calcular();
    }

    async function desfazerDaFicha(nome) {
        if (!raiz.confirm(`Desfazer a última publicação de "${nome}"? O ativo volta ao que era antes (e dá para refazer pelo mesmo botão).`)) return;
        const res = await raiz.fetch('/api/admin/ativo/desfazer', { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + (raiz.localStorage.getItem('auth_token') || '') }, body: JSON.stringify({ projeto: cfg.projeto(), nome }) });
        const j = await res.json().catch(() => ({}));
        if (!res.ok) { raiz.showToast && raiz.showToast(j.detail || 'Falha ao desfazer.'); return; }
        invalidar(); try { await cfg.recarregar(); } catch (e) { /* a tela recarrega depois */ }
        raiz.showToast && raiz.showToast(`"${nome}" voltou ao que era antes da última publicação.`);
        fecharTudo(); abrirFicha(nome);
    }

    /* ───────── tabela ───────── */
    function irParaLinha(nome) {
        if (!cfg.tbody) return;
        const tr = Array.from(cfg.tbody.querySelectorAll('tr')).find(r => { const i = r.querySelector('.cell-input'); return i && D.resolver ? U(raiz.AtivoParametros.analisarCelulaBase(i.value).base) === U(nome) : false; });
        if (!tr) { raiz.showToast && raiz.showToast('Este ativo não está na parte da base mostrada na tela (troque a aba/projeto).'); return; }
        tr.scrollIntoView({ block: 'center' }); tr.classList.add('bf-flash'); setTimeout(() => tr.classList.remove('bf-flash'), 1800);
        const i = tr.querySelector('.cell-input'); if (i) i.focus();
    }
    function iniciar(opcoes) {
        cfg = Object.assign(cfg, opcoes || {});
        if (opcoes && opcoes.barra) montarBarra(opcoes.barra);
        doc.addEventListener('keydown', e => { if ((e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 'k') { e.preventDefault(); abrirPesquisa(); } });
        // Ctrl+clique no ATIVO ou no CÓDIGO da tabela abre a ficha / o uso do código
        if (cfg.tbody) cfg.tbody.addEventListener('mousedown', e => {
            if (!(e.ctrlKey || e.metaKey) || !e.target.classList || !e.target.classList.contains('cell-input')) return;
            const tr = e.target.closest('tr'), col = Array.from(tr.children).indexOf(e.target.closest('td'));
            if (col !== 0 && col !== 6) return;
            e.preventDefault();
            const v = String(e.target.value || '').trim();
            if (!v) return;
            if (col === 0) abrirFicha(raiz.AtivoParametros.analisarCelulaBase(v).base || v); else abrirCodigo(v);
        });
    }
    raiz.BaseFicha = { estilo, modelo, iniciar, abrirPesquisa, abrirFicha, abrirCodigo, invalidar, fechar: fecharTudo };
})(window);
