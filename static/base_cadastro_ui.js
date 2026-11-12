/*
 * base_cadastro_ui.js (TASK-144) — CADASTRO RÁPIDO de ativos na Base de Orçamento: assistente em 3 passos (identidade e parâmetros → itens em grade estilo Excel → revisar).
 * Códigos já existentes preenchem descrição/MDO/filtro sozinhos; cola linhas do Excel; valida ao vivo (local) e no servidor (a mesma validação da importação) e simula o ativo
 * ANTES de gravar. "Adicionar à tabela" põe as linhas no fim da tabela do projeto (depois: Salvar Alterações / Sincronizar; a publicação só do ativo vem na parte 3).
 * Rascunho guardado automaticamente neste computador (localStorage).
 */
(function (raiz) {
    'use strict';
    const doc = raiz.document, C = raiz.BaseCadastroDados, FD = raiz.BaseFichaDados;
    const CHAVE = 'base_rascunho_ativo';
    let cfg = { projeto: () => '', podeAdicionar: () => true, adicionar: () => {} };
    let botoesRodape = [], r = C.rascunhoVazio(), passo = 1, modelo = null, servidor = null, md = null, corpo = null, sugestao = null;

    const css = `
    .bc-grid{border-collapse:collapse;width:100%}.bc-grid th{background:var(--bg-subtle,#1a202c);color:var(--text-secondary,#919bb0);text-align:left;padding:5px 6px;border:1px solid var(--border-default,#262e3d);font-size:11px;text-transform:uppercase;position:sticky;top:0}
    .bc-grid td{padding:0;border:1px solid var(--border-default,#262e3d)}.bc-grid input,.bc-grid select{background:transparent;border:0;color:var(--text-primary,#e6edf3);width:100%;padding:6px 7px;font:inherit;box-sizing:border-box;outline:none}
    .bc-grid input:focus,.bc-grid select:focus{background:#2a1a0e;box-shadow:inset 0 0 0 2px #d97706}.bc-grid td.st{padding:4px 8px;white-space:nowrap;font-size:12px}.bc-grid tr.err td{background:#2a1010}
    .bc-sug{position:fixed;z-index:9500;background:var(--bg-surface,#12171f);border:1px solid var(--border-default,#262e3d);border-radius:8px;box-shadow:0 10px 30px #000a;overflow:hidden;min-width:360px;font-size:12px}
    .bc-sug div{padding:5px 10px;cursor:pointer}.bc-sug div:hover,.bc-sug div.a{background:var(--bg-hover,#242c3d)}
    .bc-lbl{display:block;font-size:11px;color:var(--text-secondary,#919bb0);margin:0 0 3px}`;

    const el = (tag, props, filhos) => {
        const e = doc.createElement(tag);
        Object.keys(props || {}).forEach(k => { if (k === 'class') e.className = props[k]; else if (k === 'text') e.textContent = props[k]; else if (k === 'style') e.style.cssText = props[k]; else if (k === 'value') e.value = props[k]; else if (k.indexOf('on') === 0) e.addEventListener(k.slice(2), props[k]); else e.setAttribute(k, props[k]); });
        (filhos || []).forEach(f => f != null && f !== false && e.appendChild(typeof f === 'string' ? doc.createTextNode(f) : f));
        return e;
    };
    const chip = (t, c) => el('span', { class: 'bf-chip' + (c ? ' ' + c : ''), text: t });
    const avisar = m => { if (raiz.showToast) raiz.showToast(m); };

    function salvarRascunho() { try { raiz.localStorage.setItem(CHAVE, JSON.stringify({ r, passo })); } catch (e) { /* sem localStorage */ } }
    function lerRascunho() { try { const j = JSON.parse(raiz.localStorage.getItem(CHAVE) || 'null'); return j && j.r && (j.r.nome || (j.r.itens || []).some(i => i.codigo || i.condicao)) ? j : null; } catch (e) { return null; } }
    function limparRascunho() { try { raiz.localStorage.removeItem(CHAVE); } catch (e) { /* ok */ } }

    /* ───────── janela ───────── */
    function fechar() { if (md) { md.parentElement.remove(); md = null; } if (sugestao) { sugestao.remove(); sugestao = null; } doc.removeEventListener('keydown', aoTeclar, true); }
    function aoTeclar(e) { if (e.key === 'Escape') { if (sugestao) { sugestao.remove(); sugestao = null; e.stopPropagation(); return; } e.stopPropagation(); if (confirmarSaida()) fechar(); } }
    function confirmarSaida() { return true; }                                     // o rascunho já está guardado
    async function abrir(opcoes) {
        raiz.BaseFicha.estilo();
        if (!doc.getElementById('bc-css')) doc.head.appendChild(el('style', { id: 'bc-css', text: css }));
        modelo = await raiz.BaseFicha.modelo();
        const guardado = lerRascunho();
        if (guardado && !(opcoes && opcoes.novo)) { r = guardado.r; passo = guardado.passo || 1; avisar('Rascunho recuperado — use "Limpar" para começar um ativo novo.'); }
        else { r = C.rascunhoVazio(); passo = 1; }
        servidor = null;
        fechar();
        const ov = el('div', { class: 'bf-ov', style: 'z-index:9100' });
        md = el('div', { class: 'bf-md', style: 'width:min(1360px,97vw);height:min(840px,94vh);margin-top:24px' });
        ov.appendChild(md); doc.body.appendChild(ov); doc.addEventListener('keydown', aoTeclar, true);
        desenhar();
    }

    function desenhar() {
        md.innerHTML = '';
        const proj = cfg.projeto();
        md.appendChild(el('div', { class: 'bf-hd' }, [el('h2', { text: 'Novo ativo' }), proj ? chip(proj, 'o') : chip('sem projeto escolhido', 'w'), el('span', { class: 'bf-sp' }),
            el('select', { class: 'bf-b', title: 'Modelos', onchange: e => { aplicarModelo(e.target.value); e.target.value = ''; } }, [el('option', { value: '', text: 'Modelo ▾' }), el('option', { value: 'simples', text: 'Ativo simples (só itens)' }), el('option', { value: 'param', text: 'Ativo com parâmetros (cabo)' }), el('option', { value: 'composto', text: 'Ativo que chama outros ativos' })]),
            el('button', { class: 'bf-b', text: 'Limpar', onclick: () => { if (confirm('Descartar o rascunho atual?')) { r = C.rascunhoVazio(); passo = 1; servidor = null; limparRascunho(); desenhar(); } } }),
            el('button', { class: 'bf-b', text: '✕', onclick: fechar })]));
        const tabs = el('div', { class: 'bf-tabs' });
        ['1 · Identidade e parâmetros', '2 · Itens', '3 · Revisar'].forEach((t, i) => tabs.appendChild(el('span', { class: 'bf-tab' + (passo === i + 1 ? ' a' : ''), text: t, onclick: () => irPara(i + 1) })));
        md.appendChild(tabs);
        corpo = el('div', { class: 'bf-body' });
        md.appendChild(corpo);
        botoesRodape = [];
        if (passo === 1) passo1(); else if (passo === 2) passo2(); else passo3();
        const rod = el('div', { class: 'bf-hd', style: 'border-top:1px solid var(--border-default,#262e3d);border-bottom:0;background:var(--bg-subtle,#1a202c)' }, [chip('Rascunho salvo automaticamente'), el('span', { class: 'bf-sp' })]);
        if (passo > 1) rod.appendChild(el('button', { class: 'bf-b', text: '← Voltar', onclick: () => irPara(passo - 1) }));
        if (passo < 3) rod.appendChild(el('button', { class: 'bf-b', style: 'background:#d97706;border-color:#d97706;color:#111;font-weight:600', text: passo === 1 ? 'Próximo: itens →' : 'Revisar →', onclick: () => irPara(passo + 1) }));
        botoesRodape.forEach(b => rod.appendChild(b));
        md.appendChild(rod);
    }
    function irPara(n) { passo = n; salvarRascunho(); if (n === 3) { servidor = null; desenhar(); validarNoServidor(); } else desenhar(); }

    function aplicarModelo(m) {
        if (!m) return;
        const tem = r.nome || r.itens.some(i => i.codigo);
        if (tem && !confirm('Aplicar o modelo substitui o rascunho atual. Continuar?')) return;
        r = C.rascunhoVazio();
        if (m === 'param') { r.parametros = [{ nome: 'cabo', padrao: '' }]; r.itens = [Object.assign(C.itemVazio(), { condicao: '' }), Object.assign(C.itemVazio(), { condicao: 'cabo=%%' })]; }
        else if (m === 'composto') { r.itens = [Object.assign(C.itemVazio(), { codigo: ':= 2*OUTRO_ATIVO' })]; }
        else r.itens = [C.itemVazio(), C.itemVazio(), C.itemVazio()];
        salvarRascunho(); desenhar();
    }

    /* ───────── passo 1 ───────── */
    function campo(rot, valor, aoMudar, extra) {
        const i = el('input', { class: 'bf-in', value: valor || '', placeholder: extra || '' });
        i.addEventListener('input', () => { aoMudar(i.value); salvarRascunho(); lateral(); });
        return el('div', {}, [el('label', { class: 'bc-lbl', text: rot }), i]);
    }
    let lado = null;
    function lateral() {
        if (!lado) return;
        lado.innerHTML = '';
        lado.appendChild(el('h4', { text: 'Como as linhas vão ficar' }));
        const linhas = C.gerarLinhas(r, cfg.projeto());
        lado.appendChild(el('div', { class: 'bf-mono', style: 'font-size:11px;line-height:1.6;color:#c9d1d9;word-break:break-all' }, linhas.length ? linhas.slice(0, 14).map(l => el('div', { text: l.ativo + (l.codigo ? ' ▸ ' + l.codigo + ' ×' + (l.formula_i || l.fator_i) : '') })) : [el('span', { class: 'bf-mut', text: 'Preencha o nome e os itens.' })]));
        if (linhas.length > 14) lado.appendChild(el('div', { class: 'bf-mut', text: `… e mais ${linhas.length - 14}` }));
        const v = C.validarLocal(r, modelo);
        lado.appendChild(el('h4', { style: 'margin-top:14px', text: 'Validação ao vivo' }));
        if (!v.erros.length && !v.avisos.length) lado.appendChild(el('div', { class: 'bf-ok', text: '✔ nada a apontar até aqui' }));
        v.erros.slice(0, 8).forEach(x => lado.appendChild(el('div', { class: 'bf-er', text: '✖ ' + x.texto })));
        v.avisos.slice(0, 6).forEach(x => lado.appendChild(el('div', { class: 'bf-av', text: '⚠ ' + x.texto })));
        marcarErros(v);
    }
    function passo1() {
        const main = el('div', { class: 'bf-main', style: 'padding:14px' });
        lado = el('div', { class: 'bf-side', style: 'width:340px' });
        corpo.append(main, lado);
        main.appendChild(el('div', { style: 'display:grid;grid-template-columns:1.2fr 1fr 1fr;gap:12px' }, [
            campo('Nome do ativo', r.nome, v => { r.nome = v.trim(); }, 'ex.: CHAVE_FACA_OLEO'), campo('Apelidos (separados por vírgula)', r.apelidos, v => { r.apelidos = v; }, 'ex.: SUO, CFO'), campo('Descrição', r.descricao, v => { r.descricao = v; })]));
        main.appendChild(el('h4', { class: 'bf-mut', style: 'margin:18px 0 6px;font-size:11px;text-transform:uppercase', text: 'Parâmetros (opcional) — o ativo muda de itens conforme o valor' }));
        const tb = el('table', { class: 'bc-grid', style: 'width:60%' }, [el('tr', {}, ['Parâmetro', 'Padrão (se o usuário não informar)', ''].map(h => el('th', { text: h })))]);
        (r.parametros || []).forEach((p, i) => {
            const n = el('input', { value: p.nome, placeholder: 'ex.: cabo' }), d = el('input', { value: p.padrao, placeholder: 'ex.: S04' });
            n.addEventListener('input', () => { p.nome = n.value.trim(); salvarRascunho(); lateral(); }); d.addEventListener('input', () => { p.padrao = d.value.trim(); salvarRascunho(); lateral(); });
            tb.appendChild(el('tr', {}, [el('td', {}, [n]), el('td', {}, [d]), el('td', { style: 'width:34px;text-align:center' }, [el('span', { class: 'bf-link', text: '✕', onclick: () => { r.parametros.splice(i, 1); salvarRascunho(); desenhar(); } })])]));
        });
        tb.appendChild(el('tr', {}, [el('td', { colspan: 3, style: 'padding:6px 8px' }, [el('span', { class: 'bf-link', text: '＋ adicionar parâmetro', onclick: () => { r.parametros.push({ nome: '', padrao: '' }); salvarRascunho(); desenhar(); } })])]));
        main.appendChild(tb);
        main.appendChild(el('div', { class: 'bf-mut', style: 'margin-top:10px;font-size:12px', text: 'Sem parâmetros? Deixe vazio: o ativo terá só itens comuns. Com parâmetros, no passo 2 cada linha pode ter uma condição (ex.: cabo=%S336%).' }));
        lateral();
    }

    /* ───────── passo 2 ───────── */
    function statusDe(it) {
        if (!it.codigo && !it.condicao) return '';
        if (C.ehChamada(it)) return { t: '↪ chama outros ativos', c: 'bf-ok' };
        const info = modelo && it.codigo ? FD.infoDoCodigo(modelo, it.codigo) : null;
        if (!it.codigo) return { t: '', c: '' };
        return info ? { t: '✔ existe na base', c: 'bf-ok' } : { t: it.desc_codigo ? '⚠ código novo' : '⚠ código novo — informe a descrição', c: 'bf-av' };
    }
    let tabela = null;
    function marcarErros(v) {
        if (!tabela) return;
        const com = new Set(v.erros.filter(x => x.item != null).map(x => x.item));
        Array.from(tabela.querySelectorAll('tbody tr')).forEach((tr, i) => tr.classList.toggle('err', com.has(i)));
    }
    function sugerirCodigo(inp) {
        if (sugestao) { sugestao.remove(); sugestao = null; }
        const q = inp.value.trim();
        if (q.length < 3 || q.startsWith(':=')) return;
        const achados = FD.pesquisar(modelo, q, 6).filter(x => x.tipo === 'codigo');
        if (!achados.length) return;
        const b = inp.getBoundingClientRect();
        sugestao = el('div', { class: 'bc-sug', style: `left:${b.left}px;top:${b.bottom + 2}px` }, achados.map(a => el('div', { onmousedown: e => { e.preventDefault(); inp.value = a.codigo; inp.dispatchEvent(new Event('change')); sugestao && sugestao.remove(); sugestao = null; } }, [el('b', { class: 'bf-mono', text: a.codigo }), ' ' + a.descricao, el('span', { class: 'bf-mut', text: '  · ' + a.usos + ' uso(s)' })])));
        doc.body.appendChild(sugestao);
    }
    function linhaGrade(it, i) {
        const tr = el('tr', {});
        const inp = (k, ph, mono) => { const x = el('input', { value: it[k], placeholder: ph || '', class: mono ? 'bf-mono' : '' }); x.addEventListener('input', () => { it[k] = x.value; salvarRascunho(); if (k === 'codigo') sugerirCodigo(x); lateral(); statusLinha(); }); return x; };
        const sel = (k, ops) => { const x = el('select', {}, ops.map(([v, t]) => el('option', { value: v, text: t }))); x.value = it[k]; x.addEventListener('change', () => { it[k] = x.value; salvarRascunho(); lateral(); }); return x; };
        const cod = inp('codigo', 'código  (ou  := 2*OUTRO_ATIVO)', true), desc = inp('desc_codigo', 'auto'), filtro = inp('filtro', 'auto'), mdoS = sel('mdo', [['M', 'M'], ['N', 'N']]);
        const st = el('td', { class: 'st' });
        function statusLinha() { const s = statusDe(it); st.className = 'st ' + (s.c || ''); st.textContent = s.t || ''; }
        cod.addEventListener('change', () => {                                              // código conhecido: descrição, MDO e filtro vêm da base
            if (sugestao) { sugestao.remove(); sugestao = null; }
            const info = !C.ehChamada(it) && modelo ? FD.infoDoCodigo(modelo, it.codigo.trim()) : null;
            if (info) {
                if (!it.desc_codigo) { it.desc_codigo = info.descricao; desc.value = info.descricao; }
                if (!it.filtro) { it.filtro = info.filtro; filtro.value = info.filtro; }
                const m = C.mdoDe(info.mdo) === 'MAO-DE-OBRA' ? 'N' : 'M'; if (it.mdo !== m) { it.mdo = m; mdoS.value = m; }
            }
            salvarRascunho(); statusLinha(); lateral();
        });
        cod.addEventListener('blur', () => setTimeout(() => { if (sugestao) { sugestao.remove(); sugestao = null; } }, 150));
        [inp('condicao', 'cabo=%S336%  (vazio = sempre)', true), cod, desc, mdoS, inp('fator_i', '1'), inp('fator_r', '0'), sel('lv', [['A', 'A'], ['LM', 'LM'], ['LV', 'LV']]), filtro].forEach(x => tr.appendChild(el('td', {}, [x])));
        tr.appendChild(st); statusLinha();
        tr.appendChild(el('td', { style: 'width:30px;text-align:center' }, [el('span', { class: 'bf-link', text: '✕', onclick: () => { r.itens.splice(i, 1); salvarRascunho(); desenhar(); } })]));
        tr.addEventListener('keydown', e => {                                               // Enter desce na mesma coluna e cria linha no fim
            if (e.key !== 'Enter' || !e.target.matches('input')) return;
            e.preventDefault();
            const col = Array.from(tr.children).indexOf(e.target.closest('td'));
            if (!tr.nextElementSibling) { r.itens.push(C.itemVazio()); salvarRascunho(); desenhar(); }
            const prox = (tabela.querySelectorAll('tbody tr')[i + 1] || {}).children;
            const alvo = prox && prox[col] && prox[col].querySelector('input,select'); if (alvo) alvo.focus();
        });
        return tr;
    }
    function acrescentar(itens, aposIdx) {
        const vazios = r.itens.filter(x => !x.codigo && !x.condicao);
        if (vazios.length === r.itens.length) r.itens = [];                                 // só havia linhas vazias: o colado as substitui
        r.itens.splice(aposIdx == null || r.itens.length === 0 ? r.itens.length : aposIdx + 1, 0, ...itens);
        itens.forEach(it => { if (!C.ehChamada(it) && modelo) { const info = FD.infoDoCodigo(modelo, it.codigo); if (info) { if (!it.desc_codigo) it.desc_codigo = info.descricao; if (!it.filtro) it.filtro = info.filtro; it.mdo = C.mdoDe(info.mdo) === 'MAO-DE-OBRA' ? 'N' : 'M'; } } });
        if (!r.itens.some(x => !x.codigo && !x.condicao)) r.itens.push(C.itemVazio());
        salvarRascunho(); desenhar();
    }
    function caixaColar(titulo, dica, aoConfirmar) {
        const ov = el('div', { class: 'bf-ov', style: 'z-index:9300;align-items:center' });
        const ta = el('textarea', { class: 'bf-in bf-mono', rows: 12, placeholder: dica, style: 'resize:vertical' });
        const box = el('div', { class: 'bf-md', style: 'width:min(720px,94vw)' }, [el('div', { class: 'bf-hd' }, [el('h2', { text: titulo })]), el('div', { style: 'padding:12px' }, [el('div', { class: 'bf-mut', style: 'margin-bottom:8px;font-size:12px', text: dica }), ta]),
            el('div', { class: 'bf-hd', style: 'border-top:1px solid var(--border-default,#262e3d);border-bottom:0' }, [el('span', { class: 'bf-sp' }), el('button', { class: 'bf-b', text: 'Cancelar', onclick: () => ov.remove() }), el('button', { class: 'bf-b', style: 'background:#d97706;border-color:#d97706;color:#111;font-weight:600', text: 'Adicionar', onclick: () => { ov.remove(); aoConfirmar(ta.value); } })])]);
        ov.appendChild(box); doc.body.appendChild(ov); ta.focus();
    }
    function passo2() {
        const main = el('div', { class: 'bf-main', style: 'padding:14px' });
        lado = el('div', { class: 'bf-side', style: 'width:340px' });
        corpo.append(main, lado);
        const params = (r.parametros || []).map(p => p.nome).filter(Boolean);
        main.appendChild(el('div', { style: 'display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap;align-items:center' }, [
            el('button', { class: 'bf-b', text: 'Colar do Excel', onclick: () => caixaColar('Colar itens do Excel', 'Cole as linhas copiadas do Excel. Formatos aceitos: CÓDIGO ⇥ FATOR I [⇥ FATOR R ⇥ LV]  ou  as 13 colunas da base (ATIVO … ARRED). Uma linha de cabeçalho é ignorada.', t => acrescentar(C.lerColado(t))) }),
            params.length ? el('button', { class: 'bf-b', text: 'Colar combinações', title: 'Tabela por combinação de valores (como o CON)', onclick: () => caixaColar('Colar combinações', `Uma linha por combinação: ${params.slice(0, 3).join(' ⇥ ')} ⇥ CÓDIGO [⇥ FATOR]. Os valores viram máscaras (${params[0]}=%valor%).`, t => acrescentar(C.lerCombinacoes(t, params))) }) : null,
            el('button', { class: 'bf-b', text: '＋ Linha', onclick: () => { r.itens.push(C.itemVazio()); salvarRascunho(); desenhar(); } }),
            el('span', { class: 'bf-mut', style: 'font-size:12px', text: 'Enter desce · Tab avança · código existente preenche descrição, MDO e filtro · ":= 3*TERM(cabo)" chama outro ativo' })]));
        if (!r.itens.length) r.itens.push(C.itemVazio(), C.itemVazio(), C.itemVazio());
        tabela = el('table', { class: 'bc-grid' }, [el('thead', {}, [el('tr', {}, [['Condição', 170], ['Código', 190], ['Descrição (auto)', 0], ['MDO', 66], ['Fator I', 74], ['Fator R', 74], ['LV', 76], ['Filtro (auto)', 120], ['', 210], ['', 30]].map(([h, w]) => el('th', { text: h, style: w ? 'width:' + w + 'px' : '' })))]), el('tbody', {}, r.itens.map(linhaGrade))]);
        main.appendChild(tabela);
        main.addEventListener('paste', e => {                                              // colar várias linhas direto na grade
            const t = (e.clipboardData || raiz.clipboardData).getData('text');
            if (!/[\n\t]/.test(t.trim())) return;
            e.preventDefault();
            const tr = e.target.closest ? e.target.closest('tbody tr') : null, idx = tr ? Array.from(tr.parentElement.children).indexOf(tr) : null;
            acrescentar(C.lerColado(t), idx);
        });
        lateral();
    }

    /* ───────── passo 3 ───────── */
    async function validarNoServidor() {
        const linhas = C.gerarLinhas(r, cfg.projeto());
        const exemplo = (r.nome || 'ATIVO') + ((r.parametros || []).filter(p => p.nome).length ? '(' + r.parametros.filter(p => p.nome).map(p => p.padrao || '').join(',') + ')' : '');
        try {
            servidor = await (await raiz.fetch('/api/orcamento/rascunho', { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + (raiz.localStorage.getItem('auth_token') || '') },
                body: JSON.stringify({ linhas, projeto: cfg.projeto(), simular: [{ ativo: exemplo }] }) })).json();
            if (servidor.detail) servidor = { erros: [String(servidor.detail)], avisos: [], linhas: [], simulacao: null };
        } catch (e) { servidor = { erros: ['Não foi possível validar no servidor: ' + e.message], avisos: [], linhas: [], simulacao: null }; }
        if (passo === 3) desenhar();
    }
    function passo3() {
        const local = C.validarLocal(r, modelo), geradas = (servidor && servidor.linhas && servidor.linhas.length) ? servidor.linhas : C.gerarLinhas(r, cfg.projeto());
        const erros = local.erros.map(x => x.texto).concat(servidor ? servidor.erros : []), avisos = local.avisos.map(x => x.texto).concat(servidor ? servidor.avisos : []);
        const main = el('div', { class: 'bf-main', style: 'padding:14px' }), lat = el('div', { class: 'bf-side', style: 'width:360px' });
        corpo.append(main, lat);
        main.appendChild(el('div', { style: 'display:flex;gap:8px;margin-bottom:10px;align-items:center' }, [chip(`${geradas.length} linha(s) para a base`, 'k'), servidor && servidor.componentes_gerados ? chip(`${servidor.componentes_gerados} COMPONENTE(s) gerado(s)`) : null, el('span', { class: 'bf-sp' }),
            el('button', { class: 'bf-b', text: 'Copiar linhas (Excel)', onclick: () => { try { raiz.navigator.clipboard.writeText(C.paraTsv(geradas)); avisar('Linhas copiadas (13 colunas da base, sem cabeçalho).'); } catch (e) { avisar('Não foi possível copiar.'); } } }),
            el('button', { class: 'bf-b', text: 'Revalidar', onclick: () => { servidor = null; desenhar(); validarNoServidor(); } })]));
        const tb = el('table', { class: 'bf-tb' }, [el('tr', {}, ['Ativo', 'Componente', 'MDO', 'Código', 'Descrição', 'Fator I', 'Fator R', 'LV', 'Filtro'].map(h => el('th', { text: h })))]);
        geradas.forEach(l => tb.appendChild(el('tr', {}, [el('td', { class: 'bf-mono', text: l.ativo }), el('td', { class: 'bf-mono bf-mut', text: l.componente }), el('td', { text: l.mdo === 'MATERIAL' ? 'M' : l.mdo ? 'N' : '' }), el('td', { class: 'bf-mono', text: l.codigo }), el('td', { text: l.desc_codigo }),
            el('td', { class: 'n', text: String(l.formula_i || l.fator_i) }), el('td', { class: 'n', text: String(l.formula_r || l.fator_r) }), el('td', { text: l.codigo || l.ativo.indexOf(':=') === -1 ? l.lv : '' }), el('td', { text: l.filtro })])));
        main.appendChild(tb);
        lat.appendChild(el('h4', { text: 'Validação' }));
        if (!servidor) lat.appendChild(el('div', { class: 'bf-mut', text: 'Validando no servidor…' }));
        if (servidor && !erros.length) lat.appendChild(el('div', { class: 'bf-ok', text: '✔ nenhum erro — a base continua válida com este ativo' }));
        erros.forEach(x => lat.appendChild(el('div', { class: 'bf-er', text: '✖ ' + x })));
        avisos.slice(0, 10).forEach(x => lat.appendChild(el('div', { class: 'bf-av', text: '⚠ ' + x })));
        lat.appendChild(el('h4', { style: 'margin-top:14px', text: 'Simulador (com este rascunho)' }));
        const sim = servidor && servidor.simulacao && servidor.simulacao[0];
        const campoSim = el('input', { class: 'bf-in bf-mono', value: sim ? sim.ativo : '' }), saida = el('div', { style: 'margin-top:8px' });
        lat.append(el('div', { style: 'display:flex;gap:6px' }, [campoSim, el('button', { class: 'bf-b', text: 'Simular', onclick: async () => {
            saida.textContent = 'Calculando…';
            const j = await (await raiz.fetch('/api/orcamento/rascunho', { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + (raiz.localStorage.getItem('auth_token') || '') }, body: JSON.stringify({ linhas: C.gerarLinhas(r, cfg.projeto()), projeto: cfg.projeto(), simular: [{ ativo: campoSim.value }] }) })).json();
            mostrarSim(saida, j.simulacao && j.simulacao[0]);
        } })]), saida);
        mostrarSim(saida, sim);
        const pode = cfg.podeAdicionar(), bloqueado = erros.length > 0 || !servidor || !pode;
        botoesRodape.push(el('button', { class: 'bf-b', style: bloqueado ? 'opacity:.5;cursor:not-allowed' : 'background:#238636;border-color:#238636;color:#fff;font-weight:600', text: 'Adicionar à tabela', title: !pode ? 'Escolha um projeto e a aba "Projeto selecionado"' : erros.length ? 'Corrija os erros antes' : 'Põe as linhas no fim da tabela', onclick: () => {
            if (bloqueado) { avisar(!pode ? 'Escolha um projeto na aba "Projeto selecionado" para adicionar.' : 'Corrija os erros antes de adicionar.'); return; }
            cfg.adicionar(geradas); limparRascunho(); fechar();
        } }));
    }
    function mostrarSim(saida, s) {
        saida.innerHTML = '';
        if (!s) return;
        if ((s.resultado || []).length) {
            const tb = el('table', { class: 'bf-tb' }, [el('tr', {}, ['Código', 'Quantidade'].map(h => el('th', { text: h })))]);
            s.resultado.forEach(x => tb.appendChild(el('tr', {}, [el('td', { class: 'bf-mono', text: x.codigo }), el('td', { class: 'n', text: String(x.total).replace('.', ',') })])));
            saida.appendChild(tb);
        } else saida.appendChild(el('div', { class: 'bf-mut', text: 'Nenhum código gerado.' }));
        (s.nao_encontrados || []).forEach(n => saida.appendChild(el('div', { class: 'bf-er', text: '✖ ' + n })));
        (s.avisos || []).forEach(n => saida.appendChild(el('div', { class: 'bf-av', text: '⚠ ' + n })));
    }

    function iniciar(opcoes) { cfg = Object.assign(cfg, opcoes || {}); }
    raiz.BaseCadastro = { iniciar, abrir, fechar };
})(window);
