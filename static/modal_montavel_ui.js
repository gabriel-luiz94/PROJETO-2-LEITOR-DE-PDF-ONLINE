/*
 * modal_montavel_ui.js (TASK-077) — tela dos modais montáveis: botões na barra de modais do Resumo, o modal em si (blocos)
 * e o gerenciador (só admin: lista, editar JSON, exemplos). A lógica pura está em modal_montavel_engine.js (testada em Node).
 * Quem insere nas tabelas é o resumo.js (host.adicionar). Nada aqui executa código vindo da definição: só interpreta blocos.
 */
(function (raiz) {
    'use strict';
    const MM = raiz.ModalMontavel, doc = raiz.document;
    const css = 'padding:5px 10px;font-size:.75rem;';

    const EXEMPLOS = {
        ramais: { id: 'ramais_colados', nome: 'Ramais colados', tipo: 'interpretativo', blocos: [
            { tipo: 'texto', texto: 'Cole as linhas do Excel (Ctrl+V na primeira célula). Cada linha é interpretada pelas regras e vira ativo + quantidade.' },
            { tipo: 'grade', id: 'entrada', linhas_min: 8, colunas: [{ id: 'texto', rotulo: 'TEXTO' }] },
            { tipo: 'interpretar', id: 'saida', fonte: 'entrada', regras: [
                { quando: { coluna: 'texto', regex: 'TROCAR\\s+(\\d+)\\s+RS M AC' }, gera: [{ ativo: 'MAC', qtd: '{1}*20', operacao: 'I' }, { ativo: 'MAC', qtd: '{1}*15', operacao: 'R' }] },
                { quando: { coluna: 'texto', regex: 'CP[- ]?REDE' }, gera: [{ ativo: 'CPREDE', qtd: 1, operacao: 'I' }] }] },
            { tipo: 'destino', id: 'dest', opcoes: ['outros', 'cabos'], padrao: 'outros' },
            { tipo: 'acao_adicionar', fonte: 'saida', destino: 'dest', rotulo: 'Adicionar ativos' }] },
        titulos: { id: 'titulos_contrato', nome: 'Títulos de contrato', tipo: 'gerador_texto', blocos: [
            { tipo: 'campo', id: 'regiao', rotulo: 'Região', entrada: 'select', opcoes: ['NORTE', 'SUL'] },
            { tipo: 'campo', id: 'obs', rotulo: 'Observação (só no SUL)', entrada: 'texto', visivel_se: { campo: 'regiao', op: 'igual', valor: 'SUL' } },
            { tipo: 'grade', id: 'itens', linhas_min: 5, colunas: [{ id: 'codigo', rotulo: 'CÓDIGO' }, { id: 'qtd', rotulo: 'QTD' }] },
            { tipo: 'tabela_interna', id: 'precos', colunas: ['codigo', 'titulo', 'contrato'], linhas: [['MAC', 'Mufla AC', 'C-100'], ['TAM', 'Terminal AM', 'C-200']] },
            { tipo: 'saida_texto', id: 'txt', fonte: 'itens', rotulo: 'Texto para copiar',
              modelo: '{{codigo}} x{{qtd}} - {{buscar(precos, codigo, codigo, titulo)}} ({{buscar(precos, codigo, codigo, contrato)}}) {{regiao}} {{obs}}' }] },
        branco: { id: 'novo_modal', nome: 'Novo modal', tipo: 'montagem', blocos: [{ tipo: 'texto', texto: 'Descreva aqui o que este modal faz.' }] }
    };

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
    const ehAdmin = () => { try { return raiz.localStorage.getItem('is_admin') === 'true'; } catch (e) { return false; } };
    const estiloInput = 'padding:6px 8px;background:#161b22;color:#e6edf3;border:1px solid #30363d;border-radius:6px;box-sizing:border-box;';

    let estado = { modais: [], host: null, barra: null };

    async function carregar() {
        const p = projetoAtual();
        if (!p) return { versao: 1, modais: [] };
        try {
            const r = await fetch('/api/modais-montaveis?projeto_codigo=' + encodeURIComponent(p));
            if (!r.ok) return { versao: 1, modais: [] };
            return (await r.json()).dados || { versao: 1, modais: [] };
        } catch (e) { return { versao: 1, modais: [] }; }
    }

    async function gravar(dados) {
        const r = await fetch('/api/modais-montaveis', { method: 'PUT', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ projeto_codigo: projetoAtual(), dados }) });
        const j = await r.json().catch(() => ({}));
        if (!r.ok) {
            const d = j.detail;
            return { ok: false, erros: (d && d.erros) || [typeof d === 'string' ? d : 'Falha ao salvar (HTTP ' + r.status + ').'] };
        }
        return { ok: true };
    }

    /* ───────────── barra ───────────── */
    async function recarregar() {
        const dados = await carregar();
        estado.modais = dados.modais || [];
        const barra = estado.barra;
        if (!barra) return;
        barra.innerHTML = '';
        estado.modais.forEach(m => barra.appendChild(el('button', { class: 'btn-secondary', style: css + 'border-color:rgba(88,166,255,.4);color:#58a6ff;',
            title: 'Modal montável (' + m.tipo + ')', text: String(m.nome).toUpperCase(), onclick: () => abrir(m) })));
        if (ehAdmin()) barra.appendChild(el('button', { class: 'btn-secondary', style: css + 'font-weight:700;', title: 'Criar/editar modais deste projeto (admin)', text: '＋',
            onclick: abrirGerenciador }));
    }

    /* ───────────── modal ───────────── */
    function janela(titulo, largura) {
        const fundo = el('div', { class: 'mm-janela', style: 'position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:10000;display:flex;align-items:center;justify-content:center;' });
        const caixa = el('div', { style: 'background:#0d1117;color:#e6edf3;border:1px solid #30363d;border-radius:10px;padding:16px 18px;max-width:' + (largura || 900) + 'px;width:95%;max-height:90vh;overflow:auto;font-family:Inter,sans-serif;' });
        const fechar = () => { fundo.remove(); doc.removeEventListener('keydown', tecla, true); };
        const tecla = e => { if (e.key === 'Escape') { e.stopPropagation(); fechar(); } };
        caixa.appendChild(el('div', { style: 'display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;' }, [
            el('h3', { text: titulo, style: 'margin:0;font-size:1rem;' }), el('button', { class: 'btn-secondary', text: '×', style: 'padding:2px 9px;', onclick: fechar })]));
        fundo.appendChild(caixa); doc.body.appendChild(fundo); doc.addEventListener('keydown', tecla, true);
        return { caixa, fechar };
    }

    function abrir(modal) {
        const j = janela(modal.nome, 980);
        const valores = {}, grades = {}, tabelas = {}, destinos = {}, resultados = {};
        const blocosEls = [];
        const G = raiz.GradeTabela;

        (modal.blocos || []).forEach(b => {
            if (b.tipo === 'campo') valores[b.id] = b.entrada === 'select' || b.entrada === 'radio' ? (b.opcoes || [])[0] || '' : (b.padrao == null ? '' : String(b.padrao));
            if (b.tipo === 'grade') grades[b.id] = Array.from({ length: b.linhas_min || 5 }, () => Object.fromEntries(b.colunas.map(c => [c.id, ''])));
            if (b.tipo === 'tabela_interna') tabelas[b.id] = { colunas: b.colunas, linhas: b.linhas };
            if (b.tipo === 'destino') destinos[b.id] = b.padrao || (b.opcoes || ['outros'])[0];
        });

        function recalcular() {
            blocosEls.forEach(x => {
                const b = x.bloco;
                x.raiz.style.display = MM.visivel(b, valores) ? '' : 'none';
                if (b.tipo === 'interpretar') {
                    const itens = MM.interpretar(b, grades[b.fonte] || []);
                    resultados[b.id] = itens;
                    x.atualizar(itens);
                } else if (b.tipo === 'saida_texto') x.atualizar(MM.gerarTexto(b, valores, tabelas, b.fonte ? grades[b.fonte] : null));
            });
        }

        function grade(b) {
            const dados = grades[b.id];
            const tbody = el('tbody');
            const tabela = el('table', { class: 'gm-grade', style: 'width:100%;border-collapse:collapse;' }, [
                el('thead', {}, [el('tr', {}, [el('th', { text: '#', style: 'width:34px;' })].concat(b.colunas.map(c => el('th', { text: c.rotulo, style: 'text-align:left;padding:4px 6px;' }))))]), tbody]);
            function desenhar() {
                tbody.innerHTML = '';
                dados.forEach((linha, i) => {
                    const tr = el('tr', {}, [el('td', { text: String(i + 1), style: 'text-align:center;color:#8b949e;font-size:.72rem;' })]);
                    b.colunas.forEach((c, ci) => {
                        const inp = el('input', { type: 'text', value: linha[c.id], 'data-i': String(i), 'data-c': String(ci), style: estiloInput + 'width:100%;border-radius:0;' });
                        inp.addEventListener('input', () => { linha[c.id] = inp.value; recalcular(); });
                        inp.addEventListener('paste', e => {
                            const t = (e.clipboardData || raiz.clipboardData).getData('text');
                            if (!t || !(/[\n\t]/.test(t.replace(/\n$/, '')))) return;      // 1 valor só: colar normal
                            e.preventDefault();
                            const m = G ? G.parseColagem(t, 1000).linhas : t.split(/\r?\n/).filter(Boolean).map(l => l.split('\t'));
                            m.forEach((cel, k) => {
                                while (i + k >= dados.length) dados.push(Object.fromEntries(b.colunas.map(cc => [cc.id, ''])));
                                cel.forEach((v, kk) => { if (b.colunas[ci + kk]) dados[i + k][b.colunas[ci + kk].id] = String(v).trim(); });
                            });
                            desenhar(); recalcular();
                            toast(m.length + ' linha(s) coladas.');
                        });
                        tr.appendChild(el('td', { style: 'padding:0;' }, [inp]));
                    });
                    tbody.appendChild(tr);
                });
            }
            desenhar();
            const rodape = el('div', { style: 'display:flex;gap:8px;margin-top:6px;' }, [
                el('button', { class: 'btn-secondary', text: '+ Linha', style: css, onclick: () => { dados.push(Object.fromEntries(b.colunas.map(c => [c.id, '']))); desenhar(); } }),
                el('button', { class: 'btn-secondary', text: 'Limpar', style: css, onclick: () => { dados.forEach(l => Object.keys(l).forEach(k => { l[k] = ''; })); desenhar(); recalcular(); } })]);
            return { raiz: el('div', { style: 'margin:10px 0;overflow-x:auto;' }, [tabela, rodape]) };
        }

        function montar(b) {
            switch (b.tipo) {
                case 'texto': return { raiz: el('p', { text: b.texto, style: 'margin:8px 0;font-size:.82rem;color:#8b949e;line-height:1.4;' }) };
                case 'campo': {
                    let ctl;
                    if (b.entrada === 'select') ctl = el('select', { style: estiloInput }, (b.opcoes || []).map(o => el('option', { value: o, text: o })));
                    else if (b.entrada === 'radio') {
                        ctl = el('div', { style: 'display:flex;gap:12px;' }, (b.opcoes || []).map((o, k) => {
                            const r = el('input', { type: 'radio', name: 'mm-' + b.id, value: o }); if (k === 0) r.checked = true;
                            r.addEventListener('change', () => { valores[b.id] = o; recalcular(); });
                            return el('label', { style: 'display:flex;gap:4px;cursor:pointer;font-size:.85rem;' }, [r, o]);
                        }));
                    } else ctl = el('input', { type: b.entrada === 'numero' ? 'number' : 'text', value: valores[b.id], style: estiloInput });
                    if (b.entrada !== 'radio') { ctl.value = valores[b.id]; ctl.addEventListener('input', () => { valores[b.id] = ctl.value; recalcular(); }); }
                    return { raiz: el('div', { style: 'margin:8px 0;' }, [el('label', { text: b.rotulo, style: 'display:block;font-size:.72rem;color:#8b949e;margin-bottom:3px;' }), ctl]) };
                }
                case 'grade': return grade(b);
                case 'tabela_interna': {
                    const t = el('table', { style: 'width:100%;font-size:.78rem;border-collapse:collapse;' }, [
                        el('thead', {}, [el('tr', {}, b.colunas.map(c => el('th', { text: c, style: 'text-align:left;padding:3px 6px;color:#8b949e;' })))]),
                        el('tbody', {}, b.linhas.map(l => el('tr', {}, l.map(v => el('td', { text: v == null ? '' : String(v), style: 'padding:3px 6px;border-top:1px solid #21262d;' })))))]);
                    return { raiz: el('details', { style: 'margin:8px 0;' }, [el('summary', { text: 'Planilha interna: ' + b.id, style: 'cursor:pointer;font-size:.78rem;color:#8b949e;' }), t]) };
                }
                case 'interpretar': {
                    const corpo = el('tbody'), resumo = el('div', { style: 'font-size:.75rem;color:#8b949e;margin-top:4px;' });
                    const t = el('table', { style: 'width:100%;font-size:.82rem;border-collapse:collapse;' }, [
                        el('thead', {}, [el('tr', {}, ['OP', 'ATIVO', 'QTD'].map(c => el('th', { text: c, style: 'text-align:left;padding:3px 6px;color:#8b949e;' })))]), corpo]);
                    return { raiz: el('div', { style: 'margin:10px 0;border:1px solid #21262d;border-radius:6px;padding:8px;' }, [
                        el('div', { text: 'Resultado (ativos e quantidades)', style: 'font-size:.75rem;font-weight:600;color:#c9d1d9;margin-bottom:4px;' }), t, resumo]),
                        atualizar: itens => {
                            corpo.innerHTML = '';
                            itens.forEach(x => corpo.appendChild(el('tr', {}, [x.operacao, x.ativo, String(x.qtd)].map(v => el('td', { text: v, style: 'padding:3px 6px;border-top:1px solid #21262d;' })))));
                            resumo.textContent = itens.length ? itens.length + ' item(ns).' : 'Nada interpretado ainda.';
                        } };
                }
                case 'destino': {
                    const s = el('select', { style: estiloInput }, (b.opcoes || ['outros', 'cabos']).map(o => el('option', { value: o, text: o === 'outros' ? 'Tabela Outros' : 'Tabela Cabos' })));
                    s.value = destinos[b.id]; s.addEventListener('change', () => { destinos[b.id] = s.value; });
                    return { raiz: el('div', { style: 'margin:8px 0;display:flex;gap:8px;align-items:center;' }, [el('span', { text: 'Adicionar em:', style: 'font-size:.8rem;color:#8b949e;' }), s]) };
                }
                case 'acao_adicionar': return { raiz: el('div', { style: 'margin:10px 0;' }, [el('button', { class: 'btn-primary', text: b.rotulo || 'Adicionar', onclick: () => {
                    const itens = resultados[b.fonte] || [];
                    if (!itens.length) { toast('Nada para adicionar.'); return; }
                    const destino = destinos[b.destino] || b.destino;
                    if (!estado.host || !estado.host.adicionar) return;
                    estado.host.adicionar(destino, itens, modal.nome);
                    j.fechar();
                } })]) };
                case 'saida_texto': {
                    const ta = el('textarea', { readonly: 'readonly', rows: '6', style: estiloInput + 'width:100%;font-family:monospace;' });
                    const copiar = el('button', { class: 'btn-primary', text: 'Copiar', style: css, onclick: async () => {
                        try { await raiz.navigator.clipboard.writeText(ta.value); } catch (e) { ta.select(); doc.execCommand('copy'); }
                        toast('Texto copiado.');
                    } });
                    return { raiz: el('div', { style: 'margin:10px 0;' }, [el('div', { text: b.rotulo || 'Texto', style: 'font-size:.75rem;font-weight:600;color:#c9d1d9;margin-bottom:4px;' }), ta,
                        el('div', { style: 'margin-top:6px;' }, [copiar])]), atualizar: t => { ta.value = t; } };
                }
                default: return { raiz: el('div') };
            }
        }

        (modal.blocos || []).forEach(b => {
            const x = montar(b); x.bloco = b; blocosEls.push(x); j.caixa.appendChild(x.raiz);
        });
        recalcular();
    }

    /* ───────────── gerenciador (admin) ───────────── */
    async function abrirGerenciador() {
        if (!ehAdmin()) return;
        if (!projetoAtual()) { toast('Selecione um projeto.'); return; }
        const dados = await carregar();
        const j = janela('Modais montáveis — projeto ' + projetoAtual(), 900);
        const lista = el('div'), editor = el('div');
        j.caixa.appendChild(el('p', { text: 'Cada modal é montado em blocos (aba Visual) ou editado como JSON. Só administradores criam e editam; todos os usuários do projeto usam.', style: 'font-size:.78rem;color:#8b949e;margin:0 0 8px;' }));
        j.caixa.appendChild(lista); j.caixa.appendChild(editor);

        async function persistir(novaLista, msg) {
            const r = await gravar({ versao: 1, modais: novaLista });
            if (!r.ok) { return r.erros; }
            dados.modais = novaLista; toast(msg); await recarregar(); desenharLista(); editor.innerHTML = '';
            return null;
        }
        function desenharLista() {
            lista.innerHTML = '';
            (dados.modais || []).forEach(m => lista.appendChild(el('div', { style: 'display:flex;gap:8px;align-items:center;padding:6px 0;border-bottom:1px solid #21262d;' }, [
                el('div', { style: 'flex:1;' }, [el('b', { text: m.nome }), el('span', { text: ' — ' + m.tipo, style: 'color:#8b949e;font-size:.75rem;' })]),
                el('button', { class: 'btn-secondary', text: 'Editar', style: css, onclick: () => editar(m) }),
                el('button', { class: 'btn-secondary', text: 'Excluir', style: css + 'color:#f85149;', onclick: async () => {
                    if (!raiz.confirm('Excluir o modal "' + m.nome + '"?')) return;
                    const e = await persistir(dados.modais.filter(x => x.id !== m.id), 'Modal excluído.');
                    if (e) toast(e[0]);
                } })])));
            lista.appendChild(el('div', { style: 'display:flex;gap:8px;flex-wrap:wrap;margin-top:10px;' }, [
                el('button', { class: 'btn-primary', text: '+ Novo em branco', style: css, onclick: () => editar(null, EXEMPLOS.branco) }),
                el('button', { class: 'btn-secondary', text: '+ Exemplo: Ramais colados', style: css, onclick: () => editar(null, EXEMPLOS.ramais) }),
                el('button', { class: 'btn-secondary', text: '+ Exemplo: Títulos de contrato', style: css, onclick: () => editar(null, EXEMPLOS.titulos) })]));
        }
        function editar(existente, modelo) {
            const inicial = existente || modelo;
            const ED = raiz.ModalMontavelEditor;
            let ed = null, modo = ED ? 'visual' : 'json';
            const ta = el('textarea', { rows: '18', spellcheck: 'false', style: estiloInput + 'width:100%;font-family:monospace;font-size:.78rem;' });
            ta.value = JSON.stringify(inicial, null, 2);
            const visual = el('div', { class: 'mm-visual' });
            const erros = el('div', { style: 'color:#f85149;font-size:.78rem;margin-top:6px;white-space:pre-wrap;' });
            const abaV = el('button', { type: 'button', class: 'btn-primary', text: 'Visual', style: css, onclick: () => mostrar('visual') });
            const abaJ = el('button', { type: 'button', class: 'btn-secondary', text: 'JSON', style: css, onclick: () => mostrar('json') });
            function lerJson() {
                try { return JSON.parse(ta.value); } catch (e) { erros.textContent = 'JSON inválido: ' + e.message; return null; }
            }
            function mostrar(novoModo) {
                erros.textContent = '';
                if (novoModo === 'json' && modo === 'visual' && ed) ta.value = JSON.stringify(ed.obter(), null, 2);
                if (novoModo === 'visual') {
                    const m = modo === 'json' ? lerJson() : (ed ? ed.obter() : inicial);
                    if (!m || typeof m !== 'object') return;
                    visual.innerHTML = ''; ed = ED.criar(visual, m);
                }
                modo = novoModo;
                visual.style.display = modo === 'visual' ? '' : 'none'; ta.style.display = modo === 'json' ? '' : 'none';
                abaV.className = modo === 'visual' ? 'btn-primary' : 'btn-secondary'; abaJ.className = modo === 'json' ? 'btn-primary' : 'btn-secondary';
            }
            editor.innerHTML = '';
            editor.appendChild(el('div', { style: 'margin-top:12px;display:flex;justify-content:space-between;align-items:center;' }, [
                el('span', { style: 'font-size:.8rem;font-weight:600;', text: existente ? 'Editando: ' + existente.nome : 'Novo modal' }),
                ED ? el('div', { style: 'display:flex;gap:4px;' }, [abaV, abaJ]) : null]));
            editor.appendChild(visual); editor.appendChild(ta); editor.appendChild(erros);
            editor.appendChild(el('div', { style: 'display:flex;gap:8px;margin-top:8px;' }, [
                el('button', { class: 'btn-primary', text: 'Salvar', onclick: async () => {
                    const m = modo === 'visual' && ed ? ed.obter() : lerJson();
                    if (!m) return;
                    let nova = (dados.modais || []).filter(x => !existente || x.id !== existente.id);
                    nova = nova.concat([m]);
                    const e = await persistir(nova, 'Modal salvo.');
                    erros.textContent = e ? e.join('\n') : '';
                } }),
                el('button', { class: 'btn-secondary', text: 'Cancelar', onclick: () => { editor.innerHTML = ''; } })]));
            mostrar(modo);
        }
        desenharLista();
    }

    function iniciar(cfg) {
        estado.host = cfg.host; estado.barra = cfg.barra;
        return recarregar();
    }

    raiz.ModalMontavelUI = { iniciar, recarregar, abrir, EXEMPLOS };
})(window);
