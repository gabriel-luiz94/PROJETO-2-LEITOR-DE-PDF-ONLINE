/*
 * barra_modais.js (TASK-117) — configuração da BARRA DE MODAIS por projeto: ocultar/mostrar, ordem e (montáveis) excluir/copiar.
 * Aplica ocultos/ordem aos botões da barra (`[data-modal-id]` dentro de #barra-modais-itens) e monta a seção "Modais do projeto"
 * do painel Configurações. Salvo no servidor por projeto (/api/barra-modais; só admin grava). As funções puras (ordemEfetiva)
 * não tocam no DOM e são testadas em Node: tests/test_barra_modais_117.py.
 */
(function (raiz) {
    'use strict';
    /** Modais do sistema, na ordem padrão da barra. */
    const SISTEMA = [
        { id: 'resumo-rede', nome: 'Resumo da rede', icone: '📊' },
        { id: 'ajustar', nome: 'Ajustar', icone: '🛠' },
        { id: 'validar', nome: 'Validar (e opções)', icone: '✅' },
        { id: 'contexto', nome: 'Contexto', icone: '🧭' },
        { id: 'regras', nome: 'Regras de validação', icone: '📋' },
        { id: 'postes', nome: 'POSTES E ESTRUTURAS', icone: '🧱' },
        { id: 'cabos', nome: 'CABOS', icone: '〰️' },
        { id: 'conexoes', nome: 'CONEXÕES', icone: '🔌' },
        { id: 'ramais', nome: 'RAMAIS', icone: '📍' },
        { id: 'vincular', nome: 'VINCULAR CABOS', icone: '🔗' }
    ];

    /** Ordem efetiva: a salva (só ids que existem) + os que faltam na ordem padrão. */
    function ordemEfetiva(salva, padrao) {
        const existentes = new Set(padrao);
        const saida = [];
        (Array.isArray(salva) ? salva : []).forEach(id => { if (existentes.has(id) && !saida.includes(id)) saida.push(id); });
        padrao.forEach(id => { if (!saida.includes(id)) saida.push(id); });
        return saida;
    }
    function mover(ordem, id, delta) {
        const i = ordem.indexOf(id), j = i + delta;
        if (i < 0 || j < 0 || j >= ordem.length) return ordem.slice();
        const n = ordem.slice(); n.splice(i, 1); n.splice(j, 0, id);
        return n;
    }
    const api = { SISTEMA, ordemEfetiva, mover };
    if (typeof module !== 'undefined' && module.exports) { module.exports = api; return; }

    const doc = raiz.document;
    const css = 'padding:3px 7px;font-size:.74rem;';
    const toast = m => { if (raiz.showToast) raiz.showToast(m); };
    const projetoAtual = () => { try { return raiz.localStorage.getItem('projeto_selecionado_codigo') || ''; } catch (e) { return ''; } };
    const ehAdmin = () => { try { return raiz.localStorage.getItem('is_admin') === 'true'; } catch (e) { return false; } };
    let cfg = { ocultos: [], ordem: [] };
    let montaveis = [];

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

    const idsPadrao = () => SISTEMA.map(s => s.id).concat(montaveis.map(m => 'mm:' + m.id));
    const ordem = () => ordemEfetiva(cfg.ordem, idsPadrao());

    /* ───────────── aplicar na barra ───────────── */
    function aplicar() {
        const itens = doc.getElementById('barra-modais-itens');
        if (!itens) return;
        const o = ordem(), oculto = new Set(cfg.ocultos);
        const posicao = id => (o.indexOf(id) + 1) * 10;
        itens.querySelectorAll('[data-modal-id]').forEach(e => {
            const id = e.getAttribute('data-modal-id');
            e.style.order = String(posicao(id));
            e.classList.toggle('modal-oculto', oculto.has(id));
        });
        const sep = itens.querySelector('.tb-sep');                       // o separador acompanha o primeiro modal de cadastro (POSTES)
        if (sep) sep.style.order = String(posicao('postes') - 1);
        const mais = itens.querySelector('[data-modal-fixo]');            // o ＋ do gerenciador (admin) fica sempre por último
        if (mais) mais.style.order = '100000';
    }

    /* ───────────── servidor ───────────── */
    async function carregar() {
        const p = projetoAtual();
        cfg = { ocultos: [], ordem: [] };
        if (!p) return;
        try {
            const r = await fetch('/api/barra-modais?projeto_codigo=' + encodeURIComponent(p));
            if (r.ok) { const d = (await r.json()).dados || {}; cfg = { ocultos: d.ocultos || [], ordem: d.ordem || [] }; }
        } catch (e) { /* sem rede: barra completa */ }
    }
    async function gravar(nova) {
        const r = await fetch('/api/barra-modais', { method: 'PUT', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ projeto_codigo: projetoAtual(), dados: { versao: 1, ocultos: nova.ocultos, ordem: nova.ordem } }) });
        if (!r.ok) {
            const j = await r.json().catch(() => ({}));
            const d = j.detail;
            toast((d && d.erros && d.erros[0]) || (typeof d === 'string' ? d : 'Não foi possível salvar a barra de modais (HTTP ' + r.status + ').'));
            return false;
        }
        cfg = nova;
        aplicar(); desenhar();
        return true;
    }

    /* ───────────── seção do painel Configurações ───────────── */
    function desenhar() {
        const corpo = doc.getElementById('cfg-modais-corpo');
        if (!corpo) return;
        const admin = ehAdmin();
        corpo.innerHTML = '';
        const projeto = projetoAtual();
        corpo.appendChild(el('div', { style: 'font-size:.72rem;color:#8b949e;margin-bottom:6px;', text: projeto ? 'Projeto: ' + projeto + (admin ? '' : ' (só o administrador altera)') : 'Selecione um projeto.' }));
        if (!projeto) return;
        const o = ordem(), oculto = new Set(cfg.ocultos);
        const sistema = new Map(SISTEMA.map(s => [s.id, s]));
        const mont = new Map(montaveis.map(m => ['mm:' + m.id, m]));
        const lista = el('div', { class: 'cfgm-lista' });
        o.forEach(id => {
            const s = sistema.get(id), m = mont.get(id);
            if (!s && !m) return;
            const off = oculto.has(id);
            const bt = (txt, titulo, fn, extra) => el('button', { type: 'button', class: 'btn-secondary', style: css + (extra || ''), title: titulo, text: txt, onclick: fn, ...(admin ? {} : { disabled: 'disabled' }) });
            const linha = el('div', { class: 'cfgm-linha' + (off ? ' off' : ''), 'data-id': id }, [
                el('span', { class: 'cfgm-nome', title: s ? 'Modal do sistema' : 'Modal montável (' + m.tipo + ')', text: (s ? s.icone : '🧮') + ' ' + (s ? s.nome : m.nome) }),
                el('span', { class: 'cfgm-tag ' + (s ? 't-sis' : 't-mm'), text: s ? 'sistema' : 'montável' }),
                bt(off ? '🚫' : '👁', off ? 'Oculto na barra — clique para mostrar' : 'Visível na barra — clique para ocultar', () => gravar({ ocultos: off ? cfg.ocultos.filter(x => x !== id) : cfg.ocultos.concat([id]), ordem: o })),
                bt('▲', 'Subir', () => gravar({ ocultos: cfg.ocultos, ordem: mover(o, id, -1) })),
                bt('▼', 'Descer', () => gravar({ ocultos: cfg.ocultos, ordem: mover(o, id, 1) })),
                m ? bt('✏️', 'Editar este modal montável', () => raiz.ModalMontavelUI.abrirGerenciador()) : null,
                m ? bt('🗑', 'Excluir este modal montável do projeto', async () => {
                    if (!raiz.confirm('Excluir o modal "' + m.nome + '" deste projeto?\nA versão anterior continua no histórico do gerenciador (admin).')) return;
                    const e = await raiz.ModalMontavelUI.excluir(m.id);
                    if (e) toast(e[0]); else { toast('Modal excluído.'); await gravar({ ocultos: cfg.ocultos.filter(x => x !== id), ordem: o.filter(x => x !== id) }); }
                }, 'color:#f85149;') : null
            ]);
            lista.appendChild(linha);
        });
        corpo.appendChild(lista);
        if (admin) {
            corpo.appendChild(el('div', { class: 'painel-linha', style: 'flex-wrap:wrap;margin-top:8px;' }, [
                el('button', { type: 'button', class: 'btn-primary', style: css, text: '＋ Novo modal montável', onclick: () => raiz.ModalMontavelUI.abrirGerenciador() }),
                el('button', { type: 'button', class: 'btn-secondary', style: css, text: '⧉ Copiar de outro projeto…', onclick: dialogoCopiar }),
                el('button', { type: 'button', class: 'btn-secondary', style: css, text: '↺ Restaurar barra padrão', onclick: () => {
                    if (raiz.confirm('Mostrar todos os modais e voltar à ordem padrão neste projeto?')) gravar({ ocultos: [], ordem: [] });
                } })]));
        }
        corpo.appendChild(el('div', { style: 'font-size:.68rem;color:#6e7681;margin-top:6px;', text: '👁 oculta/mostra o botão só neste projeto (nada é apagado). 🗑 só nos montáveis.' }));
    }

    function dialogoCopiar() {
        const sel = doc.getElementById('select-projeto');
        const projetos = sel ? Array.from(sel.options).map(o => ({ codigo: o.dataset.codigo || o.value, nome: (o.textContent || o.value || '').trim() })).filter(p => p.codigo && p.codigo !== projetoAtual()) : [];
        if (!projetos.length) { toast('Não há outro projeto para copiar.'); return; }
        const fundo = el('div', { style: 'position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:10000;display:flex;align-items:center;justify-content:center;' });
        const caixa = el('div', { style: 'background:#0d1117;color:#e6edf3;border:1px solid #30363d;border-radius:10px;padding:16px 18px;max-width:460px;width:95%;font-family:Inter,sans-serif;' });
        const fechar = () => fundo.remove();
        const origem = el('select', { style: 'background:#161b22;color:#e6edf3;border:1px solid #30363d;border-radius:6px;padding:5px 8px;' }, projetos.map(p => el('option', { value: p.codigo, text: p.nome })));
        const area = el('div', { style: 'margin:10px 0;max-height:240px;overflow:auto;font-size:.82rem;' });
        let achados = [];
        async function listar() {
            area.textContent = 'Carregando…'; achados = [];
            try {
                const r = await fetch('/api/modais-montaveis?projeto_codigo=' + encodeURIComponent(origem.value));
                achados = r.ok ? (((await r.json()).dados || {}).modais || []) : [];
            } catch (e) { achados = []; }
            area.innerHTML = '';
            if (!achados.length) { area.textContent = 'Esse projeto não tem modais montáveis.'; return; }
            achados.forEach(m => area.appendChild(el('label', { style: 'display:flex;gap:8px;align-items:center;padding:3px 0;' }, [el('input', { type: 'checkbox', 'data-id': m.id, checked: 'checked' }), el('span', { text: m.nome + ' — ' + m.tipo })])));
        }
        origem.addEventListener('change', listar);
        caixa.appendChild(el('b', { text: 'Copiar modais de outro projeto' }));
        caixa.appendChild(el('div', { style: 'margin-top:8px;font-size:.8rem;' }, ['Origem: ', origem, ' → destino: ' + projetoAtual()]));
        caixa.appendChild(area);
        caixa.appendChild(el('div', { style: 'font-size:.72rem;color:#8b949e;' , text: 'Só os montáveis são copiados (os do sistema existem em todos os projetos). Id repetido no destino ganha sufixo _2. Quem estava oculto na origem fica oculto aqui.' }));
        caixa.appendChild(el('div', { style: 'display:flex;gap:8px;margin-top:12px;justify-content:flex-end;' }, [
            el('button', { type: 'button', class: 'btn-secondary', text: 'Cancelar', onclick: fechar }),
            el('button', { type: 'button', class: 'btn-primary', text: 'Copiar', onclick: async () => {
                const ids = Array.from(area.querySelectorAll('input:checked')).map(i => i.getAttribute('data-id'));
                if (!ids.length) { toast('Marque ao menos um modal.'); return; }
                const r = await fetch('/api/barra-modais/copiar', { method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ origem: origem.value, destino: projetoAtual(), ids, copiar_ocultos: true }) });
                const j = await r.json().catch(() => ({}));
                if (!r.ok) { const d = j.detail; toast((d && d.erros && d.erros[0]) || (typeof d === 'string' ? d : 'Falha ao copiar (HTTP ' + r.status + ').')); return; }
                toast((j.copiados || []).length + ' modal(is) copiado(s).');
                fechar();
                await recarregar();
                if (raiz.ModalMontavelUI) await raiz.ModalMontavelUI.recarregar();
            } })]));
        fundo.appendChild(caixa); doc.body.appendChild(fundo);
        fundo.addEventListener('mousedown', e => { if (e.target === fundo) fechar(); });
        listar();
    }

    async function recarregar() {
        await carregar();
        montaveis = raiz.ModalMontavelUI ? raiz.ModalMontavelUI.lista() : [];
        aplicar(); desenhar();
    }

    function iniciar() {
        if (raiz.ModalMontavelUI) {
            raiz.ModalMontavelUI.aoRecarregar = lista => { montaveis = lista || []; aplicar(); desenhar(); };      // criou/excluiu/trocou de projeto
        }
        return recarregar();
    }

    raiz.BarraModais = Object.assign({ iniciar, recarregar, aplicar }, api);
})(typeof window !== 'undefined' ? window : globalThis);
