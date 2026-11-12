/*
 * resumo_layout.js (TASK-111) — comportamento do layout da tela Resumo do Orçamento (só DOM, sem dados das tabelas):
 * painéis laterais (Configurações / Assistente IA) ocultos por padrão, abrir/fechar/fixar; barra de modais ocultável;
 * layout lado a lado × empilhadas; mostrar/ocultar/recolher cada tabela; modal Resumo da rede; chip do projeto.
 * O estado fica em localStorage (por usuário/navegador). Nada aqui toca em tableStates (isso é do resumo.js).
 */
(function () {
    'use strict';
    const CHAVE = 'resumo_layout_v1';
    const PADRAO = {
        config: { aberto: false, fixado: false },
        ia: { aberto: false, fixado: false },
        modaisRecolhida: false,
        layout: 'lado',
        inverter: false,
        cabos: { visivel: true, recolhida: false },
        outros: { visivel: true, recolhida: false }
    };

    function clonar(o) { return JSON.parse(JSON.stringify(o)); }
    function lerEstado() {
        let e = clonar(PADRAO);
        try {
            const salvo = JSON.parse(localStorage.getItem(CHAVE) || 'null');
            if (salvo && typeof salvo === 'object') {
                ['config', 'ia', 'cabos', 'outros'].forEach(k => { if (salvo[k] && typeof salvo[k] === 'object') Object.assign(e[k], salvo[k]); });
                if (typeof salvo.modaisRecolhida === 'boolean') e.modaisRecolhida = salvo.modaisRecolhida;
                if (salvo.layout === 'lado' || salvo.layout === 'empilhadas') e.layout = salvo.layout;
                if (typeof salvo.inverter === 'boolean') e.inverter = salvo.inverter;
            }
        } catch (err) { /* sem armazenamento: usa o padrão */ }
        return e;
    }
    let estado = lerEstado();
    function gravar() { try { localStorage.setItem(CHAVE, JSON.stringify(estado)); } catch (err) { /* ignora */ } }
    const $ = id => document.getElementById(id);

    /* ── Painéis ── */
    const painel = { config: $('painel-config'), ia: $('painel-ia') };
    const botaoPainel = { config: $('btn-toggle-config'), ia: $('btn-toggle-ia') };

    function aplicarPainel(qual) {
        const el = painel[qual], st = estado[qual];
        if (!el) return;
        el.hidden = !st.aberto;
        el.classList.toggle('flutuante', st.aberto && !st.fixado);
        const fix = el.querySelector('[data-acao="fixar"]');
        if (fix) fix.setAttribute('aria-pressed', st.fixado ? 'true' : 'false');
        if (botaoPainel[qual]) botaoPainel[qual].setAttribute('aria-pressed', st.aberto ? 'true' : 'false');
    }
    function definirPainel(qual, aberto) {
        estado[qual].aberto = aberto;
        aplicarPainel(qual);
        gravar();
        window.dispatchEvent(new Event('resize'));          // as linhas vazias das tabelas se recalculam
        if (aberto && qual === 'ia') { const i = $('chat-input'); if (i) setTimeout(() => i.focus(), 30); }
    }
    ['config', 'ia'].forEach(q => {
        if (botaoPainel[q]) botaoPainel[q].addEventListener('click', () => definirPainel(q, !estado[q].aberto));
        const el = painel[q];
        if (!el) return;
        el.querySelector('[data-acao="fechar"]').addEventListener('click', () => definirPainel(q, false));
        el.querySelector('[data-acao="fixar"]').addEventListener('click', () => {
            estado[q].fixado = !estado[q].fixado;
            aplicarPainel(q); gravar();
            window.dispatchEvent(new Event('resize'));
        });
    });
    // painel solto (flutuante): fecha ao clicar fora ou com Esc
    document.addEventListener('mousedown', e => {
        ['config', 'ia'].forEach(q => {
            const st = estado[q], el = painel[q];
            if (!el || !st.aberto || st.fixado) return;
            if (el.contains(e.target) || (botaoPainel[q] && botaoPainel[q].contains(e.target))) return;
            if (e.target.closest && e.target.closest('.modal-overlay:not(.hidden), #ctx-menu-resumo, .ac-dropdown')) return;      // modais abertos a partir do painel não o fecham
            definirPainel(q, false);
        });
    });
    document.addEventListener('keydown', e => {
        if (e.key !== 'Escape') return;
        if (document.querySelector('.modal-overlay:not(.hidden)')) return;
        ['config', 'ia'].forEach(q => { if (estado[q].aberto && !estado[q].fixado) definirPainel(q, false); });
    });

    /* ── Barra de modais ── */
    const barraModais = $('barra-modais'), btnModais = $('btn-toggle-modais');
    function aplicarModais() {
        if (!barraModais) return;
        barraModais.classList.toggle('recolhida', estado.modaisRecolhida);
        btnModais.setAttribute('aria-expanded', estado.modaisRecolhida ? 'false' : 'true');
        btnModais.textContent = estado.modaisRecolhida ? '▸ Modais (clique para mostrar)' : '▾ Modais';
    }
    if (btnModais) btnModais.addEventListener('click', () => { estado.modaisRecolhida = !estado.modaisRecolhida; aplicarModais(); gravar(); window.dispatchEvent(new Event('resize')); });

    /* ── Layout e tabelas ── */
    const linhaTabelas = document.querySelector('.tables-row');
    const secao = { cabos: $('section-cabos'), outros: $('section-outros') };
    const chip = { cabos: $('chip-tabela-cabos'), outros: $('chip-tabela-outros') };

    function aplicarLayout() {
        if (linhaTabelas) linhaTabelas.classList.toggle('empilhadas', estado.layout === 'empilhadas');
        document.querySelectorAll('#seg-layout button[data-layout]').forEach(b => b.classList.toggle('on', b.dataset.layout === estado.layout));
        // TASK-150: "⇄ Inverter" — Demais ativos (postes) primeiro (esquerda / em cima) e Cabos depois (direita / embaixo)
        if (linhaTabelas) linhaTabelas.classList.toggle('invertidas', !!estado.inverter);
        const inv = $('btn-inverter-tabelas');
        if (inv) { inv.classList.toggle('on', !!estado.inverter); inv.setAttribute('aria-pressed', estado.inverter ? 'true' : 'false'); }
    }
    function aplicarTabela(qual) {
        const el = secao[qual], st = estado[qual];
        if (!el) return;
        el.classList.toggle('oculta', !st.visivel);
        el.classList.toggle('recolhida', st.visivel && st.recolhida);
        if (chip[qual]) {
            chip[qual].classList.toggle('on', st.visivel);
            chip[qual].setAttribute('aria-pressed', st.visivel ? 'true' : 'false');
            chip[qual].textContent = (st.visivel ? '👁 ' : '🚫 ') + (qual === 'cabos' ? 'Cabos' : 'Demais ativos');
        }
        const rec = el.querySelector('.btn-recolher');
        if (rec) rec.textContent = st.recolhida ? '⌃ Mostrar' : '⌄ Ocultar';
    }
    ['cabos', 'outros'].forEach(q => {
        const el = secao[q];
        if (!el) return;
        const acoes = el.querySelector('.section-actions');
        if (acoes && !acoes.querySelector('.btn-recolher')) {
            const b = document.createElement('button');
            b.type = 'button'; b.className = 'btn-recolher';
            b.title = 'Recolher/mostrar só o corpo desta tabela (o título fica)';
            b.addEventListener('click', () => { estado[q].recolhida = !estado[q].recolhida; aplicarTabela(q); gravar(); window.dispatchEvent(new Event('resize')); });
            acoes.appendChild(b);
        }
        if (chip[q]) chip[q].addEventListener('click', () => {
            estado[q].visivel = !estado[q].visivel;
            if (estado[q].visivel) estado[q].recolhida = false;
            aplicarTabela(q); gravar(); window.dispatchEvent(new Event('resize'));
        });
    });
    document.querySelectorAll('#seg-layout button[data-layout]').forEach(b => b.addEventListener('click', () => {
        estado.layout = b.dataset.layout; aplicarLayout(); gravar(); window.dispatchEvent(new Event('resize'));
    }));
    if ($('btn-inverter-tabelas')) $('btn-inverter-tabelas').addEventListener('click', () => {
        estado.inverter = !estado.inverter; aplicarLayout(); gravar(); window.dispatchEvent(new Event('resize'));
    });

    /* ── Importar JSON: o resumo.js abre o arquivo e mostra o diálogo Carregar / Adicionar / Subtrair (TASK-118) ── */
    const btnImp = $('btn-importar-json'), inpImp = $('input-importar-json');
    if (btnImp && inpImp) btnImp.addEventListener('click', () => inpImp.click());

    /* ── Modal Resumo da rede ── */
    const modalRede = $('modal-resumo-rede');
    function abrirResumoRede() {
        const corpo = $('modal-resumo-rede-corpo');
        const itens = [['Postes instalando', 'resumo-rede-postes'], ['Rede de média instalando', 'resumo-rede-media'],
                       ['Rede de baixa instalando', 'resumo-rede-baixa'], ['Equipamentos instalando', 'resumo-rede-equipamentos']];
        corpo.innerHTML = '';
        itens.forEach(([rotulo, id]) => {
            const cx = document.createElement('div'); cx.className = 'cx';
            const s = document.createElement('span'); s.textContent = rotulo;
            const b = document.createElement('b'); b.textContent = ($(id) && $(id).textContent) || '0';
            cx.appendChild(s); cx.appendChild(b); corpo.appendChild(cx);
        });
        modalRede.classList.remove('hidden');
    }
    if ($('btn-resumo-rede')) $('btn-resumo-rede').addEventListener('click', abrirResumoRede);
    if ($('btn-fechar-resumo-rede')) $('btn-fechar-resumo-rede').addEventListener('click', () => modalRede.classList.add('hidden'));

    /* ── Chip do projeto (o seletor fica no painel Configurações) ── */
    function atualizarChipProjeto() {
        const sel = $('select-projeto'), c = $('chip-projeto');
        if (!sel || !c) return;
        const o = sel.options[sel.selectedIndex];
        c.textContent = 'Projeto: ' + (o ? (o.textContent || o.value || '—').trim() : '—');
    }
    const selProj = $('select-projeto');
    if (selProj) {
        selProj.addEventListener('change', atualizarChipProjeto);
        new MutationObserver(atualizarChipProjeto).observe(selProj, { childList: true });
        setInterval(atualizarChipProjeto, 2000);          // o valor é definido por código (sem evento) ao carregar os projetos
    }

    /* ── Restaurar padrão ── */
    if ($('btn-restaurar-layout')) $('btn-restaurar-layout').addEventListener('click', () => {
        estado = clonar(PADRAO);
        ['config', 'ia'].forEach(aplicarPainel); aplicarModais(); aplicarLayout(); ['cabos', 'outros'].forEach(aplicarTabela); gravar();
        window.dispatchEvent(new Event('resize'));
    });

    ['config', 'ia'].forEach(aplicarPainel); aplicarModais(); aplicarLayout(); ['cabos', 'outros'].forEach(aplicarTabela);
    atualizarChipProjeto();

    window.ResumoLayout = { estado: () => clonar(estado), definirPainel, abrirResumoRede };
})();
