/*
 * base_grade.js (TASK-090) — a tabela da Base de Orçamento (static/orcamento.html) no estilo Excel: seleção de intervalo (clique, arraste, Shift+clique,
 * Shift+setas, Ctrl+A), Ctrl+C em TSV, Delete limpa o intervalo, setas/Enter navegam, largura de coluna ao conteúdo (duplo clique na borda do título),
 * linha nova acima/abaixo da selecionada, duplicar e excluir várias, menu do botão direito. Só DOM: os dados continuam nos <input>/<select> da tabela
 * e só vão para o servidor em "Salvar Alterações" (nada aqui grava).
 */
(function (raiz) {
    'use strict';
    const doc = raiz.document;
    let tbody = null, tabela = null, anc = null, foco = null, arrastando = false, menu = null;

    const linhas = () => Array.from(tbody.rows).filter(tr => tr.style.display !== 'none' && tr.querySelector('.cell-input'));
    const COLS_DADOS = tr => Math.max(0, tr.cells.length - 1);                         // a última coluna é AÇÃO
    const campoDe = (tr, c) => (tr.cells[c] ? tr.cells[c].querySelector('input,select') : null);
    const textoDoCampo = el => !el ? '' : (el.tagName === 'SELECT' ? el.value : el.value);

    function coord(td) { const tr = td.parentElement; return { tr, c: Array.from(tr.cells).indexOf(td) }; }

    function intervalo() {
        if (!anc || !foco) return null;
        const ls = linhas(), i1 = ls.indexOf(anc.tr), i2 = ls.indexOf(foco.tr);
        if (i1 < 0 || i2 < 0) return null;
        return { ls, r1: Math.min(i1, i2), r2: Math.max(i1, i2), c1: Math.min(anc.c, foco.c), c2: Math.max(anc.c, foco.c) };
    }
    const multi = iv => !!iv && (iv.r2 > iv.r1 || iv.c2 > iv.c1);

    function pintar() {
        tbody.querySelectorAll('td.cel-sel').forEach(td => td.classList.remove('cel-sel'));
        tbody.querySelectorAll('tr.linha-sel').forEach(tr => tr.classList.remove('linha-sel'));
        const iv = intervalo();
        if (!iv) return;
        for (let r = iv.r1; r <= iv.r2; r++) {
            const tr = iv.ls[r];
            tr.classList.add('linha-sel');
            if (multi(iv)) for (let c = iv.c1; c <= iv.c2; c++) if (tr.cells[c] && c < COLS_DADOS(tr)) tr.cells[c].classList.add('cel-sel');
        }
    }

    function limpar() { anc = foco = null; pintar(); }

    /** Linha de referência para inserir/duplicar: a última linha do intervalo (ou a do campo em foco). */
    function linhaReferencia() {
        const iv = intervalo();
        if (iv) return iv.ls[iv.r2];
        const a = doc.activeElement, tr = a && a.closest ? a.closest('#' + tbody.id + ' tr, tbody tr') : null;
        return tr && tr.parentElement === tbody ? tr : null;
    }
    function linhasSelecionadas() {
        const iv = intervalo();
        if (iv) return iv.ls.slice(iv.r1, iv.r2 + 1);
        const r = linhaReferencia();
        return r ? [r] : [];
    }

    function inserirLinha(tr, posicao) {
        const ref = linhaReferencia();
        if (!ref || posicao === 'fim') tbody.appendChild(tr);
        else if (posicao === 'acima') tbody.insertBefore(tr, ref);
        else tbody.insertBefore(tr, ref.nextSibling);
        const f = tr.querySelector('input:not([disabled]),select');
        if (f) { f.focus(); if (f.select) f.select(); }
        tr.scrollIntoView({ block: 'nearest' });
    }

    function duplicar() {
        const sel = linhasSelecionadas();
        if (!sel.length) { aviso('Clique numa célula da linha que deseja duplicar.'); return; }
        let depois = sel[sel.length - 1];
        sel.forEach(orig => {
            const copia = orig.cloneNode(true);
            const o = orig.querySelectorAll('input,select'), c = copia.querySelectorAll('input,select');
            o.forEach((el, i) => { if (c[i]) c[i].value = el.value; });
            tbody.insertBefore(copia, depois.nextSibling);
            depois = copia;
        });
        aviso(`${sel.length} linha(s) duplicada(s). Lembre de Salvar Alterações.`);
    }

    function excluirSelecionadas() {
        const sel = linhasSelecionadas();
        if (!sel.length) { aviso('Selecione as linhas (clique ou arraste nas células).'); return; }
        if (sel.length > 1 && !raiz.confirm(`Excluir ${sel.length} linhas? (só vale depois de Salvar Alterações)`)) return;
        sel.forEach(tr => tr.remove());
        limpar();
        aviso(`${sel.length} linha(s) removida(s) da tela. Lembre de Salvar Alterações.`);
    }

    function aviso(m) { if (raiz.showToast) raiz.showToast(m); }

    // ── copiar / limpar ──
    function textoDoIntervalo(iv) {
        const out = [];
        for (let r = iv.r1; r <= iv.r2; r++) {
            const linha = [];
            for (let c = iv.c1; c <= iv.c2; c++) linha.push(String(textoDoCampo(campoDe(iv.ls[r], c))).replace(/[\t\r\n]+/g, ' '));
            out.push(linha.join('\t'));
        }
        return out.join('\n');
    }
    function limparIntervalo(iv) {
        for (let r = iv.r1; r <= iv.r2; r++) for (let c = iv.c1; c <= iv.c2; c++) {
            const el = campoDe(iv.ls[r], c);
            if (!el || el.disabled) continue;
            if (el.tagName === 'SELECT') el.value = 'A'; else el.value = '';
        }
    }

    // ── largura de coluna ao conteúdo ──
    let canvas = null;
    function largura(txt, el) {
        if (!canvas) canvas = doc.createElement('canvas');
        const ctx = canvas.getContext('2d'), cs = raiz.getComputedStyle(el);
        ctx.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
        return ctx.measureText(String(txt)).width;
    }
    function ajustarColunas(soIdx) {
        const ths = Array.from(tabela.querySelectorAll('thead tr:first-child th')), ls = Array.from(tbody.rows).filter(tr => tr.querySelector('.cell-input'));
        ths.forEach((th, i) => {
            if (soIdx !== undefined && soIdx !== null && i !== soIdx) return;
            const nome = (th.childNodes[0] ? th.childNodes[0].textContent : '').trim().toUpperCase();
            let max = largura(nome, th) * 1.08 + 24;
            ls.forEach(tr => {
                const el = campoDe(tr, i);
                if (!el) return;
                const txt = el.tagName === 'SELECT' ? (el.selectedOptions[0] ? el.selectedOptions[0].text : '') : el.value;
                if (txt) max = Math.max(max, largura(txt, el) + 26);
            });
            const w = Math.round(Math.min(Math.max(max, 46), 620));
            th.style.width = th.style.minWidth = th.style.maxWidth = w + 'px';
        });
        tabela.style.tableLayout = 'fixed';
        tabela.style.width = ths.reduce((a, th) => a + (parseFloat(th.style.width) || th.offsetWidth), 0) + 'px';
    }
    function restaurarColunas() {
        tabela.querySelectorAll('thead th').forEach(th => { th.style.width = th.style.minWidth = th.style.maxWidth = ''; });
        tabela.style.tableLayout = ''; tabela.style.width = '';
    }

    // ── menu do botão direito ──
    function fecharMenu() { if (menu) { menu.remove(); menu = null; } }
    function abrirMenu(e, td) {
        e.preventDefault(); fecharMenu();
        const iv = intervalo(), dentro = iv && iv.ls.includes(td.parentElement);
        if (!dentro) { anc = foco = coord(td); pintar(); }
        const n = linhasSelecionadas().length, idxCol = Array.from(td.parentElement.cells).indexOf(td);
        const itens = [['Inserir linha acima', () => raiz.adicionarLinha('acima')], ['Inserir linha abaixo', () => raiz.adicionarLinha('abaixo')],
            [n > 1 ? `Duplicar ${n} linhas` : 'Duplicar linha', duplicar], [n > 1 ? `Excluir ${n} linhas` : 'Excluir linha', excluirSelecionadas], null,
            ['Ajustar esta coluna ao conteúdo', () => ajustarColunas(idxCol)], ['Ajustar todas as colunas', () => ajustarColunas()], ['Restaurar larguras', restaurarColunas]];
        menu = doc.createElement('div');
        menu.style.cssText = 'position:fixed;z-index:9000;background:#161b22;border:1px solid #30363d;border-radius:8px;padding:5px;min-width:230px;box-shadow:0 8px 24px rgba(0,0,0,.5);font-size:.8rem;';
        itens.forEach(i => {
            if (!i) { const s = doc.createElement('div'); s.style.cssText = 'height:1px;background:#30363d;margin:4px 0;'; menu.appendChild(s); return; }
            const d = doc.createElement('div'); d.textContent = i[0]; d.style.cssText = 'padding:6px 10px;border-radius:5px;cursor:pointer;color:#c9d1d9;';
            d.onmouseenter = () => { d.style.background = '#21262d'; }; d.onmouseleave = () => { d.style.background = ''; };
            d.addEventListener('click', () => { fecharMenu(); i[1](); });
            menu.appendChild(d);
        });
        menu.style.left = Math.min(e.clientX, raiz.innerWidth - 250) + 'px'; menu.style.top = Math.min(e.clientY, raiz.innerHeight - 250) + 'px';
        doc.body.appendChild(menu);
    }

    function iniciar(opcoes) {
        tbody = opcoes.tbody; tabela = tbody.closest('table');
        tbody.addEventListener('mousedown', e => {
            if (e.button !== 0) return;
            const td = e.target.closest('td');
            if (!td || td.cellIndex >= COLS_DADOS(td.parentElement) || (e.target.closest && e.target.closest('.btn-del'))) return;
            const c = coord(td);
            if (e.shiftKey && anc) { e.preventDefault(); foco = c; if (doc.activeElement && doc.activeElement.blur) doc.activeElement.blur(); pintar(); return; }
            anc = foco = c; arrastando = true; pintar();
        });
        tbody.addEventListener('mouseover', e => {
            if (!arrastando) return;
            const td = e.target.closest('td');
            if (!td || td.cellIndex >= COLS_DADOS(td.parentElement)) return;
            const c = coord(td);
            if (c.tr === foco.tr && c.c === foco.c) return;
            foco = c;
            if (multi(intervalo())) { doc.body.classList.add('grade-arrastando'); const s = raiz.getSelection && raiz.getSelection(); if (s) s.removeAllRanges(); if (doc.activeElement && doc.activeElement.blur) doc.activeElement.blur(); }
            pintar();
        });
        doc.addEventListener('mouseup', () => { arrastando = false; doc.body.classList.remove('grade-arrastando'); });
        tbody.addEventListener('focusin', e => {            // a seleção acompanha o foco do teclado (Tab, setas, Enter)
            if (arrastando) return;
            const td = e.target.closest ? e.target.closest('td') : null;
            if (!td || td.cellIndex >= COLS_DADOS(td.parentElement)) return;
            const iv = intervalo();
            if (iv && multi(iv) && iv.ls.includes(td.parentElement) && td.cellIndex >= iv.c1 && td.cellIndex <= iv.c2) return;      // foco dentro do intervalo: mantém
            anc = foco = coord(td); pintar();
        });
        tbody.addEventListener('contextmenu', e => { const td = e.target.closest('td'); if (td && td.cellIndex < COLS_DADOS(td.parentElement)) abrirMenu(e, td); });
        doc.addEventListener('click', e => { if (menu && !menu.contains(e.target)) fecharMenu(); });
        doc.addEventListener('keydown', e => {
            if (e.key === 'Escape') { fecharMenu(); return; }
            const a = doc.activeElement, naTabela = a && a.closest && a.closest('tbody') === tbody, corpo = a === doc.body;
            if (!naTabela && !(corpo && intervalo())) return;
            const iv = intervalo(), ehCampo = a && (a.tagName === 'INPUT' || a.tagName === 'SELECT'), mod = e.ctrlKey || e.metaKey;
            if (mod && (e.key === 'a' || e.key === 'A')) {
                if (ehCampo && a.tagName === 'INPUT' && a.value && !(a.selectionStart === 0 && a.selectionEnd === a.value.length)) return;     // 1º Ctrl+A: texto do campo
                const ls = linhas(); if (!ls.length) return;
                e.preventDefault(); anc = { tr: ls[0], c: 0 }; foco = { tr: ls[ls.length - 1], c: COLS_DADOS(ls[0]) - 1 };
                if (a && a.blur) a.blur(); pintar(); return;
            }
            if (mod && (e.key === 'c' || e.key === 'C') && multi(iv)) {
                e.preventDefault();
                const t = textoDoIntervalo(iv);
                if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(() => aviso('Copiado (' + t.split('\n').length + ' linha(s)).'));
                return;
            }
            if ((e.key === 'Delete' || e.key === 'Backspace') && iv && (multi(iv) || corpo)) {
                e.preventDefault(); limparIntervalo(iv); aviso('Células limpas (só vale depois de Salvar Alterações).'); return;
            }
            if (e.shiftKey && !mod && /^Arrow(Up|Down|Left|Right)$/.test(e.key)) {
                if (ehCampo && /Left|Right/.test(e.key)) return;          // Shift+←/→ dentro do campo seleciona texto
                if (a && a.tagName === 'SELECT') return;
                if (!iv) return;
                e.preventDefault();
                const ls = iv.ls, i = ls.indexOf(foco.tr);
                if (e.key === 'ArrowUp') foco = { tr: ls[Math.max(0, i - 1)], c: foco.c };
                if (e.key === 'ArrowDown') foco = { tr: ls[Math.min(ls.length - 1, i + 1)], c: foco.c };
                if (e.key === 'ArrowLeft') foco = { tr: foco.tr, c: Math.max(0, foco.c - 1) };
                if (e.key === 'ArrowRight') foco = { tr: foco.tr, c: Math.min(COLS_DADOS(foco.tr) - 1, foco.c + 1) };
                if (a && a.blur) a.blur(); pintar(); return;
            }
            if (naTabela && ehCampo && a.tagName === 'INPUT' && !mod && !e.altKey && (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Enter')) {
                const ls = linhas(), tr = a.closest('tr'), c = a.closest('td').cellIndex, i = ls.indexOf(tr);
                const d = (e.key === 'ArrowUp' || (e.key === 'Enter' && e.shiftKey)) ? -1 : 1, alvo = ls[i + d];
                e.preventDefault();
                if (alvo) { const f = campoDe(alvo, c); if (f && !f.disabled) { f.focus(); if (f.select) f.select(); } }
            }
        });
        // duplo clique na borda direita do título = ajusta SÓ aquela coluna (como no Excel)
        tabela.querySelectorAll('thead tr:first-child th').forEach(th => {
            th.addEventListener('dblclick', e => { if (e.clientX >= th.getBoundingClientRect().right - 9) ajustarColunas(Array.from(th.parentElement.children).indexOf(th)); });
        });
    }

    raiz.BaseGrade = { iniciar, inserirLinha, linhaReferencia, duplicar, excluirSelecionadas, ajustarColunas, restaurarColunas, limpar,
        pontoDeColagem() { const iv = intervalo(); return iv ? { tr: iv.ls[iv.r1], c: iv.c1 } : null; } };
})(typeof window !== 'undefined' ? window : globalThis);
