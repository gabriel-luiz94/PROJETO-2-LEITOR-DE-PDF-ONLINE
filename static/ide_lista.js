/*
 * ide_lista.js — lista de sugestões no estilo do VS Code (a "suggest widget"): painel escuro colado no campo, item selecionado em azul, valor à esquerda e detalhe cinza à direita.
 * Quem usa: o campo Operação do Resumo (todas as opções aparecem ao entrar no campo; ao digitar, filtra) e a caixa de parâmetros (ativo_dica.js usa as mesmas cores).
 *   const l = IdeLista.criar({ aoEscolher: item => … });
 *   l.abrir(campo, [{ valor: 'I', detalhe: 'instalar' }, …], indiceSelecionado);   l.mover(+1);   l.escolherAtual();   l.fechar();   l.aberta()
 * Só DOM; sem dependências.
 */
(function (raiz) {
    'use strict';
    const doc = raiz.document;
    const CSS = '.ide-lista{position:fixed;z-index:9500;min-width:210px;max-width:420px;max-height:264px;overflow:auto;background:#252526;color:#cccccc;border:1px solid #454545;border-radius:4px;'
        + 'box-shadow:0 6px 18px rgba(0,0,0,.55);font:12px/1.5 "Cascadia Code",Consolas,Menlo,monospace;padding:2px 0}'
        + '.ide-lista .il-item{display:flex;align-items:center;gap:14px;padding:1px 8px;cursor:pointer;white-space:nowrap}'
        + '.ide-lista .il-item:hover{background:#2a2d2e}.ide-lista .il-item.sel{background:#094771;color:#fff}'
        + '.ide-lista .il-valor{font-weight:600}.ide-lista .il-det{margin-left:auto;color:#8b8b8b;font-size:11px;font-family:Inter,system-ui,sans-serif}.ide-lista .il-item.sel .il-det{color:#c5d9ea}'
        + '.ide-lista .il-vazio{padding:3px 10px;color:#8b8b8b;font-family:Inter,system-ui,sans-serif}';

    function garantirCss() {
        if (doc.getElementById('ide-lista-css')) return;
        const st = doc.createElement('style'); st.id = 'ide-lista-css'; st.textContent = CSS; doc.head.appendChild(st);
    }

    function criar(opcoes) {
        const aoEscolher = (opcoes && opcoes.aoEscolher) || (() => {});
        let el = null, itens = [], sel = -1, campo = null;

        function desenhar() {
            el.innerHTML = '';
            if (!itens.length) { const v = doc.createElement('div'); v.className = 'il-vazio'; v.textContent = 'Nenhuma opção'; el.appendChild(v); return; }
            itens.forEach((it, i) => {
                const d = doc.createElement('div'); d.className = 'il-item' + (i === sel ? ' sel' : ''); d.setAttribute('role', 'option');
                const v = doc.createElement('span'); v.className = 'il-valor'; v.textContent = it.valor; d.appendChild(v);
                if (it.detalhe) { const t = doc.createElement('span'); t.className = 'il-det'; t.textContent = it.detalhe; d.appendChild(t); }
                d.addEventListener('mousedown', e => { e.preventDefault(); sel = i; escolherAtual(); });             // mousedown (não click): o campo não perde o foco
                el.appendChild(d);
            });
            const ativo = el.querySelector('.il-item.sel'); if (ativo && ativo.scrollIntoView) ativo.scrollIntoView({ block: 'nearest' });
        }
        function posicionar() {
            if (!campo || !el) return;
            const r = campo.getBoundingClientRect();
            el.style.minWidth = Math.max(210, r.width) + 'px';
            const abaixo = r.bottom + 2, h = el.offsetHeight || 200;
            el.style.left = Math.max(4, Math.min(r.left, raiz.innerWidth - el.offsetWidth - 4)) + 'px';
            el.style.top = (abaixo + h > raiz.innerHeight - 4 && r.top - h - 2 > 4 ? r.top - h - 2 : abaixo) + 'px';
        }
        function abrir(alvo, lista, indice) {
            garantirCss();
            campo = alvo; itens = lista || []; sel = indice == null ? -1 : indice;
            if (!el) { el = doc.createElement('div'); el.className = 'ide-lista'; el.setAttribute('role', 'listbox'); doc.body.appendChild(el); }
            desenhar(); posicionar();
        }
        function fechar() { if (el) { el.remove(); el = null; } itens = []; sel = -1; campo = null; }
        function mover(passo) {
            if (!el || !itens.length) return;
            sel = sel < 0 ? (passo > 0 ? 0 : itens.length - 1) : (sel + passo + itens.length) % itens.length;
            desenhar();
        }
        function escolherAtual() {
            if (!el || sel < 0 || !itens[sel]) return false;
            const it = itens[sel], alvo = campo;
            fechar();
            aoEscolher(it, alvo);
            return true;
        }
        return { abrir, fechar, mover, escolherAtual, aberta: () => !!el, doCampo: () => campo, indice: () => sel, itens: () => itens.slice() };
    }

    raiz.IdeLista = { criar };
    if (typeof module !== 'undefined' && module.exports) module.exports = raiz.IdeLista;
})(typeof window !== 'undefined' ? window : globalThis);
