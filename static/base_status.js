/*
 * base_status.js — barra de status das operações da Base de Orçamento (Salvar, Sincronizar, Validar, Publicar, Desfazer, Recarregar).
 * Uma barra só, fixa no canto: "⏳ Publicando…" → "Recarregando a base…" → "✔ Publicado (5 linhas)" (some sozinha) ou "✖ motivo" (fica até fechar).
 *   const op = BaseStatus.iniciar('Salvando alterações…');  op.etapa('Recarregando…');  op.ok('Salvo');  op.erro('Falhou: …');
 *   await BaseStatus.rodar('Publicando…', async op => { …; op.etapa('…'); return resultado; }, { ok: 'Publicado' });
 * `iniciar` também desabilita os botões `[data-status-bloqueia]` durante a operação (evita clique duplo). Só DOM.
 */
(function (raiz) {
    'use strict';
    const doc = raiz.document;
    let barra = null, texto = null, giro = null, fechar = null, fim = null, ativas = 0, seq = 0;

    function garantir() {
        if (barra || !doc || !doc.body) return !!barra;
        const st = doc.createElement('style');
        st.textContent = '#bs-barra{position:fixed;left:16px;bottom:44px;z-index:100000;display:none;align-items:center;gap:10px;max-width:min(560px,calc(100vw - 32px));padding:9px 14px;border-radius:8px;font:13px system-ui,sans-serif;background:#161b22;color:#e6edf3;border:1px solid #30363d;box-shadow:0 6px 24px rgba(0,0,0,.45)}' +
            '#bs-barra.bs-ok{border-color:#238636}#bs-barra.bs-erro{border-color:#da3633;background:#2a1010}' +
            '#bs-giro{width:14px;height:14px;border:2px solid #8b949e;border-top-color:#58a6ff;border-radius:50%;animation:bs-gira .8s linear infinite;flex:none}' +
            '#bs-barra.bs-ok #bs-giro,#bs-barra.bs-erro #bs-giro{animation:none;border:0;width:auto;height:auto}' +
            '#bs-x{margin-left:6px;background:none;border:0;color:#8b949e;cursor:pointer;font-size:15px}@keyframes bs-gira{to{transform:rotate(360deg)}}';
        doc.head.appendChild(st);
        barra = doc.createElement('div'); barra.id = 'bs-barra'; barra.setAttribute('role', 'status'); barra.setAttribute('aria-live', 'polite');
        giro = doc.createElement('span'); giro.id = 'bs-giro';
        texto = doc.createElement('span'); texto.id = 'bs-txt';
        fechar = doc.createElement('button'); fechar.id = 'bs-x'; fechar.textContent = '✕'; fechar.title = 'Fechar'; fechar.onclick = () => { barra.style.display = 'none'; };
        barra.append(giro, texto, fechar);
        doc.body.appendChild(barra);
        return true;
    }
    function mostrar(msg, tipo) {
        if (!garantir()) return;
        clearTimeout(fim);
        barra.className = tipo ? 'bs-' + tipo : '';
        giro.textContent = tipo === 'ok' ? '✔' : tipo === 'erro' ? '✖' : '';
        texto.textContent = msg;
        fechar.style.display = tipo === 'erro' ? '' : 'none';
        barra.style.display = 'flex';
    }
    function bloquear(sim) {
        if (!doc) return;
        doc.querySelectorAll('[data-status-bloqueia]').forEach(b => { if (sim) { b.dataset.bsAntes = b.disabled ? '1' : ''; b.disabled = true; } else if (!b.dataset.bsAntes) b.disabled = false; });
    }

    /** Começa uma operação. Devolve { etapa(txt), ok(txt), erro(txt) }; só a operação mais recente mexe na barra. */
    function iniciar(msg) {
        const meu = ++seq;
        ativas++; if (ativas === 1) bloquear(true);
        mostrar(msg, '');
        let fechada = false;
        const encerrar = () => { if (!fechada) { fechada = true; ativas = Math.max(0, ativas - 1); if (!ativas) bloquear(false); } };
        return {
            etapa(txt) { if (meu === seq && !fechada) mostrar(txt, ''); },
            ok(txt) { encerrar(); if (meu === seq) { mostrar(txt || 'Pronto', 'ok'); fim = setTimeout(() => { if (barra && meu === seq) barra.style.display = 'none'; }, 4000); } },
            erro(txt) { encerrar(); if (meu === seq) mostrar(txt || 'Falhou', 'erro'); },
        };
    }
    /** Roda `fn(op)` com a barra: ok(opcoes.ok) se terminar, erro(mensagem) se lançar (e relança). */
    async function rodar(msg, fn, opcoes) {
        const op = iniciar(msg);
        try { const r = await fn(op); op.ok(typeof (opcoes && opcoes.ok) === 'function' ? opcoes.ok(r) : (opcoes && opcoes.ok)); return r; }
        catch (e) { op.erro((opcoes && opcoes.erro ? opcoes.erro + ': ' : '') + (e && e.message ? e.message : String(e))); throw e; }
    }

    const api = { iniciar, rodar };
    raiz.BaseStatus = api;
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
