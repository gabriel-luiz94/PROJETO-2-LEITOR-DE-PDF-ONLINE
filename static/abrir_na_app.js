/*
 * abrir_na_app.js (TASK-128) — abre telas do programa DENTRO da aplicação desktop, em vez de uma nova aba no navegador.
 *   • abrir(url, titulo): tela cheia por cima da página atual (iframe), com botão Fechar — a página de baixo (ex.: o Resumo, com as tabelas em edição) não é descarregada.
 *   • navegar(url): troca a página atual (para fluxos que já levam tudo pelo localStorage, como Leitor → Resumo).
 * Na versão online (servidor, `/api/version` → mode "server") continua abrindo uma nova aba do navegador, como sempre foi.
 * As telas abertas aqui fecham pelo botão "Voltar" delas (voltarOuFechar → postMessage) ou pelo Fechar da moldura.
 */
(function (raiz) {
    'use strict';
    const doc = raiz.document;
    let modo = null;                                           // 'desktop' | 'server' (desconhecido = trata como desktop: abrir dentro do programa nunca quebra)
    try {
        fetch('/api/version').then(r => (r.ok ? r.json() : null)).then(d => { if (d && d.mode) modo = d.mode; }).catch(() => {});
    } catch (e) { /* sem fetch: segue como desktop */ }

    const noServidor = () => modo === 'server';

    function fechar() {
        const ov = doc.getElementById('abrir-na-app-overlay');
        if (ov) ov.remove();
        doc.removeEventListener('keydown', aoTeclar, true);
    }
    function aoTeclar(e) { if (e.key === 'Escape') fechar(); }

    function abrir(url, titulo) {
        if (noServidor()) { raiz.open(url, '_blank'); return null; }
        fechar();
        const ov = doc.createElement('div');
        ov.id = 'abrir-na-app-overlay';
        ov.style.cssText = 'position:fixed;inset:0;z-index:100000;background:#0d1117;display:flex;flex-direction:column;';
        const barra = doc.createElement('div');
        barra.style.cssText = 'display:flex;align-items:center;gap:12px;padding:6px 12px;background:#161b22;border-bottom:1px solid #30363d;color:#e6edf3;font:600 13px system-ui,sans-serif;';
        const tit = doc.createElement('span'); tit.textContent = titulo || ''; tit.style.flex = '1';
        const btn = doc.createElement('button');
        btn.textContent = '✕ Fechar'; btn.title = 'Voltar para a tela anterior (Esc)';
        btn.style.cssText = 'background:#21262d;color:#e6edf3;border:1px solid #30363d;border-radius:6px;padding:4px 12px;cursor:pointer;font:inherit;';
        btn.addEventListener('click', fechar);
        barra.append(tit, btn);
        const quadro = doc.createElement('iframe');
        quadro.src = url; quadro.title = titulo || '';
        quadro.style.cssText = 'flex:1;border:0;width:100%;background:#0d1117;';
        ov.append(barra, quadro);
        doc.body.appendChild(ov);
        doc.addEventListener('keydown', aoTeclar, true);
        return ov;
    }

    function navegar(url) {
        if (noServidor()) { raiz.open(url, '_blank'); return; }
        raiz.location.href = url;
    }

    // Quem foi aberto numa moldura avisa a página de baixo para fechá-la
    raiz.addEventListener('message', e => { if (e.origin === raiz.location.origin && e.data && e.data.tipo === 'fechar-tela-app') fechar(); });

    /** Para o botão "Voltar" das telas abertas por `abrir`: dentro da moldura fecha a moldura; fora dela vai para `destino`. */
    function voltarOuFechar(destino) {
        if (raiz.self !== raiz.top) raiz.parent.postMessage({ tipo: 'fechar-tela-app' }, raiz.location.origin);
        else raiz.location.href = destino || '/';
    }

    raiz.AbrirNaApp = { abrir, navegar, fechar, voltarOuFechar };
})(window);
