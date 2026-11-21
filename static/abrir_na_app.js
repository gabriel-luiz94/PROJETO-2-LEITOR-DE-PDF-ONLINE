/*
 * abrir_na_app.js (TASK-128) — abre telas do programa DENTRO da aplicação desktop, em vez de uma nova aba no navegador.
 *   • abrir(url, titulo, chave): com chave, janela auxiliar no Resumo. Sem chave, mantém a abertura de tela cheia no desktop.
 *   • navegar(url): troca a página atual (para fluxos que já levam tudo pelo localStorage, como Leitor → Resumo).
 * Na versão online, as telas sem chave continuam abrindo em uma nova aba do navegador.
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
        const url = ov ? ((ov.querySelector('iframe') || {}).src || '') : '';
        if (ov) ov.remove();
        doc.removeEventListener('keydown', aoTeclar, true);
        // TASK-142: a página de baixo pode ter dados que a tela fechada alterou (ex.: a Tabela de orçamento) — avisa para recarregar
        if (ov) { try { doc.dispatchEvent(new CustomEvent('tela-app-fechada', { detail: { url } })); } catch (e) { /* navegador antigo */ } }
    }
    function aoTeclar(e) { if (e.key === 'Escape') fechar(); }

    const paineis = new Map();
    const areaDeTrabalho = doc.getElementById('work');
    const topoDaArea = () => Math.round(areaDeTrabalho?.getBoundingClientRect().top || 0) + 8;
    function porNaFrente(el) {
        const ordem = [...paineis.values()].map(p => p.el).filter(outro => outro !== el);
        ordem.sort((a, b) => Number(a.style.zIndex) - Number(b.style.zIndex));
        ordem.push(el);
        ordem.forEach((janela, i) => { janela.style.zIndex = 8400 + i; });
    }

    function trazerPainel(chave) {
        const p = paineis.get(chave);
        if (!p) return false;
        p.el.classList.remove('janela-minimizada');
        p.minimizar.textContent = 'Minimizar';
        porNaFrente(p.el);
        p.barra.focus({ preventScroll: true });
        return true;
    }

    function removerPainel(chave) {
        const p = paineis.get(chave);
        if (!p) return;
        paineis.delete(chave);
        p.el.remove();
        if (p.url) doc.dispatchEvent(new CustomEvent('tela-app-fechada', { detail: { url: p.url } }));
    }

    function fecharPainel(chave) {
        const p = paineis.get(chave);
        if (p && p.aoFechar?.() !== false) removerPainel(chave);
    }

    function painel(chave, titulo, conteudo, opcoes = {}) {
        if (trazerPainel(chave)) return paineis.get(chave);
        ajustarArea();
        const el = doc.createElement('section');
        el.className = 'janela-auxiliar';
        el.setAttribute('aria-label', titulo);
        el.style.setProperty('--janela-deslocamento', (paineis.size * 30) + 'px');
        const barra = doc.createElement('header');
        barra.tabIndex = -1;
        const nome = doc.createElement('strong');
        nome.textContent = titulo;
        barra.append(nome);
        const botao = (rotulo, acao) => {
            const b = doc.createElement('button');
            b.type = 'button'; b.textContent = rotulo;
            b.addEventListener('click', acao);
            barra.append(b);
            return b;
        };
        const minimizar = botao('Minimizar', () => {
            el.classList.toggle('janela-minimizada');
            minimizar.textContent = el.classList.contains('janela-minimizada') ? 'Restaurar' : 'Minimizar';
        });
        const ampliar = botao('Ampliar', () => {
            el.classList.toggle('janela-ampliada');
            ampliar.textContent = el.classList.contains('janela-ampliada') ? 'Reduzir' : 'Ampliar';
        });
        if (opcoes.aoReencaixar) botao('Voltar ao modal', opcoes.aoReencaixar);
        botao('Fechar', () => fecharPainel(chave));
        el.append(barra, conteudo);
        el.addEventListener('pointerdown', () => porNaFrente(el), true);
        el.addEventListener('focusin', () => porNaFrente(el));
        if (conteudo.tagName === 'IFRAME') {
            conteudo.addEventListener('load', () => {
                conteudo.contentDocument?.addEventListener('pointerdown', () => porNaFrente(el), true);
            });
        }
        barra.addEventListener('pointerdown', e => {
            if (e.button !== 0 || e.target.closest('button') || el.classList.contains('janela-ampliada')) return;
            const box = el.getBoundingClientRect();
            const dx = e.clientX - box.left, dy = e.clientY - box.top;
            barra.setPointerCapture(e.pointerId);
            barra.onpointermove = ev => {
                el.style.left = Math.max(0, Math.min(raiz.innerWidth - el.offsetWidth, ev.clientX - dx)) + 'px';
                el.style.top = Math.max(topoDaArea(), Math.min(raiz.innerHeight - el.offsetHeight, ev.clientY - dy)) + 'px';
            };
            barra.onlostpointercapture = () => { barra.onpointermove = null; };
        });
        doc.body.append(el);
        const p = { el, barra, minimizar, conteudo, aoFechar: opcoes.aoFechar, url: opcoes.url };
        paineis.set(chave, p);
        trazerPainel(chave);
        return p;
    }

    function ajustarArea() {
        const y = topoDaArea();
        doc.documentElement.style.setProperty('--janela-area-topo', y + 'px');
        paineis.forEach(({ el }) => {
            if (!el.style.left || el.classList.contains('janela-ampliada')) return;
            el.style.left = Math.max(0, Math.min(parseFloat(el.style.left), raiz.innerWidth - el.offsetWidth)) + 'px';
            el.style.top = Math.max(y, Math.min(parseFloat(el.style.top), raiz.innerHeight - el.offsetHeight)) + 'px';
        });
    }
    raiz.addEventListener('resize', ajustarArea);
    if (areaDeTrabalho && raiz.ResizeObserver) new raiz.ResizeObserver(ajustarArea).observe(areaDeTrabalho);

    function abrirEmPainel(url, titulo, chave) {
        const existente = paineis.get(chave);
        if (existente) {
            if (chave === 'resultado' || existente.url !== url) {
                existente.url = url;
                existente.conteudo.src = url;
            }
            trazerPainel(chave);
            return existente.el;
        }
        const quadro = doc.createElement('iframe');
        quadro.src = url; quadro.title = titulo;
        return painel(chave, titulo, quadro, { url }).el;
    }

    function abrir(url, titulo, chave) {
        if (chave) return abrirEmPainel(url, titulo, chave);
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
    raiz.addEventListener('message', e => {
        if (e.origin !== raiz.location.origin || e.data?.tipo !== 'fechar-tela-app') return;
        for (const [chave, p] of paineis) {
            if (p.conteudo.contentWindow === e.source) { fecharPainel(chave); return; }
        }
        if (doc.querySelector('#abrir-na-app-overlay iframe')?.contentWindow === e.source) fechar();
    });

    /** Para o botão "Voltar" das telas abertas por `abrir`: dentro da moldura fecha a moldura; fora dela vai para `destino`. */
    function voltarOuFechar(destino) {
        if (raiz.self !== raiz.top) raiz.parent.postMessage({ tipo: 'fechar-tela-app' }, raiz.location.origin);
        else raiz.location.href = destino || '/';
    }

    raiz.AbrirNaApp = { abrir, navegar, fechar, voltarOuFechar, painel, trazerPainel, fecharPainel, removerPainel };
})(window);
