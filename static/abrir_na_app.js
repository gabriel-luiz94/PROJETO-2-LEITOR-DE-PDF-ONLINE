/*
 * abrir_na_app.js (TASK-128) — abre telas do programa DENTRO da aplicação desktop, em vez de uma nova aba no navegador.
 *   • abrir(url, titulo, chave): com chave, visualização do Resumo em tela inteira, minimizável e destacável no desktop.
 *     Sem chave, mantém a moldura simples no desktop.
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
        raiz.JanelaAuxiliar.configurar(p.minimizar, 'minimizar', 'Minimizar');
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
        const icone = raiz.JanelaAuxiliar.botao;
        const minimizar = icone('minimizar', 'Minimizar', () => {
            const recolhida = el.classList.toggle('janela-minimizada');
            raiz.JanelaAuxiliar.configurar(minimizar, recolhida ? 'maximizar' : 'minimizar', recolhida ? 'Restaurar' : 'Minimizar');
        });
        const ampliar = icone('maximizar', 'Maximizar', () => {
            const ampla = el.classList.toggle('janela-ampliada');
            raiz.JanelaAuxiliar.configurar(ampliar, ampla ? 'restaurar' : 'maximizar', ampla ? 'Restaurar tamanho' : 'Maximizar');
        });
        barra.append(minimizar, ampliar);
        if (opcoes.aoReencaixar) botao('Voltar ao modal', opcoes.aoReencaixar);
        barra.append(icone('fechar', 'Fechar', () => fecharPainel(chave)));
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

    const visualizacoes = new Map();
    let epocaObra = 0, naFrente = null;
    const ponte = () => raiz.pywebview?.api;
    const avisar = mensagem => raiz.showToast?.(mensagem);
    const conteudoDa = p => p.quadro?.contentDocument && p.quadro.contentWindow.JanelaConteudo;
    const donoProjeto = () => raiz.GridRProjeto;

    function abaDe(p) {
        const rodape = doc.getElementById('janelas-rodape');
        p.aba.hidden = p.estado !== 'minimizada';
        rodape.hidden = ![...visualizacoes.values()].some(outra => !outra.aba.hidden);
    }

    function exibir(p, inteira) {
        // Só uma visualização ocupa a área principal de cada vez. A anterior continua no rodapé.
        if (naFrente && naFrente !== p) minimizarVisualizacao(naFrente);
        naFrente = p;
        p.estado = 'embutida';
        if (inteira) p.el.classList.remove('janela-restaurada');
        p.el.hidden = false;
        raiz.JanelaAuxiliar.configurar(p.maximo, p.el.classList.contains('janela-restaurada') ? 'maximizar' : 'restaurar',
            p.el.classList.contains('janela-restaurada') ? 'Maximizar' : 'Restaurar tamanho');
        abaDe(p);
    }

    function minimizarVisualizacao(p) {
        p.estado = 'minimizada';
        p.el.hidden = true; // mantém o iframe e seu trabalho
        if (naFrente === p) naFrente = null;
        abaDe(p);
    }

    function fecharVisualizacao(p) {
        if (p.nativa) desligarNativa(p);
        visualizacoes.delete(p.chave);
        if (naFrente === p) naFrente = null;
        p.el.remove();
        abaDe(p);
        p.aba.remove();
        doc.dispatchEvent(new CustomEvent('tela-app-fechada', { detail: { url: p.url } }));
    }

    function criarQuadro(p, estado, pronto) {
        const quadro = doc.createElement('iframe');
        quadro.title = p.titulo;
        quadro.dataset.auxiliar = p.chave;
        quadro.src = estado ? p.url + '?janela_transferida=1' : p.url;
        if (estado) quadro.hidden = true;
        p.quadro = quadro;
        quadro.addEventListener('load', async () => {
            if (p.quadro !== quadro) return;
            try {
                if (estado) await conteudoDa(p).restaurar(estado);
                quadro.hidden = false;
                pronto?.();
                if (p.recalculoPendente && conteudoDa(p)) {
                    p.recalculoPendente = false;
                    conteudoDa(p).recalcular(raiz.__orcamentoDaSessao);
                }
            } catch (e) {
                avisar('Não foi possível restaurar a visualização.');
                pronto?.(e);
            }
        }, { once: true });
        p.el.append(quadro);
    }

    function desligarNativa(p, fechar = true, limparProjeto = true) {
        const n = p.nativa;
        p.nativa = null;
        n.canal.close();
        clearTimeout(n.tempo);
        n.pedidos.forEach(fn => fn(n.resultado));
        if (limparProjeto && n.ativa && p.chave === 'resultado') {
            donoProjeto()?.__registrarColetor(null);
            donoProjeto()?.__registrarAoAbrir(null);
        }
        if (fechar) ponte()?.janela_auxiliar_fechar(n.id).catch(() => {});
    }

    function recuperarFalha(p, mensagem) {
        desligarNativa(p);
        if (p.quadro) p.quadro.inert = false;
        p.destacar.disabled = !ponte()?.janela_auxiliar_abrir;
        if (p.recalculoPendente && conteudoDa(p)) {
            p.recalculoPendente = false;
            conteudoDa(p).recalcular(raiz.__orcamentoDaSessao);
        }
        avisar(mensagem);
    }

    function resultadoNativo(p) {
        const n = p.nativa;
        if (!n?.ativa) return null;
        return new Promise(resolve => {
            const id = ++n.sequencia;
            const tempo = setTimeout(() => {
                n.pedidos.delete(id);
                resolve(n.resultado);
            }, 1200);
            n.pedidos.set(id, resultado => { clearTimeout(tempo); resolve(resultado); });
            n.enviar('coletar', { id });
        });
    }

    function consumirNativa(p, msg) {
        const n = p.nativa, projeto = donoProjeto();
        if (!n || !msg) return;
        if (msg.acao === 'pronta' && !n.ativa) {
            const estado = conteudoDa(p)?.capturar();
            if (!estado) { recuperarFalha(p, 'Aguarde a conclusão da operação para destacar a janela.'); return; }
            p.quadro.inert = true;
            n.enviar('iniciar', { estado, editado: projeto?.resultadoEditado() });
        } else if (msg.acao === 'conectada' && !n.ativa) {
            clearTimeout(n.tempo);
            conteudoDa(p)?.suspender?.();
            p.quadro.remove();
            p.quadro = null;
            p.estado = 'destacada';
            p.el.hidden = true;
            if (naFrente === p) naFrente = null;
            abaDe(p);
            n.ativa = true;
            if (p.chave === 'resultado') {
                projeto?.__registrarColetor(() => resultadoNativo(p));
                projeto?.__registrarAoAbrir(info => {
                    n.infoObra = info;
                    n.enviar('obra', { info });
                });
            }
            n.enviar('ativar');
            if (p.recalculoPendente) {
                p.recalculoPendente = false;
                n.enviar('calcular', { pedido: raiz.__orcamentoDaSessao });
            }
        } else if (msg.acao === 'falha' && !n.ativa) {
            if (msg.erro) console.error('Falha ao transferir a visualização:', msg.erro);
            recuperarFalha(p, 'Não foi possível conectar a janela nativa.');
        } else if (n.ativa && msg.epoca === n.epoca) {
            if (p.chave === 'resultado') {
                if (msg.acao === 'editado' || msg.editado) projeto?.__marcarEdicaoResultado();
                if (['resultado', 'coleta', 'retornar', 'fechar'].includes(msg.acao)) n.resultado = msg.resultado;
                if (msg.acao === 'coleta') {
                    n.pedidos.get(msg.id)?.(msg.resultado);
                    n.pedidos.delete(msg.id);
                }
                if (msg.acao === 'gravar') {
                    const parcial = msg.parcial || {};
                    // Só os dados produzidos pelo Resultado transitam por este canal.
                    const permitido = {};
                    if (parcial.resultado) { permitido.resultado = n.resultado = parcial.resultado; }
                    if (parcial.sap) permitido.sap = parcial.sap;
                    if (Object.keys(permitido).length) (msg.agora ? projeto?.__enviarAgora(permitido) : projeto?.__enviar(permitido));
                }
                if (msg.acao === 'salvar') projeto?.__salvarClick();
                if (msg.acao === 'salvar-como') projeto?.__salvarComoClick();
            }
            if (msg.acao === 'fechar' || msg.acao === 'retornar') {
                if (msg.resultado && p.chave === 'resultado') projeto?.__enviar({ resultado: msg.resultado });
                if (msg.acao === 'fechar') { fecharVisualizacao(p); return; }
                n.reencaixando = true;
                criarQuadro(p, msg.estado, erro => {
                    if (p.nativa !== n) return;
                    if (erro) {
                        p.quadro.remove(); p.quadro = null;
                        n.reencaixando = false;
                        n.enviar('continuar');
                        return;
                    }
                    if (p.chave === 'resultado' && n.epoca !== msg.epoca) conteudoDa(p)?.obra(n.infoObra);
                    desligarNativa(p, true, false);
                    if (msg.editado && p.chave === 'resultado') projeto?.__marcarEdicaoResultado();
                    p.destacar.disabled = false;
                    exibir(p, true);
                });
            }
        }
    }

    async function destacarVisualizacao(p) {
        const api = ponte();
        if (!api?.janela_auxiliar_abrir || p.nativa || !raiz.BroadcastChannel) return;
        if (!conteudoDa(p)?.capturar()) { avisar('Aguarde a conclusão da operação para destacar a janela.'); return; }
        const id = raiz.crypto.randomUUID().replace(/-/g, '');
        const canal = new BroadcastChannel('orcacore-janela-' + id);
        const n = p.nativa = {
            id, canal, epoca: epocaObra, ativa: false, resultado: null, sequencia: 0, pedidos: new Map(),
            enviar: (acao, dados = {}) => canal.postMessage({ acao, epoca: n.epoca, ...dados })
        };
        p.destacar.disabled = true;
        canal.onmessage = ({ data }) => consumirNativa(p, data);
        n.tempo = setTimeout(() => recuperarFalha(p, 'A janela nativa não respondeu.'), 12000);
        try {
            const resposta = await api.janela_auxiliar_abrir(p.chave, id);
            if (p.nativa === n && !resposta?.ok) recuperarFalha(p, resposta?.erro || 'Não foi possível destacar a janela.');
            if (p.nativa !== n && resposta?.ok) await api.janela_auxiliar_fechar(id);
        } catch (e) {
            if (p.nativa === n) recuperarFalha(p, 'Não foi possível destacar a janela.');
        }
    }

    function nativaFechada(id) {
        const p = [...visualizacoes.values()].find(v => v.nativa?.id === id);
        if (!p) return;
        if (p.nativa.ativa) {
            desligarNativa(p, false);
            if (!p.quadro) fecharVisualizacao(p);
        } else recuperarFalha(p, 'A janela nativa foi fechada.');
    }

    function abrirEmPainel(url, titulo, chave) {
        const existente = visualizacoes.get(chave);
        if (existente) {
            if (chave === 'resultado') {
                if (existente.nativa?.ativa && !existente.nativa.reencaixando) existente.nativa.enviar('calcular', { pedido: raiz.__orcamentoDaSessao });
                else if (!existente.nativa && conteudoDa(existente)) conteudoDa(existente).recalcular(raiz.__orcamentoDaSessao);
                else existente.recalculoPendente = true;
            }
            if (!existente.nativa?.ativa) exibir(existente, true);
            return existente.el;
        }
        const el = doc.createElement('section');
        el.className = 'janela-inteira';
        el.setAttribute('aria-label', titulo);
        const barra = doc.createElement('header');
        barra.className = 'janela-titulo';
        const nome = doc.createElement('strong');
        nome.textContent = titulo;
        const aba = doc.createElement('div');
        aba.className = 'janela-aba';
        const nomeAba = doc.createElement('strong');
        nomeAba.textContent = chave === 'resultado' ? 'Resultado' : titulo;
        const p = { el, aba, titulo, chave, url };
        const b = raiz.JanelaAuxiliar.botao;
        p.destacar = b('destacar', 'Abrir em nova janela', () => destacarVisualizacao(p));
        const atualizarBotao = () => {
            p.destacar.disabled = !ponte()?.janela_auxiliar_abrir || !raiz.BroadcastChannel;
        };
        atualizarBotao();
        raiz.addEventListener('pywebviewready', atualizarBotao, { once: true });
        p.maximo = b('maximizar', 'Maximizar', () => {
            el.classList.toggle('janela-restaurada');
            exibir(p, false);
        });
        barra.append(nome, p.destacar, b('minimizar', 'Minimizar', () => minimizarVisualizacao(p)),
            p.maximo, b('fechar', 'Fechar', () => fecharVisualizacao(p)));
        aba.append(nomeAba, b('maximizar', 'Maximizar ' + nomeAba.textContent, () => exibir(p, true)),
            b('fechar', 'Fechar ' + nomeAba.textContent, () => fecharVisualizacao(p)));
        visualizacoes.set(chave, p);
        doc.getElementById('janelas-rodape').append(aba);
        el.append(barra);
        doc.body.append(el);
        criarQuadro(p);
        exibir(p, true);
        return el;
    }

    raiz.addEventListener('obra-resumo-substituida', () => {
        epocaObra++;
        const p = visualizacoes.get('resultado');
        if (!p?.nativa) return;
        if (!p.nativa.ativa) { recuperarFalha(p, 'A obra mudou durante a abertura da janela.'); return; }
        p.nativa.resultado = null;
        p.nativa.infoObra = null;
        p.nativa.epoca = epocaObra;
        p.nativa.pedidos.forEach(fn => fn(null));
        p.nativa.pedidos.clear();
        p.nativa.enviar('invalidar');
    });

    raiz.addEventListener('pagehide', () => {
        visualizacoes.forEach(p => { if (p.nativa) desligarNativa(p); });
    });

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
        for (const p of visualizacoes.values()) {
            if (p.quadro?.contentWindow === e.source) { fecharVisualizacao(p); return; }
        }
        if (doc.querySelector('#abrir-na-app-overlay iframe')?.contentWindow === e.source) fechar();
    });

    /** Para o botão "Voltar" das telas abertas por `abrir`: dentro da moldura fecha a moldura; fora dela vai para `destino`. */
    function voltarOuFechar(destino) {
        if (raiz.self !== raiz.top) raiz.parent.postMessage({ tipo: 'fechar-tela-app' }, raiz.location.origin);
        else raiz.location.href = destino || '/';
    }

    raiz.AbrirNaApp = { abrir, navegar, fechar, voltarOuFechar, painel, trazerPainel, fecharPainel, removerPainel, nativaFechada };
})(window);
