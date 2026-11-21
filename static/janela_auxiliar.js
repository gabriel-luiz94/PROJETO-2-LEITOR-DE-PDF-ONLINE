/* TASK-215 — controles compartilhados e ponto de entrada da visualização nativa. */
(function (raiz) {
    'use strict';
    const doc = raiz.document;
    const icones = {
        destacar: 'M8 3H3v10h10V8 M8 8l6-6 M9 2h5v5',
        retornar: 'M7 3H3v10h10V9 M14 3L8 9 M8 4v5h5',
        minimizar: 'M3 10h10',
        maximizar: 'M3 3h10v10H3z',
        restaurar: 'M5 3h8v8 M3 5h8v8H3z',
        fechar: 'M4 4l8 8 M12 4l-8 8'
    };
    function configurar(botao, acao, rotulo) {
        botao.dataset.janelaAcao = acao;
        botao.title = rotulo;
        botao.setAttribute('aria-label', rotulo);
        botao.innerHTML = '<svg aria-hidden="true" viewBox="0 0 16 16" fill="none" stroke="currentColor"><path d="' + icones[acao] + '"/></svg>';
    }
    function botao(acao, rotulo, executar) {
        const b = doc.createElement('button');
        b.type = 'button';
        b.className = 'janela-icone';
        configurar(b, acao, rotulo);
        b.addEventListener('click', executar);
        return b;
    }
    raiz.JanelaAuxiliar = { botao, configurar };

    const moldura = doc.getElementById('janela-nativa');
    if (!moldura) return;
    const telas = { resultado: ['/resultado_orcamento', 'Resultado'], base: ['/orcamento', 'Base de Orçamento'] };
    const ref = /^(resultado|base)=([0-9a-f]{32})$/.exec(raiz.location.hash.slice(1));
    const status = doc.getElementById('janela-nativa-status');
    if (!ref || !raiz.BroadcastChannel) { status.textContent = 'Visualização indisponível.'; return; }

    const tipo = ref[1], [url, titulo] = telas[tipo];
    const canal = new BroadcastChannel('orcacore-janela-' + ref[2]);
    const barra = moldura.querySelector('header');
    barra.querySelector('strong').textContent = doc.title = 'OrçaCore · ' + titulo;
    let epoca = 0, coletor = null, aoAbrir = null, editado = false, ligada = false;
    const enviar = (acao, dados = {}) => canal.postMessage({ acao, epoca, ...dados });
    const lerResultado = () => coletor?.() || null;
    raiz.__janelaNativa = true;
    raiz.GridRProjeto = {
        __registrarColetor(fn) { coletor = fn; if (ligada) enviar('resultado', { resultado: lerResultado(), editado }); },
        __registrarAoAbrir(fn) { aoAbrir = fn; },
        __marcarEdicaoResultado() { editado = true; if (ligada) enviar('editado'); },
        __limparEdicaoResultado() { editado = false; },
        __resultadoEditado: () => editado,
        __pendenteResultado: lerResultado,
        __enviar(parcial) { if (ligada) enviar('gravar', { parcial }); },
        __enviarAgora(parcial) { if (ligada) enviar('gravar', { parcial, agora: true }); },
        __salvarClick() { if (ligada) enviar('salvar'); },
        __salvarComoClick() { if (ligada) enviar('salvar-como'); },
        __notificarAbertura(info) { if (ligada) aoAbrir?.(info); }
    };

    const quadro = doc.createElement('iframe');
    quadro.title = titulo;
    quadro.dataset.auxiliar = tipo;
    quadro.hidden = true;
    quadro.src = url + '?janela_transferida=1';
    const conteudo = () => quadro.contentWindow.JanelaConteudo;
    const sair = acao => {
        if (!ligada) return;
        const estado = conteudo()?.capturar();
        if (!estado) { status.hidden = false; status.textContent = 'Aguarde o término da operação para trocar de janela.'; return; }
        ligada = false;
        quadro.inert = true;
        enviar(acao, { estado, resultado: lerResultado(), editado });
    };
    raiz.fecharVisualizacaoNativa = () => sair('fechar');

    barra.append(botao('retornar', 'Retornar à visualização principal', () => sair('retornar')));

    raiz.addEventListener('message', e => {
        if (e.origin === raiz.location.origin && e.source === quadro.contentWindow && e.data?.tipo === 'fechar-tela-app') sair('fechar');
    });
    let carregado = false, filaIniciada = false;
    function conectar() {
        if (!carregado || filaIniciada || !raiz.pywebview?.api.salvar_arquivo) return;
        filaIniciada = true;
        if (conteudo()) enviar('pronta');
        else enviar('falha');
    }
    raiz.addEventListener('pywebviewready', conectar, { once: true });
    quadro.addEventListener('load', () => { carregado = true; conectar(); }, { once: true });
    moldura.append(quadro);
    canal.onmessage = async ({ data }) => {
        if (!data) return;
        if (data.acao === 'iniciar') {
            epoca = data.epoca;
            try {
                await conteudo().restaurar(data.estado);
                editado = !!data.editado;
                quadro.hidden = false;
                status.hidden = true;
                enviar('conectada');
            } catch (e) { enviar('falha', { erro: e.message }); }
        } else if (data.acao === 'ativar') {
            ligada = true;
            enviar('resultado', { resultado: lerResultado(), editado });
        } else if (data.acao === 'invalidar') {
            epoca = data.epoca;
            raiz.dispatchEvent(new Event('obra-resumo-substituida'));
        } else if (data.epoca === epoca) {
            if (data.acao === 'calcular') conteudo()?.recalcular(data.pedido);
            if (data.acao === 'obra') aoAbrir?.(data.info);
            if (data.acao === 'coletar') enviar('coleta', { id: data.id, resultado: lerResultado(), editado });
            if (data.acao === 'continuar') { ligada = true; quadro.inert = false; }
        }
    };
    raiz.addEventListener('pagehide', () => canal.close(), { once: true });
})(window);
