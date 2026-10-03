/* autonomo.js — Tela de controle do modo autônomo (TASK-031, fase D). Só administrador.
   Tudo o que vem do servidor entra na página por textContent (nenhum HTML montado a partir de dado externo). */
(function () {
    'use strict';
    const $ = id => document.getElementById(id);
    const ROTULO = { processando: 'Processando', aguardando_confirmacao: 'Aguardando confirmação', ok: 'Concluída', com_pendencias: 'Com pendências', erro: 'Erro', revertida: 'Revertida' };
    let cfgAtual = null, usuarioAtual = null, ligado = false, ocupado = false, formSujo = false;

    if (!localStorage.getItem('auth_token')) { location.href = '/login'; return; }
    if (localStorage.getItem('is_admin') !== 'true') { alert('Acesso negado. Apenas administradores podem usar o modo autônomo.'); location.href = '/'; return; }

    function el(tag, cls, texto) {
        const e = document.createElement(tag);
        if (cls) e.className = cls;
        if (texto !== undefined && texto !== null) e.textContent = texto;
        return e;
    }
    function botao(texto, cls, onclick, titulo) {
        const b = el('button', cls || '', texto);
        b.type = 'button';
        if (titulo) b.title = titulo;
        b.onclick = onclick;
        return b;
    }
    async function api(metodo, url, corpo) {
        const r = await fetch(url, { method: metodo, headers: corpo !== undefined ? { 'Content-Type': 'application/json' } : {}, body: corpo !== undefined ? JSON.stringify(corpo) : undefined });
        if (r.status === 401) { location.href = '/login'; throw new Error('Sessão expirada.'); }
        if (r.status === 403) { location.href = '/'; throw new Error('Acesso negado.'); }
        let dados = null;
        try { dados = await r.json(); } catch (e) { /* sem corpo */ }
        if (!r.ok) throw new Error((dados && (typeof dados.detail === 'string' ? dados.detail : JSON.stringify(dados.detail))) || `Erro ${r.status}`);
        return dados;
    }
    function aviso(texto) { const a = $('aviso'); a.hidden = !texto; a.textContent = texto || ''; }
    function msgCfg(texto, tipo) { const m = $('cfg-msg'); m.hidden = !texto; m.className = 'msg ' + (tipo || 'ok'); m.textContent = texto || ''; }
    async function tentar(fn) { try { aviso(''); return await fn(); } catch (e) { aviso(e.message); } }

    /* ── estado ── */
    function stat(valor, rotulo) { const d = el('div', 'stat'); d.append(el('b', '', valor), el('span', '', rotulo)); return d; }
    function renderEstado(st) {
        const s = st.status;
        ligado = !!s.ligado;
        const selo = $('selo');
        selo.textContent = ligado ? 'Ligado' : 'Desligado';
        selo.className = 'selo ' + (ligado ? 'on' : 'off');
        const b = $('btn-ligar');
        b.textContent = ligado ? 'Desligar' : 'Ligar';
        b.className = ligado ? '' : 'primario';
        $('caminho-entrada').textContent = st.pastas.entrada;
        const fmt = v => v ? v.replace('T', ' ') : '—';
        const box = $('stats');
        box.replaceChildren(stat(s.processando || (ligado ? 'Aguardando arquivos' : '—'), 'Agora'), stat(fmt(s.ultima_varredura), 'Última varredura'),
            stat(String(s.na_fila), 'Arquivos na pasta'), stat(String(s.processados), 'Processados nesta sessão'), stat(String(st.aguardando_confirmacao), 'Aguardando confirmação'),
            stat(String(s.erros), 'Com erro nesta sessão'));
        $('ultimo-erro').textContent = s.ultimo_erro ? `Último erro: ${s.ultimo_erro}` : '';
        if (!st.configurado) aviso('Escolha abaixo em nome de quem as obras serão salvas e clique em “Salvar configuração” antes de ligar.');
    }

    /* ── pendências de confirmação ── */
    async function renderPendencias(lista) {
        const card = $('card-pend'), alvo = $('pend');
        const aguardando = lista.filter(x => x.status === 'aguardando_confirmacao');
        card.hidden = aguardando.length === 0;
        if (!aguardando.length) { alvo.replaceChildren(); return; }
        const detalhes = await Promise.all(aguardando.map(x => api('GET', `/api/autonomo/execucoes/${x.id}`).catch(() => null)));
        alvo.replaceChildren();
        detalhes.filter(Boolean).forEach(ex => {
            const bloco = el('div', 'pend');
            bloco.append(el('h3', '', `${ex.arquivo} · projeto ${ex.projeto_codigo} · ${ex.pendencias.length} exclusão(ões) pendente(s)`));
            ex.pendencias.forEach(p => {
                const linha = el('div', 'linha');
                linha.append(el('div', 'txt', p.descricao));
                linha.append(botao('Sim', 'bom', () => decidir(ex.id, 'confirmar', [p.indice], `Confirmar esta exclusão?\n\n${p.descricao}`), 'Confirma esta exclusão'));
                linha.append(botao('Não', 'perigo', () => decidir(ex.id, 'rejeitar', [p.indice], null), 'Mantém a linha como está'));
                bloco.append(linha);
            });
            const acoes = el('div', 'acoes');
            acoes.append(botao('Sim para todos', 'primario', () => decidir(ex.id, 'confirmar', null, `Confirmar TODAS as ${ex.pendencias.length} exclusão(ões) de “${ex.arquivo}”?`),
                'Confirma todas as exclusões deste arquivo'));
            acoes.append(botao('Não para todas', 'perigo', () => decidir(ex.id, 'rejeitar', null, `Manter TODAS as linhas de “${ex.arquivo}”? O arquivo será concluído com pendências.`)));
            acoes.append(botao('Detalhes', '', () => abrirDetalhes(ex.id)));
            bloco.append(acoes);
            alvo.append(bloco);
        });
    }
    async function decidir(id, acao, indices, pergunta) {
        if (pergunta && !confirm(pergunta)) return;
        await tentar(async () => { await api('POST', `/api/autonomo/execucoes/${id}/${acao}`, { indices }); await atualizar(true); });
    }

    /* ── histórico ── */
    function renderHistorico(lista) {
        const alvo = $('hist');
        if (!lista.length) { alvo.replaceChildren(el('p', 'vazio', 'Nenhuma execução ainda. Coloque um arquivo .dxf ou .pdf em uma subpasta da pasta de entrada.')); return; }
        const t = el('table', 'tabela');
        const cab = el('tr');
        ['Quando', 'Arquivo', 'Projeto', 'Situação', 'Mensagem', ''].forEach(h => cab.append(el('th', '', h)));
        const cabeca = el('thead');
        cabeca.append(cab);
        t.append(cabeca);
        const corpo = el('tbody');
        lista.forEach(x => {
            const tr = el('tr');
            tr.append(el('td', '', (x.criado_em || '').replace('T', ' ')), el('td', '', x.arquivo), el('td', '', x.projeto_codigo));
            tr.append(el('td', 'b-' + x.status, ROTULO[x.status] || x.status), el('td', '', x.mensagem || ''));
            const ac = el('td');
            ac.append(botao('Detalhes', '', () => abrirDetalhes(x.id)));
            if (x.status === 'ok' || x.status === 'com_pendencias') {
                ac.append(' ', botao('Reverter', '', () => reverter(x), 'Refaz a obra e o orçamento com as tabelas originais (sem os ajustes)'));
            }
            if (x.arquivo_caminho && x.status !== 'aguardando_confirmacao') ac.append(' ', botao('Reprocessar', '', () => reprocessar(x), 'Roda o arquivo original de novo (cria uma nova execução)'));
            tr.append(ac);
            corpo.append(tr);
        });
        t.append(corpo);
        alvo.replaceChildren(t);
    }
    function reverter(x) {
        if (!confirm(`Reverter “${x.arquivo}”?\n\nA obra e o orçamento serão refeitos com as tabelas ORIGINAIS, sem nenhum ajuste.`)) return;
        tentar(async () => { await api('POST', `/api/autonomo/execucoes/${x.id}/reverter`); await atualizar(true); });
    }
    function reprocessar(x) {
        if (!confirm(`Processar “${x.arquivo}” de novo?\n\nSerá criada uma nova execução; a anterior continua no histórico.`)) return;
        tentar(async () => { await api('POST', `/api/autonomo/execucoes/${x.id}/reprocessar`); await atualizar(true); });
    }

    /* ── detalhes ── */
    function lista(titulo, itens, vazio) {
        const f = document.createDocumentFragment();
        f.append(el('h4', '', titulo));
        if (!itens || !itens.length) { f.append(el('p', 'vazio', vazio || 'Nenhum.')); return f; }
        const ul = el('ul');
        itens.forEach(i => ul.append(el('li', '', typeof i === 'string' ? i : JSON.stringify(i))));
        f.append(ul);
        return f;
    }
    async function abrirDetalhes(id) {
        await tentar(async () => {
            const ex = await api('GET', `/api/autonomo/execucoes/${id}`);
            const r = ex.relatorio || {};
            const corpo = $('modal-corpo');
            corpo.replaceChildren();
            $('modal-titulo').textContent = `${ex.arquivo} · ${ROTULO[ex.status] || ex.status}`;
            corpo.append(el('p', '', ex.mensagem || ''));
            if (ex.pasta_saida) { corpo.append(el('h4', '', 'Resultados gravados em')); corpo.append(el('div', 'caminho', ex.pasta_saida)); }
            if (ex.obra_id) corpo.append(el('p', 'sub', `Obra gerada: ${ex.obra_id} (aparece em “Carregar obra” no Resumo, para o usuário dono).`));
            if (ex.pendencias && ex.pendencias.length) corpo.append(lista('Exclusões aguardando confirmação', ex.pendencias.map(p => p.descricao)));
            if (r.erro) corpo.append(el('div', 'msg erro', r.erro));
            const aj = r.ajustes || {};
            corpo.append(lista('Ajustes aplicados', aj.aplicados, 'Nenhum ajuste foi aplicado.'));
            if ((aj.rejeitados || []).length) corpo.append(lista('Exclusões rejeitadas (linhas mantidas)', aj.rejeitados));
            if ((aj.descartados || []).length) corpo.append(lista('Descartados pela validação de contrato', aj.descartados.map(d => `${d.linha_id}: ${d.motivo}`)));
            if ((aj.ignorados_por_linha_alterada || []).length) corpo.append(lista('Ignorados (a linha mudou)', aj.ignorados_por_linha_alterada));
            if ((r.ajustes_ignorados || []).length) corpo.append(lista('Ajustes cadastrados com erro (não rodaram)', r.ajustes_ignorados.map(a => `${a.ajuste}: ${(a.erros || []).join('; ')}`)));
            if (r.validacao_depois) corpo.append(lista('Validação depois dos ajustes', [`${r.validacao_depois.erro} erro(s), ${r.validacao_depois.aviso} aviso(s), ${r.validacao_depois.info} info`]));
            corpo.append(lista('Achados restantes', (r.achados_restantes || []).map(a => `[${a.severidade}] ${a.linha_id}: ${a.mensagem}`), 'Nenhum achado.'));
            if (r.orcamento) corpo.append(lista('Orçamento', [`${r.orcamento.linhas} linha(s) de orçamento`, ...(r.orcamento.nao_encontrados.length ? [`Ativos fora da base de orçamento: ${r.orcamento.nao_encontrados.join(', ')}`] : [])]));
            corpo.append(lista('Etapas', (r.etapas || []).map(e => `${e.etapa} — ${e.status} (${e.ms} ms)${e.detalhe ? ' ' + JSON.stringify(e.detalhe) : ''}`)));
            $('modal').classList.remove('oculto');
        });
    }

    /* ── configuração ── */
    async function carregarUsuarios() {
        let admins = [];
        try {
            const us = await api('GET', '/api/admin/users');
            admins = us.filter(u => (u.is_admin || u.role === 'admin') && u.ativo !== 0 && u.ativo !== false);
        } catch (e) { /* a lista é só uma conveniência */ }
        return admins;
    }
    async function montarFormulario(c, admins) {
        const sel = $('cfg-user');
        sel.replaceChildren();
        const ops = [];
        if (usuarioAtual && usuarioAtual.user_id) ops.push([String(usuarioAtual.user_id), `Eu mesmo (${usuarioAtual.email})`]);
        admins.forEach(u => { if (!ops.some(o => o[0] === String(u.id))) ops.push([String(u.id), u.email]); });
        if (c.user_id && !ops.some(o => o[0] === c.user_id)) ops.push([c.user_id, `Usuário ${c.user_id}`]);
        const vazio = new Option('— escolha —', '');
        sel.append(vazio);
        ops.forEach(([v, t]) => sel.append(new Option(t, v)));
        sel.value = c.user_id || '';
        $('cfg-base').value = c.pasta_base || '';
        $('cfg-intervalo').value = c.intervalo_s;
        $('cfg-estab').value = c.estabilizacao_s;
        $('cfg-p-entrada').value = c.pasta_entrada || '';
        $('cfg-p-processados').value = c.pasta_processados || '';
        $('cfg-p-erros').value = c.pasta_erros || '';
        $('cfg-p-saida').value = c.pasta_saida || '';
        formSujo = false;
    }
    async function salvarConfig() {
        msgCfg('');
        try {
            const corpo = {
                user_id: $('cfg-user').value, pasta_base: $('cfg-base').value.trim(), intervalo_s: parseInt($('cfg-intervalo').value, 10), estabilizacao_s: parseInt($('cfg-estab').value, 10),
                pasta_entrada: $('cfg-p-entrada').value.trim(), pasta_processados: $('cfg-p-processados').value.trim(), pasta_erros: $('cfg-p-erros').value.trim(), pasta_saida: $('cfg-p-saida').value.trim()
            };
            if (Number.isNaN(corpo.intervalo_s) || Number.isNaN(corpo.estabilizacao_s)) throw new Error('Informe números nos campos de segundos.');
            const r = await api('PUT', '/api/autonomo/config', corpo);
            cfgAtual = r.config;
            formSujo = false;
            msgCfg('Configuração salva.', 'ok');
            await atualizar(true);
        } catch (e) { msgCfg(e.message, 'erro'); }
    }

    /* ── ciclo ── */
    function queryExecucoes(extra) {
        const p = new URLSearchParams();
        if ($('filtro').value) p.set('status', $('filtro').value);
        if ($('filtro-projeto').value) p.set('projeto', $('filtro-projeto').value);
        if (extra) Object.entries(extra).forEach(([k, v]) => p.set(k, v));
        const s = p.toString();
        return '/api/autonomo/execucoes' + (s ? `?${s}` : '');
    }
    async function atualizar(forcar) {
        if (ocupado) return;
        ocupado = true;
        try {
            const [st, hist] = await Promise.all([api('GET', '/api/autonomo/status'), api('GET', queryExecucoes())]);
            renderEstado(st);
            renderHistorico(hist.execucoes);
            const precisaBuscarPendencias = $('filtro').value || $('filtro-projeto').value;
            const todas = precisaBuscarPendencias
                ? (await api('GET', queryExecucoes({ status: 'aguardando_confirmacao' }))).execucoes
                : hist.execucoes;
            await renderPendencias(todas);
        } catch (e) { aviso(e.message); } finally { ocupado = false; }
    }
    async function carregarProjetosFiltro() {
        try {
            const r = await api('GET', '/api/projetos');
            const projetos = Array.isArray(r) ? r : (r.projetos || []);
            const sel = $('filtro-projeto');
            const atual = sel.value;
            sel.replaceChildren(new Option('Todos os projetos', ''));
            projetos.forEach(p => sel.append(new Option(`${p.codigo} — ${p.nome}`, p.codigo)));
            sel.value = atual;
        } catch (e) { /* o filtro é só uma conveniência: sem a lista, segue funcionando sem opções */ }
    }
    async function iniciar() {
        await tentar(async () => {
            const c = await api('GET', '/api/autonomo/config');
            cfgAtual = c.config; usuarioAtual = c.usuario_atual;
            await montarFormulario(cfgAtual, await carregarUsuarios());
            await carregarProjetosFiltro();
            await atualizar(true);
        });
    }

    $('btn-ligar').onclick = () => tentar(async () => {
        if (ligado) { await api('POST', '/api/autonomo/desligar'); } else { await api('POST', '/api/autonomo/ligar'); }
        await atualizar(true);
    });
    $('btn-varrer').onclick = () => tentar(async () => { await api('POST', '/api/autonomo/varrer'); await atualizar(true); });
    $('btn-atualizar').onclick = () => atualizar(true);
    $('filtro').onchange = () => atualizar(true);
    $('filtro-projeto').onchange = () => atualizar(true);
    $('btn-salvar').onclick = salvarConfig;
    $('modal-fechar').onclick = () => $('modal').classList.add('oculto');
    $('modal').onclick = e => { if (e.target === $('modal')) $('modal').classList.add('oculto'); };
    document.addEventListener('keydown', e => { if (e.key === 'Escape') $('modal').classList.add('oculto'); });
    document.querySelectorAll('#cfg-user, #cfg-base, #cfg-intervalo, #cfg-estab, #cfg-p-entrada, #cfg-p-processados, #cfg-p-erros, #cfg-p-saida').forEach(e => e.addEventListener('input', () => { formSujo = true; }));
    setInterval(() => { if (!document.hidden && $('modal').classList.contains('oculto')) atualizar(false); }, 5000);
    iniciar();
})();
