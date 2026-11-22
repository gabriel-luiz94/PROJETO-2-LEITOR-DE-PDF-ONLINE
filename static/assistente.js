/*
 * assistente.js (TASK-191) — o lado do navegador do assistente de IA com ferramentas e skills.
 *   · funções PURAS (testadas em Node, tests/test_assistente_front_191.py): ler o stream NDJSON, montar a cópia das tabelas que vai ao servidor,
 *     descrever e APLICAR um plano de alterações com checagem de que a linha ainda é a mesma;
 *   · DOM (só no navegador): o cartão do plano (Aplicar / Descartar / Desfazer) e o editor de Skills (listar, editar, criar, apagar, gerar o catálogo).
 * O servidor nunca altera a tabela: ele PROPÕE; quem aplica é o clique do usuário aqui.
 */
(function (raiz) {
    'use strict';

    const CAMPOS_LINHA = ['entidade', 'operacao', 'ativo', 'qtdAtivos', 'congelada', 'bloqueada', 'vinculoEstruturas', 'vinculoPostes', 'estruturasFicticias', 'cabosFicticios', 'cabosFicticiosSlots', '_poste', '_geo', 'fases', 'tensao', 'nivel', 'textoOriginal'];
    const LIMITE_LINHAS = 2000;

    /** Lê um bloco de texto do stream: devolve os eventos das linhas COMPLETAS e o resto (linha ainda pela metade). Linha que não é JSON é ignorada. */
    function linhasNdjson(buffer) {
        const partes = String(buffer == null ? '' : buffer).split('\n');
        const resto = partes.pop();
        const eventos = [];
        partes.forEach(l => {
            if (!l.trim()) return;
            try { const e = JSON.parse(l); if (e && typeof e === 'object' && !Array.isArray(e)) eventos.push(e); } catch (x) { /* linha corrompida: segue */ }
        });
        return { eventos, resto };
    }

    /** As linhas da tela como o servidor precisa: {id = posição na tela, entidade, operacao, ativo, ...}. Linhas vazias (null) ficam de fora, mas o id continua sendo a posição real. */
    function tabelaParaServidor(estados) {
        const saida = { cabos: [], outros: [] };
        ['cabos', 'outros'].forEach(t => {
            const dados = (estados && estados[t] && estados[t].data) || [];
            for (let i = 0; i < dados.length; i++) {
                const r = dados[i];
                if (!r || typeof r !== 'object') continue;
                if (saida[t].length >= LIMITE_LINHAS) { saida.tabela_incompleta = true; break; }
                const o = { id: i };
                CAMPOS_LINHA.forEach(c => { if (r[c] !== undefined && r[c] !== null && r[c] !== '') o[c] = r[c]; });
                saida[t].push(o);
            }
        });
        return saida;
    }

    const ROTULO_TABELA = { cabos: 'Cabos', outros: 'Outros' };
    const SIMBOLO = { adicionar: '＋', editar: '✎', remover: '−', obra: '📦' };

    /** Uma linha do plano em palavras: {simbolo, classe, tabela, operacao, texto, antes, avisos}. */
    function descreverItem(it) {
        if (it.acao === 'obra') {
            const verbo = { adicionar: 'Adicionar', subtrair: 'Subtrair', carregar: 'Carregar (substitui a tela)' }[it.modo] || it.modo;
            return { simbolo: SIMBOLO.obra, classe: 'obra', tabela: '', operacao: '', texto: `${verbo} a obra «${it.nome || it.obra_id}»`, antes: '', avisos: it.avisos || [] };
        }
        return {
            simbolo: SIMBOLO[it.acao] || '?', classe: it.acao, tabela: ROTULO_TABELA[it.tabela] || it.tabela, operacao: it.operacao || '',
            texto: it.ativo || '', antes: it.acao === 'editar' && it.antes ? [it.antes.operacao, it.antes.ativo].filter(Boolean).join(' · ') : '', avisos: it.avisos || [],
        };
    }

    /**
     * Aplica o plano em `estados` ({cabos: {data}, outros: {data}}). `entidadeDe(ativo, tabela)` dá a entidade das linhas novas/editadas.
     * Editar/remover exige que a linha permaneça editável e corresponda à proposta;
     * senão a alteração é IGNORADA e listada. Devolve {adicionadas, editadas, removidas, ignoradas: [{n, motivo}]}.
     */
    function aplicarPlano(estados, itens, entidadeDe) {
        const r = { adicionadas: 0, editadas: 0, removidas: 0, ignoradas: [] };
        const ent = (a, t, linha) => (typeof entidadeDe === 'function' ? entidadeDe(a, t, linha) : (t === 'cabos' ? 'CABO' : '0'));
        (itens || []).forEach(it => {
            if (!it || !['adicionar', 'editar', 'remover'].includes(it.acao) || !estados[it.tabela]) return;
            const dados = estados[it.tabela].data;
            if (it.acao === 'adicionar') {
                dados.push({ entidade: ent(it.ativo, it.tabela), operacao: it.operacao || 'I', ativo: it.ativo });
                r.adicionadas++;
                return;
            }
            const atual = Number.isInteger(it.id) ? dados[it.id] : null;
            if (!atual) { r.ignoradas.push({ n: it.n, motivo: 'a linha não existe mais' }); return; }
            if (atual.bloqueada || atual.congelada) { r.ignoradas.push({ n: it.n, motivo: 'a linha está protegida' }); return; }
            if (it.antes && (String(atual.ativo || '').trim() !== String(it.antes.ativo || '').trim() ||
                (it.antes.operacao !== undefined && (atual.operacao || '') !== it.antes.operacao))) { r.ignoradas.push({ n: it.n, motivo: 'a linha mudou depois do pedido' }); return; }
            if (it.acao === 'remover') { dados[it.id] = null; r.removidas++; return; }
            atual.operacao = it.operacao || atual.operacao || 'I';
            atual.ativo = it.ativo;
            if (!atual.entidadeManual) atual.entidade = ent(it.ativo, it.tabela, atual);
            r.editadas++;
        });
        ['cabos', 'outros'].forEach(t => { if (estados[t]) estados[t].data = estados[t].data.filter(x => x !== null); });
        return r;
    }

    function resumoAplicacao(r) {
        const p = [];
        if (r.adicionadas) p.push(`${r.adicionadas} adicionada(s)`);
        if (r.editadas) p.push(`${r.editadas} editada(s)`);
        if (r.removidas) p.push(`${r.removidas} removida(s)`);
        let t = p.length ? p.join(', ') : 'nada aplicado';
        if (r.ignoradas.length) t += `; ${r.ignoradas.length} ignorada(s) (${[...new Set(r.ignoradas.map(x => x.motivo))].join('; ')})`;
        return t;
    }

    /** Texto seguro (sem HTML) com quebras de linha preservadas por CSS (white-space: pre-wrap). */
    function textoSeguro(el, texto) { el.textContent = texto; return el; }

    /** Atalhos do formulário do provedor reserva: DeepSeek, Qwen etc. entram como «OpenAI-compatível» (URL base + nome do modelo). Confira o nome do modelo na doc de cada serviço: eles mudam. */
    const PRESETS_RESERVA = [
        { id: 'deepseek', rotulo: 'DeepSeek', provider: 'openai', url: 'https://api.deepseek.com', modelo: 'deepseek-flash', dica: 'Crie uma chave em platform.deepseek.com. O teste consulta Cabos e confirma a resposta final.' },
        { id: 'qwen', rotulo: 'Qwen (Alibaba · Singapura)', provider: 'openai', url: 'https://dashscope-intl.aliyuncs.com/compatible-mode/v1', modelo: 'qwen-plus', dica: 'Chave do Model Studio (região Singapura): ela só vale na região em que foi criada.' },
        { id: 'openrouter', rotulo: 'OpenRouter', provider: 'openai', url: 'https://openrouter.ai/api/v1', modelo: '', dica: 'Informe o modelo (os gratuitos terminam em :free e têm limite de 20 pedidos/min e 50/dia).' },
        { id: 'ollama', rotulo: 'Ollama (no seu PC)', provider: 'openai', url: 'http://localhost:11434/v1', modelo: '', dica: 'Informe o modelo instalado (ex.: um Qwen). O Ollama não usa chave: digite qualquer texto.' },
        { id: 'gemini', rotulo: 'Gemini (outra chave)', provider: 'gemini', url: '', modelo: '', dica: '' },
        { id: 'claude', rotulo: 'Claude', provider: 'claude', url: '', modelo: '', dica: '' },
    ];

    /** Corpo do PUT /api/assistente/reserva a partir do formulário: URL só vale para OpenAI-compatível; chave vazia = manter a salva. */
    function corpoReserva(f) {
        const provider = String(f.provider || 'openai');
        return { ativa: !!f.ativa, provider, base_url: provider === 'openai' ? String(f.url || '').trim() : '', model: String(f.modelo || '').trim(),
                 api_key: String(f.chave || '').trim() || null, limpar_chave: !!f.limparChave };
    }

    const api = { PRESETS_RESERVA, corpoReserva, linhasNdjson, tabelaParaServidor, descreverItem, aplicarPlano, resumoAplicacao, textoSeguro };

    // ── DOM ───────────────────────────────────────────────────────────────────────────────────────────
    if (typeof document !== 'undefined') {
        const mk = (tag, cls, texto) => { const e = document.createElement(tag); if (cls) e.className = cls; if (texto != null) e.textContent = texto; return e; };

        /** O cartão do plano dentro do chat. `acoes`: {aplicar(itens) → resumo | Promise, desfazer(), obra(item) → Promise<msg>}. */
        api.renderPlano = function (plano, acoes) {
            const card = mk('div', 'chat-plano');
            const itens = plano.itens || [];
            const topo = mk('div', 'cp-topo');
            topo.appendChild(mk('b', '', itens.some(i => i.acao === 'obra') ? 'Obra proposta' : 'Plano proposto'));
            if (!itens.some(i => i.acao === 'obra')) topo.appendChild(mk('span', 'cp-qtd', `${itens.length} alteração(ões)`));
            card.appendChild(topo);
            if (plano.resumo) card.appendChild(mk('div', 'cp-resumo', plano.resumo));
            const ul = mk('ul', 'cp-lista');
            itens.forEach(it => {
                const d = descreverItem(it);
                const li = mk('li', 'cp-' + d.classe);
                li.appendChild(mk('span', 'cp-simb', d.simbolo));
                if (d.tabela) li.appendChild(mk('span', 'cp-tab', d.tabela));
                if (d.operacao) li.appendChild(mk('span', 'cp-op', d.operacao));
                li.appendChild(mk('code', '', d.texto));
                if (d.antes) { const a = mk('s', 'cp-antes', d.antes); a.title = 'era assim'; li.appendChild(a); }
                (d.avisos || []).forEach(av => li.appendChild(mk('div', 'cp-aviso', '⚠ ' + av)));
                ul.appendChild(li);
            });
            card.appendChild(ul);
            const barra = mk('div', 'cp-acoes');
            const btnOk = mk('button', 'btn-primary', itens.some(i => i.acao === 'obra') ? 'Continuar…' : 'Aplicar');
            btnOk.type = 'button';
            const btnNao = mk('button', 'btn-secondary', 'Descartar');
            btnNao.type = 'button';
            barra.append(btnOk, btnNao);
            if (acoes.previa && itens.every(i => i.acao !== 'obra')) {
                const btnPrevia = mk('button', 'btn-secondary', 'Simular orçamento');
                btnPrevia.type = 'button';
                const saida = mk('div', 'cp-resumo');
                saida.setAttribute('role', 'status');
                barra.appendChild(btnPrevia);
                card.appendChild(saida);
                btnPrevia.addEventListener('click', async () => {
                    btnPrevia.disabled = true;
                    saida.textContent = 'Calculando prévia…';
                    try {
                        const r = await acoes.previa(itens);
                        const titulo = r.total_mudancas ? `${r.total_mudancas} item(ns) de orçamento com quantidade alterada` : 'Nenhuma quantidade mudou no orçamento.';
                        saida.replaceChildren(mk('b', '', titulo));
                        const mudancas = mk('ul', 'cp-lista');
                        r.mudancas.forEach(i => mudancas.appendChild(mk('li', '', `${i.operacao} · ${i.codigo} · ${i.mdo} · ${i.antes} → ${i.depois}`)));
                        saida.appendChild(mudancas);
                        if (r.total_mudancas > r.mudancas.length) saida.appendChild(mk('p', '', 'Exibindo os primeiros 20 itens.'));
                        if (r.nao_encontrados.length) {
                            const detalhe = mk('details', 'cp-aviso');
                            detalhe.append(mk('summary', '', `${r.nao_encontrados.length} item(ns) ausente(s) da base`),
                                mk('p', '', r.nao_encontrados.join(', ')));
                            saida.appendChild(detalhe);
                        }
                        r.avisos.forEach(av => saida.appendChild(mk('p', 'cp-aviso', typeof av === 'string' ? av : (av.mensagem || JSON.stringify(av)))));
                    } catch (e) {
                        saida.textContent = e.message;
                    } finally { btnPrevia.disabled = false; }
                });
            }
            card.appendChild(barra);
            const fim = (msg, desfazer) => {
                barra.textContent = '';
                barra.appendChild(mk('span', 'cp-estado', msg));
                if (desfazer) {
                    const d = mk('button', 'btn-secondary', 'Desfazer');
                    d.type = 'button';
                    d.addEventListener('click', () => fim(acoes.desfazer(), false));
                    barra.appendChild(d);
                }
            };
            btnNao.addEventListener('click', () => fim('Descartado.', false));
            btnOk.addEventListener('click', async () => {
                btnOk.disabled = btnNao.disabled = true;
                try {
                    if (itens.some(i => i.acao === 'obra')) fim(await acoes.obra(itens[0]), false);
                    else {
                        const r = await acoes.aplicar(itens);
                        fim(r.mensagem, r.aplicado);
                    }
                } catch (e) { fim('Erro ao aplicar: ' + e.message, false); }
            });
            return card;
        };


        /** Formulário do provedor reserva automático (dentro da «Configuração de IA»). Devolve {recarregar}. */
        api.reservaUI = function (el, opcoes) {
            const toast = (opcoes && opcoes.toast) || (m => alert(m));
            el.textContent = '';
            const caixa = mk('div', 'rs-caixa');
            caixa.appendChild(mk('p', 'rs-titulo', 'PROVEDOR RESERVA (AUTOMÁTICO)'));
            caixa.appendChild(mk('p', 'rs-ajuda', 'Se o provedor principal falhar (sobrecarga, cota, chave, rede), o assistente continua sozinho neste, e avisa no chat. DeepSeek, Qwen, OpenRouter e Ollama entram como «OpenAI-compatível».'));
            const lAtiva = mk('label', 'rs-ativa');
            const ativa = mk('input'); ativa.type = 'checkbox';
            lAtiva.append(ativa, document.createTextNode(' Usar provedor reserva'));
            const preset = mk('select', 'rs-campo');
            const o0 = mk('option', '', 'Escolher um atalho…'); o0.value = ''; preset.appendChild(o0);
            PRESETS_RESERVA.forEach(p => { const o = mk('option', '', p.rotulo); o.value = p.id; preset.appendChild(o); });
            const provider = mk('select', 'rs-campo');
            [['openai', 'OpenAI-compatível (DeepSeek, Qwen, OpenRouter, Ollama…)'], ['gemini', 'Gemini'], ['claude', 'Claude']].forEach(([v, t]) => { const o = mk('option', '', t); o.value = v; provider.appendChild(o); });
            const url = mk('input', 'rs-campo'); url.placeholder = 'URL base (ex.: https://api.deepseek.com)'; url.autocomplete = 'off';
            const modelo = mk('input', 'rs-campo'); modelo.placeholder = 'Nome do modelo (ex.: deepseek-flash)'; modelo.autocomplete = 'off';
            const chave = mk('input', 'rs-campo'); chave.type = 'password'; chave.autocomplete = 'new-password'; chave.placeholder = 'Chave da API do reserva';
            const estadoChave = mk('p', 'rs-ajuda');
            const dica = mk('p', 'rs-dica');
            const barra = mk('div', 'rs-barra');
            const salvar = mk('button', 'btn-primary', 'Salvar reserva'); salvar.type = 'button';
            const testar = mk('button', 'btn-secondary', 'Testar'); testar.type = 'button';
            const estado = mk('span', 'rs-estado');
            barra.append(salvar, testar, estado);
            caixa.append(lAtiva, preset, provider, url, modelo, chave, estadoChave, dica, barra);
            el.appendChild(caixa);
            let temChave = false, pode = true;
            const ajustarProvider = () => { url.style.display = provider.value === 'openai' ? '' : 'none'; };
            provider.addEventListener('change', ajustarProvider);
            preset.addEventListener('change', () => {
                const p = PRESETS_RESERVA.find(x => x.id === preset.value);
                if (!p) return;
                provider.value = p.provider; url.value = p.url; if (p.modelo) modelo.value = p.modelo; else if (p.provider === 'openai') modelo.value = '';
                dica.textContent = p.dica || ''; ajustarProvider();
            });
            const marcarEstado = (txt, ruim) => { estado.textContent = txt; estado.style.color = ruim ? '#ff7b72' : '#3fb950'; };
            async function recarregar(manterMensagem) {
                try {
                    const r = await fetch('/api/assistente/reserva');
                    if (!r.ok) return;
                    const j = await r.json();
                    ativa.checked = !!j.ativa; provider.value = j.provider; url.value = j.base_url; modelo.value = j.model; temChave = !!j.tem_chave; pode = !!j.pode_editar;
                    chave.value = '';
                    estadoChave.textContent = temChave ? 'Chave salva ✓. Informe a chave correspondente ao trocar de provedor ou URL.' : 'Nenhuma chave salva.';
                    [ativa, preset, provider, url, modelo, chave, salvar, testar].forEach(c => { c.disabled = !pode; });
                    if (!pode) marcarEstado('Só o administrador configura o reserva.', true); else if (!manterMensagem) estado.textContent = '';
                    ajustarProvider();
                } catch (e) { /* sem rede: mantém o formulário como está */ }
            }
            async function gravar() {
                const r = await fetch('/api/assistente/reserva', { method: 'PUT', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(corpoReserva({ ativa: ativa.checked, provider: provider.value, url: url.value, modelo: modelo.value, chave: chave.value })) });
                const j = await r.json().catch(() => ({}));
                if (!r.ok) throw new Error(j.detail || 'HTTP ' + r.status);
                return j;
            }
            salvar.addEventListener('click', async () => {
                salvar.disabled = true;
                try { await gravar(); marcarEstado('Reserva salvo.', false); toast('Provedor reserva salvo.'); await recarregar(true); }
                catch (e) { marcarEstado(e.message, true); } finally { salvar.disabled = !pode; }
            });
            testar.addEventListener('click', async () => {
                testar.disabled = true; marcarEstado('Salvando e testando…', false);
                try {
                    await gravar();
                    const r = await fetch('/api/assistente/reserva/testar', { method: 'POST' });
                    const j = await r.json().catch(() => ({}));
                    if (j.ok) marcarEstado(`Funcionou: ${j.modelo} concluiu a consulta e respondeu em ${(j.ms / 1000).toFixed(1)} s.`, false); else marcarEstado(j.erro || ('HTTP ' + r.status), true);
                    await recarregar(true);
                } catch (e) { marcarEstado(e.message, true); } finally { testar.disabled = !pode; }
            });
            ajustarProvider();
            recarregar();
            return { recarregar };
        };

        /** Janela de Skills: lista à esquerda, editor à direita. `opcoes.projeto()` = código do projeto; `opcoes.toast(msg)`. */
        api.abrirSkills = async function (opcoes) {
            const projeto = opcoes.projeto();
            const toast = opcoes.toast || (m => alert(m));
            let dados;
            try {
                const r = await fetch('/api/assistente/skills?projeto=' + encodeURIComponent(projeto));
                if (!r.ok) throw new Error('HTTP ' + r.status);
                dados = await r.json();
            } catch (e) { toast('Não consegui listar as skills: ' + e.message); return; }
            const pode = !!dados.pode_editar;
            const fundo = mk('div', 'modal-overlay');
            const caixa = mk('div', 'modal-content sk-caixa');
            fundo.appendChild(caixa);
            caixa.appendChild(mk('h3', '', 'Skills do assistente' + (dados.projeto ? ' — ' + dados.projeto : '')));
            caixa.appendChild(mk('p', 'sk-ajuda', 'Guias que o assistente consulta antes de agir. A skill do projeto vale mais que a geral. ' + (pode ? '' : 'Só o administrador edita.')));
            const corpo = mk('div', 'sk-corpo');
            const lista = mk('div', 'sk-lista');
            const editor = mk('div', 'sk-editor');
            corpo.append(lista, editor);
            caixa.appendChild(corpo);
            const rodape = mk('div', 'sk-rodape');
            const btnFechar = mk('button', 'btn-secondary', 'Fechar');
            btnFechar.type = 'button';
            btnFechar.addEventListener('click', () => fundo.remove());
            fundo.addEventListener('mousedown', e => { if (e.target === fundo) fundo.remove(); });
            if (pode) {
                const btnNovo = mk('button', 'btn-secondary', '＋ Nova skill');
                btnNovo.type = 'button';
                const btnCat = mk('button', 'btn-secondary', '⟳ Gerar catálogo de ativos');
                btnCat.type = 'button';
                btnCat.title = 'Lê a Base de Orçamento do projeto e grava a skill «ativos» (famílias e códigos válidos)';
                btnNovo.addEventListener('click', () => abrir(null));
                btnCat.addEventListener('click', async () => {
                    btnCat.disabled = true;
                    try {
                        const r = await fetch('/api/assistente/skills/gerar-catalogo?projeto=' + encodeURIComponent(projeto), { method: 'POST' });
                        const j = await r.json().catch(() => ({}));
                        if (!r.ok) throw new Error(j.detail || 'HTTP ' + r.status);
                        toast(`Catálogo gerado: ${j.ativos} ativos.`);
                        await recarregar();
                        abrir({ escopo: 'projeto', nome: 'ativos' });
                    } catch (e) { toast('Não gerei o catálogo: ' + e.message); } finally { btnCat.disabled = false; }
                });
                rodape.append(btnNovo, btnCat);
            }
            rodape.appendChild(btnFechar);
            caixa.appendChild(rodape);
            document.body.appendChild(fundo);

            const escopoDe = s => (s.escopo === 'projeto' ? dados.escopo_projeto : '_geral');
            async function recarregar() {
                const r = await fetch('/api/assistente/skills?projeto=' + encodeURIComponent(projeto));
                if (r.ok) dados = await r.json();
                desenharLista();
            }
            function desenharLista() {
                lista.textContent = '';
                if (!dados.skills.length) lista.appendChild(mk('div', 'sk-vazio', 'Nenhuma skill ainda.'));
                dados.skills.forEach(s => {
                    const b = mk('button', 'sk-item');
                    b.type = 'button';
                    b.appendChild(mk('b', '', s.nome));
                    b.appendChild(mk('span', 'sk-tag', (s.escopo === 'projeto' ? 'projeto' : 'geral') + (s.origem === 'usuario' ? ' · sua' : '')));
                    b.appendChild(mk('small', '', [s.descricao, s.versao && 'Revisão ' + s.versao].filter(Boolean).join(' · ')));
                    b.addEventListener('click', () => abrir(s));
                    lista.appendChild(b);
                });
            }
            async function abrir(s) {
                editor.textContent = '';
                const nome = mk('input', 'sk-nome');
                nome.placeholder = 'nome_da_skill (minúsculas, números, _ e -)';
                const escopo = mk('select', 'sk-escopo');
                [...(dados.escopo_projeto ? [['projeto', 'Deste projeto']] : []), ['geral', 'Todos os projetos']].forEach(([v, t]) => { const o = mk('option', '', t); o.value = v; escopo.appendChild(o); });
                const area = mk('textarea', 'sk-texto');
                area.spellcheck = false;
                area.placeholder = 'descricao: quando usar esta skill\n# Título\nRegras e passos, em linhas curtas...';
                if (s) {
                    nome.value = s.nome; nome.disabled = true; escopo.value = s.escopo; escopo.disabled = true;
                    try {
                        const r = await fetch(`/api/assistente/skills/${encodeURIComponent(s.pasta || escopoDe(s))}/${encodeURIComponent(s.nome)}`);
                        area.value = r.ok ? (await r.json()).texto : '';
                    } catch (e) { area.value = ''; }
                }
                area.readOnly = !pode;
                editor.append(nome, escopo, area);
                if (pode) {
                    const barra = mk('div', 'sk-barra');
                    const salvar = mk('button', 'btn-primary', 'Salvar');
                    salvar.type = 'button';
                    salvar.addEventListener('click', async () => {
                        const alvo = { escopo: escopo.value, nome: nome.value.trim() };
                        try {
                            const r = await fetch(`/api/assistente/skills/${encodeURIComponent(escopoDe(alvo))}/${encodeURIComponent(alvo.nome)}`, {
                                method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ texto: area.value }),
                            });
                            const j = await r.json().catch(() => ({}));
                            if (!r.ok) throw new Error(j.detail || 'HTTP ' + r.status);
                            toast('Skill salva.');
                            await recarregar();
                            abrir(dados.skills.find(x => x.nome === alvo.nome) || null);
                        } catch (e) { toast('Não salvei: ' + e.message); }
                    });
                    barra.appendChild(salvar);
                    if (s && s.origem === 'usuario') {
                        const apagar = mk('button', 'btn-secondary', 'Apagar a minha');
                        apagar.type = 'button';
                        apagar.addEventListener('click', async () => {
                            if (!confirm(`Apagar a skill «${s.nome}»? Se houver uma versão que veio com o programa, ela volta a valer.`)) return;
                            const r = await fetch(`/api/assistente/skills/${encodeURIComponent(s.pasta || escopoDe(s))}/${encodeURIComponent(s.nome)}`, { method: 'DELETE' });
                            if (!r.ok) { toast('Não apaguei (' + r.status + ').'); return; }
                            toast('Skill apagada.');
                            editor.textContent = '';
                            await recarregar();
                        });
                        barra.appendChild(apagar);
                    }
                    editor.appendChild(barra);
                } else if (s) {
                    editor.appendChild(mk('div', 'sk-vazio', 'Somente leitura.'));
                }
                if (!s) nome.focus();
            }
            desenharLista();
            if (pode && opcoes.rascunho && dados.escopo_projeto) {
                await abrir(null);
                const area = editor.querySelector('textarea');
                area.value = opcoes.rascunho;
            } else if (dados.skills.length) abrir(dados.skills[0]);
        };
    }

    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.Assistente = api;
})(typeof window !== 'undefined' ? window : globalThis);
