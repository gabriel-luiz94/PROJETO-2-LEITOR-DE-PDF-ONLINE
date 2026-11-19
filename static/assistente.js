/*
 * assistente.js (TASK-191) — o lado do navegador do assistente de IA com ferramentas e skills.
 *   · funções PURAS (testadas em Node, tests/test_assistente_front_191.py): ler o stream NDJSON, montar a cópia das tabelas que vai ao servidor,
 *     descrever e APLICAR um plano de alterações com checagem de que a linha ainda é a mesma;
 *   · DOM (só no navegador): o cartão do plano (Aplicar / Descartar / Desfazer) e o editor de Skills (listar, editar, criar, apagar, gerar o catálogo).
 * O servidor nunca altera a tabela: ele PROPÕE; quem aplica é o clique do usuário aqui.
 */
(function (raiz) {
    'use strict';

    const CAMPOS_LINHA = ['entidade', 'operacao', 'ativo', 'qtdAtivos', 'congelada', 'vinculoEstruturas', 'estruturasFicticias'];
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
            for (let i = 0; i < dados.length && saida[t].length < LIMITE_LINHAS; i++) {
                const r = dados[i];
                if (!r || typeof r !== 'object') continue;
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
            texto: it.ativo || '', antes: it.acao === 'editar' && it.antes ? it.antes.ativo : '', avisos: it.avisos || [],
        };
    }

    /**
     * Aplica o plano em `estados` ({cabos: {data}, outros: {data}}). `entidadeDe(ativo, tabela)` dá a entidade das linhas novas/editadas.
     * Editar/remover só vale se a linha da posição `id` ainda tem o mesmo texto de antes (o usuário pode ter mexido na tela depois do pedido);
     * senão a alteração é IGNORADA e listada. Devolve {adicionadas, editadas, removidas, ignoradas: [{n, motivo}]}.
     */
    function aplicarPlano(estados, itens, entidadeDe) {
        const r = { adicionadas: 0, editadas: 0, removidas: 0, ignoradas: [] };
        const ent = (a, t) => (typeof entidadeDe === 'function' ? entidadeDe(a, t) : (t === 'cabos' ? 'CABO' : '0'));
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
            if (it.antes && String(atual.ativo || '').trim() !== String(it.antes.ativo || '').trim()) { r.ignoradas.push({ n: it.n, motivo: 'a linha mudou depois do pedido' }); return; }
            if (it.acao === 'remover') { dados[it.id] = null; r.removidas++; return; }
            atual.operacao = it.operacao || atual.operacao || 'I';
            atual.ativo = it.ativo;
            atual.entidade = ent(it.ativo, it.tabela);
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

    const api = { linhasNdjson, tabelaParaServidor, descreverItem, aplicarPlano, resumoAplicacao, textoSeguro };

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
            card.appendChild(barra);
            const fim = (msg, desfazer) => {
                barra.textContent = '';
                barra.appendChild(mk('span', 'cp-estado', msg));
                if (desfazer) {
                    const d = mk('button', 'btn-secondary', 'Desfazer');
                    d.type = 'button';
                    d.addEventListener('click', () => { acoes.desfazer(); barra.textContent = ''; barra.appendChild(mk('span', 'cp-estado', 'Desfeito.')); });
                    barra.appendChild(d);
                }
            };
            btnNao.addEventListener('click', () => fim('Descartado.', false));
            btnOk.addEventListener('click', async () => {
                btnOk.disabled = btnNao.disabled = true;
                try {
                    if (itens.some(i => i.acao === 'obra')) fim(await acoes.obra(itens[0]), false);
                    else fim('Aplicado: ' + await acoes.aplicar(itens) + '. Ctrl+Z desfaz.', true);
                } catch (e) { fim('Erro ao aplicar: ' + e.message, false); }
            });
            return card;
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
                    b.appendChild(mk('small', '', s.descricao || ''));
                    b.addEventListener('click', () => abrir(s));
                    lista.appendChild(b);
                });
            }
            async function abrir(s) {
                editor.textContent = '';
                const nome = mk('input', 'sk-nome');
                nome.placeholder = 'nome_da_skill (minúsculas, números, _ e -)';
                const escopo = mk('select', 'sk-escopo');
                [['projeto', 'Deste projeto'], ['geral', 'Todos os projetos']].forEach(([v, t]) => { const o = mk('option', '', t); o.value = v; escopo.appendChild(o); });
                const area = mk('textarea', 'sk-texto');
                area.spellcheck = false;
                area.placeholder = 'descricao: quando usar esta skill\n# Título\nRegras e passos, em linhas curtas...';
                if (s) {
                    nome.value = s.nome; nome.disabled = true; escopo.value = s.escopo; escopo.disabled = true;
                    try {
                        const r = await fetch(`/api/assistente/skills/${encodeURIComponent(escopoDe(s))}/${encodeURIComponent(s.nome)}`);
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
                            const r = await fetch(`/api/assistente/skills/${encodeURIComponent(escopoDe(s))}/${encodeURIComponent(s.nome)}`, { method: 'DELETE' });
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
            if (dados.skills.length) abrir(dados.skills[0]);
        };
    }

    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.Assistente = api;
})(typeof window !== 'undefined' ? window : globalThis);
