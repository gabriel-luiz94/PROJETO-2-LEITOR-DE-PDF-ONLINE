/* ═══════════════════════════════════════════════════════════════════════════
   Gaveta "Regras de validação" do Resumo do projeto (TASK-019).
   Monta a gaveta, as abas (Regras · Grupos · Prompts da IA · Testar) e carrega sob demanda o editor
   (regras_editor.js) e os prompts (painel_prompts.js). Segue o projeto de #select-projeto; o padrão de todos os
   projetos é uma opção dentro do painel. Só admin edita (o backend também exige); operador vê Regras/Grupos em leitura.
   Conteúdo sempre via textContent/value (nunca innerHTML).
═══════════════════════════════════════════════════════════════════════════ */
(function () {
    const ehAdmin = () => localStorage.getItem('is_admin') === 'true';
    const SCRIPTS = ['/static/regras_editor.js', '/static/painel_ajustes.js', '/static/painel_prompts.js'];
    let montado = false, carregando = null;

    // Mensagem curta usada pelos editores (toast do Resumo).
    window.rpMensagem = function (tipo, msg) {
        if (typeof window.showToast === 'function') window.showToast((tipo === 'error' ? 'Erro: ' : '') + msg);
        else alert(msg);
    };

    function no(tag, cls, texto) {
        const e = document.createElement(tag);
        if (cls) e.className = cls;
        if (texto !== undefined) e.textContent = texto;
        return e;
    }
    function botao(id, texto, cls) {
        const b = no('button', cls || 'rp-btn', texto);
        b.type = 'button'; if (id) b.id = id;
        return b;
    }
    function campo(rotulo, el) {
        const l = no('label', 'block text-xs text-[#8b949e]', rotulo);
        el.classList.add('block', 'mt-1');
        l.appendChild(el);
        return l;
    }
    const ADM = 'rp-admin';   // elementos só do admin (escondidos para o operador)

    // Aba Execução (TASK-026): o que o botão Validar roda e o "Validar ao montar". Preferências LOCAIS de qualquer usuário
    // (localStorage, via window.validacaoPrefs de resumo.js) — não dependem de a gaveta estar aberta para valerem.
    function montarPainelExecucao() {
        const p = no('div', 'rp-painel ativo'); p.id = 'rp-execucao';
        p.appendChild(no('p', 'rp-ajuda', 'Escolha o que o botão Validar roda e se a validação acontece ao montar o orçamento. Vale só neste navegador.'));
        const prefs = () => window.validacaoPrefs;
        const caixa = (id, rotulo, dica) => {
            const l = no('label', 'flex items-center gap-2 text-sm'); l.title = dica;
            const c = document.createElement('input'); c.type = 'checkbox'; c.id = id;
            l.append(c, rotulo);
            return l;
        };
        const grupo = no('div', 'space-y-2');
        grupo.append(
            caixa('vmodo-det', 'Determinística — contrato das planilhas e regras', 'Formato, operação e base técnica; barato e sem IA'),
            caixa('vmodo-dominio', 'Regras de domínio do projeto (as que o admin ligou)', 'Só vale com a determinística ligada'),
            caixa('vmodo-ia', 'IA — revisão por IA', 'Usa a chave de IA'),
            caixa('vmodo-pular-ia', 'Pular a IA se o contrato tiver erro', 'Economiza a chave: a IA só roda depois de o contrato estar sem erro'));
        const auto = no('div', 'mt-3 flex items-center gap-3 text-sm');
        auto.appendChild(no('span', '', 'Validar ao montar o orçamento:'));
        [['nao', 'Não'], ['sim', 'Sim']].forEach(([v, r]) => {
            const l = no('label', 'inline-flex items-center gap-1');
            const i = document.createElement('input'); i.type = 'radio'; i.name = 'validacao_auto'; i.value = v;
            l.append(i, r);
            auto.appendChild(l);
        });
        const efeito = no('div', 'rp-leitura-aviso'); efeito.id = 'vmodo-efeito'; efeito.style.marginTop = '14px'; efeito.style.color = '#8b949e';
        p.append(grupo, auto, efeito);

        const mapa = { det: 'vmodo-det', dominio: 'vmodo-dominio', ia: 'vmodo-ia', pularIa: 'vmodo-pular-ia' };
        function sincronizar() {
            if (!prefs()) return;
            const m = prefs().lerModos();
            Object.entries(mapa).forEach(([k, id]) => { p.querySelector('#' + id).checked = !!m[k]; });
            p.querySelector('#vmodo-dominio').disabled = !m.det;
            p.querySelector('#vmodo-pular-ia').disabled = !(m.det && m.ia);
            const ligada = prefs().validacaoAutomaticaLigada();
            p.querySelectorAll('input[name="validacao_auto"]').forEach(r => { r.checked = r.value === (ligada ? 'sim' : 'nao'); });
            efeito.textContent = 'O botão Validar roda: ' + prefs().resumoDoQueRoda().longo;
        }
        Object.entries(mapa).forEach(([k, id]) => p.querySelector('#' + id).addEventListener('change', e => {
            const m = prefs().lerModos(); m[k] = e.target.checked; prefs().gravarModos(m); sincronizar();
        }));
        p.querySelectorAll('input[name="validacao_auto"]').forEach(r => r.addEventListener('change', () => {
            prefs().definirValidacaoAutomatica(r.value === 'sim'); sincronizar();
        }));
        p.addEventListener('rp-mostrar', sincronizar);
        sincronizar();
        return p;
    }

    function montarPainelRegras() {
        const p = no('div', 'rp-painel'); p.id = 'rp-regras';
        p.appendChild(no('p', 'rp-ajuda', 'O padrão vale para todos os projetos. O projeto pode ajustar uma regra, ocultá-la (−, reexibe com +) ou criar regras só dele. Quando o padrão muda, o projeto acompanha, exceto onde ajustou.'));
        const status = no('div', 'text-sm text-[#8b949e] mb-2'); status.id = 'rdStatus';
        const erros = no('div', 'rp-caixa-erro hidden'); erros.id = 'rdErros';
        const avisos = no('div', 'rp-caixa-aviso hidden'); avisos.id = 'rdAvisos';
        const lista = no('div', 'space-y-2'); lista.id = 'rdLista';
        const assistente = no('div', 'hidden mt-3 border rounded-md p-3'); assistente.id = 'rdAssistente';
        const dl = document.createElement('datalist'); dl.id = 'rdSugestoes';
        const barra = no('div', 'flex flex-wrap gap-2 mt-3 ' + ADM);
        barra.append(botao('rdAdicionar', 'Adicionar regra'), botao('rdSalvar', 'Salvar', 'btn-primary'),
            botao('rdNovas', 'Adicionar regras novas da semente'), botao('rdSemente', 'Restaurar semente'), botao('rdHistorico', 'Histórico'));
        barra.querySelector('#rdNovas').title = 'Só no padrão: acrescenta as regras da semente que ainda não estão na lista, desligadas';
        const hist = no('div', 'rp-hist hidden'); hist.id = 'rdHistLista';
        p.append(status, erros, avisos, lista, assistente, dl, barra, hist);
        return p;
    }

    function montarPainelGrupos() {
        const p = no('div', 'rp-painel'); p.id = 'rp-grupos';
        p.appendChild(no('p', 'rp-ajuda', 'Um grupo é uma lista nomeada de ativos (ex.: @ESTRUTURA_MT) usada nas regras. Mudar o grupo muda todas as regras que o usam; num projeto, redefinir um grupo vale só para ele. Salve pela aba Regras.'));
        const g = no('div', 'space-y-2'); g.id = 'rdGrupos';
        const novo = botao('rdNovoGrupo', 'Novo grupo', 'rp-btn mt-3 ' + ADM);
        p.append(g, novo);
        return p;
    }

    function montarPainelAjustes() {
        const p = no('div', 'rp-painel'); p.id = 'rp-ajustes';
        p.appendChild(no('p', 'rp-ajuda', 'Ajustes são ações recorrentes sobre as tabelas Cabos e Outros (substituir texto, ordenar, excluir ou adicionar linhas…). Cada um pode ser pré-visualizado nas tabelas atuais; nada é aplicado sem você ver o que muda. A edição fica em rascunho até clicar em Salvar (grava local e no Supabase).'));
        const status = no('div', 'text-sm text-[#8b949e] mb-2'); status.id = 'ajStatus';
        const erros = no('div', 'rp-caixa-erro hidden'); erros.id = 'ajErros';
        const avisos = no('div', 'rp-caixa-aviso hidden'); avisos.id = 'ajAvisos';
        const lista = no('div', 'space-y-2'); lista.id = 'ajLista';
        const dl = document.createElement('datalist'); dl.id = 'ajRegrasSugestoes';
        const barra = no('div', 'flex flex-wrap gap-2 mt-3 ' + ADM);
        barra.append(botao('ajAdicionar', 'Novo ajuste'), botao('ajSalvar', 'Salvar', 'btn-primary'),
            botao('ajSemente', 'Restaurar semente'), botao('ajHistorico', 'Histórico'));
        const hist = no('div', 'rp-hist hidden'); hist.id = 'ajHistLista';
        p.append(status, erros, avisos, lista, dl, barra, hist);
        return p;
    }

    function montarPainelPrompts() {
        const p = no('div', 'rp-painel ' + ADM); p.id = 'rp-prompts';
        p.appendChild(no('p', 'rp-ajuda', 'Instruções enviadas à IA para checar e corrigir as planilhas. Cada projeto usa a versão própria, se existir; senão herda o padrão. Cada salvamento guarda a versão anterior no histórico.'));
        const linha = no('div', 'rp-linha');
        const sel = no('select', 'px-3 py-1'); sel.id = 'pvPrompt';
        const st = no('div', 'text-sm text-[#8b949e]'); st.id = 'pvStatus';
        linha.append(campo('Prompt', sel), st);
        const erros = no('div', 'rp-caixa-erro hidden'); erros.id = 'pvErros';
        const ta = no('textarea', 'w-full px-3 py-2 font-mono text-xs'); ta.id = 'pvTexto'; ta.rows = 22; ta.spellcheck = false;
        const barra = no('div', 'flex flex-wrap gap-2 mt-3');
        barra.append(botao('pvSalvar', 'Salvar', 'btn-primary'), botao('pvSemente', 'Restaurar semente'), botao('pvHistorico', 'Histórico'));
        const hist = no('div', 'rp-hist hidden'); hist.id = 'pvHistLista';
        p.append(linha, erros, ta, barra, hist);
        return p;
    }

    function montarPainelTestar() {
        const p = no('div', 'rp-painel ' + ADM); p.id = 'rp-testar';
        p.appendChild(no('p', 'rp-ajuda', 'Roda as regras em edição (sem salvar) contra linhas de exemplo, uma por linha no formato "OPERAÇÃO ativo", ex.: "I DT11/300 1-CFU" em Outros, "I P50 ABC 6 m" em Cabos. Regras desligadas ou ocultas não são avaliadas.'));
        const grade = no('div', 'grid grid-cols-1 md:grid-cols-3 gap-3');
        const cabos = no('textarea', 'w-full px-3 py-2 font-mono text-xs'); cabos.id = 'rdTesteCabos'; cabos.rows = 5; cabos.spellcheck = false;
        const outros = no('textarea', 'w-full px-3 py-2 font-mono text-xs'); outros.id = 'rdTesteOutros'; outros.rows = 5; outros.spellcheck = false;
        const lc = campo('Cabos', cabos), lo = campo('Outros', outros);
        lo.classList.add('md:col-span-2');
        grade.append(lc, lo);
        const acoes = no('div', 'mt-3 flex flex-wrap items-center gap-3');
        const todas = document.createElement('input'); todas.type = 'checkbox'; todas.id = 'rdTesteTodas';
        const lt = no('label', 'text-sm text-[#8b949e]'); lt.append(todas, ' Mostrar também por que as outras linhas não dispararam');
        acoes.append(botao('rdTestar', 'Testar rascunho'), lt);
        const res = no('div', 'mt-3 text-sm'); res.id = 'rdTesteResultado';
        p.append(grade, acoes, res);
        return p;
    }

    function projetoDoResumo() {
        const sel = document.getElementById('select-projeto');
        const o = sel && sel.selectedIndex >= 0 ? sel.options[sel.selectedIndex] : null;
        return o && o.dataset.codigo ? { codigo: o.dataset.codigo, rotulo: o.textContent } : null;
    }

    function preencherProjetos() {
        const sel = document.getElementById('rdProjeto');
        if (!sel) return;
        const anterior = sel.value;
        const proj = projetoDoResumo();
        sel.replaceChildren();
        if (proj) sel.appendChild(new Option(`Projeto selecionado: ${proj.rotulo}`, proj.codigo));
        sel.appendChild(new Option('Padrão de todos os projetos', 'DEFAULT'));
        if (anterior === 'DEFAULT') sel.value = 'DEFAULT';
    }

    async function recarregarAbas() {
        if (typeof rdCarregar === 'function') await rdCarregar();
        if (typeof ajCarregar === 'function') await ajCarregar();   // depois das regras: o vínculo sugere ids de regras
        if (ehAdmin() && typeof pvCarregarLista === 'function') pvCarregarLista();
    }

    function ativarAba(nome) {
        document.querySelectorAll('#painel-regras .rp-aba').forEach(b => b.classList.toggle('ativa', b.dataset.aba === nome));
        document.querySelectorAll('#painel-regras .rp-painel').forEach(p => p.classList.toggle('ativo', p.id === 'rp-' + nome));
        const aviso = document.getElementById('rp-aviso-leitura');   // a aba Execução é editável por todos
        if (aviso) aviso.style.display = nome === 'execucao' ? 'none' : '';
        const alvo = document.getElementById('rp-' + nome);
        if (alvo) alvo.dispatchEvent(new Event('rp-mostrar'));   // a aba Execução relê as preferências ao ser mostrada
    }

    function montar() {
        if (montado) return;
        montado = true;
        if (!document.getElementById('painel-regras-css')) {
            const l = document.createElement('link');
            l.rel = 'stylesheet'; l.href = '/static/painel_regras.css?v=2'; l.id = 'painel-regras-css';
            document.head.appendChild(l);
        }
        const d = no('aside', ''); d.id = 'painel-regras'; d.setAttribute('aria-label', 'Regras de validação');
        const topo = no('div', 'rp-topo');
        const sel = no('select', 'px-3 py-1'); sel.id = 'rdProjeto';
        topo.append(no('h2', '', 'Regras de validação'), campo('Editando', sel));
        const fechar = botao('', 'Fechar', 'rp-btn'); fechar.style.marginLeft = 'auto'; fechar.onclick = fecharPainel;
        topo.appendChild(fechar);
        const abas = no('div', 'rp-abas');
        [['execucao', 'Execução', false], ['regras', 'Regras', false], ['grupos', 'Grupos', false], ['ajustes', 'Ajustes', false], ['prompts', 'Prompts da IA', true], ['testar', 'Testar', true]].forEach(([k, r, adm]) => {
            const b = no('button', 'rp-aba' + (k === 'execucao' ? ' ativa' : '') + (adm ? ' ' + ADM : ''), r);
            b.type = 'button'; b.dataset.aba = k; b.onclick = () => ativarAba(k);
            abas.appendChild(b);
        });
        const corpo = no('div', 'rp-corpo');
        if (!ehAdmin()) { const av = no('div', 'rp-leitura-aviso', 'Somente leitura: só o administrador edita regras, grupos e ajustes.'); av.id = 'rp-aviso-leitura'; av.style.display = 'none'; corpo.appendChild(av); }
        corpo.append(montarPainelExecucao(), montarPainelRegras(), montarPainelGrupos(), montarPainelAjustes(), montarPainelPrompts(), montarPainelTestar());
        d.append(topo, abas, corpo);
        document.body.appendChild(d);
        if (!ehAdmin()) d.querySelectorAll('.' + ADM).forEach(e => e.classList.add('hidden'));
        preencherProjetos();
        sel.addEventListener('change', recarregarAbas);
    }

    function carregarScript(src) {
        return new Promise((ok, falha) => {
            const s = document.createElement('script');
            s.src = src; s.onload = ok; s.onerror = () => falha(new Error('Falha ao carregar ' + src));
            document.head.appendChild(s);
        });
    }

    async function carregarEditores() {
        if (carregando) return carregando;
        carregando = (async () => {
            for (const src of SCRIPTS) {
                if (src.includes('prompts') && !ehAdmin()) continue;   // prompts: só admin
                await carregarScript(src);
            }
            await iniciarRegrasDominio();
            await iniciarAjustes();
            if (ehAdmin()) await iniciarPromptsValidacao();
        })();
        return carregando;
    }

    async function abrirPainel(aba) {
        montar();
        if (aba) ativarAba(aba);
        preencherProjetos();
        document.getElementById('painel-regras').classList.add('aberto');
        try {
            await carregarEditores();
            await recarregarAbas();
        } catch (e) {
            window.rpMensagem('error', e.message);
        }
    }
    // "Cadastrar como ajuste" (proposta da IA aceita no Resumo): abre a aba Ajustes com um ajuste novo em rascunho.
    // Atalho ⚙ ao lado do Validar: abre a gaveta direto numa aba (ex.: 'execucao').
    window.rpAbrirNaAba = function (aba) {
        const d = document.getElementById('painel-regras');
        if (d && d.classList.contains('aberto')) { ativarAba(aba); return Promise.resolve(); }   // já aberta: só troca de aba (não perde rascunho)
        return abrirPainel(aba);
    };

    window.rpCadastrarComoAjuste = async function (acoes, nome) {
        await abrirPainel();
        ativarAba('ajustes');
        if (typeof ajNovoComAcoes === 'function') ajNovoComAcoes(acoes, nome);
        window.rpMensagem('success', 'Ajuste criado em rascunho: revise, ligue e clique em Salvar.');
    };

    function fecharPainel() {
        const d = document.getElementById('painel-regras');
        if (d) d.classList.remove('aberto');
    }

    document.addEventListener('DOMContentLoaded', () => {
        const btn = document.getElementById('btn-regras-validacao');
        if (btn) btn.addEventListener('click', () => {
            const d = document.getElementById('painel-regras');
            if (d && d.classList.contains('aberto')) fecharPainel(); else abrirPainel();
        });
        // O painel segue o projeto do Resumo.
        const sel = document.getElementById('select-projeto');
        if (sel) sel.addEventListener('change', () => {
            if (!montado) return;
            preencherProjetos();
            if (document.getElementById('painel-regras').classList.contains('aberto') && carregando) recarregarAbas();
        });
        document.addEventListener('keydown', e => { if (e.key === 'Escape') fecharPainel(); });
    });
})();
