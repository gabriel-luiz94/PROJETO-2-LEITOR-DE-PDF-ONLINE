/**
 * resumo.js  –  Aba "Resumo do Orçamento"
 * Tabelas de Cabos e Demais Ativos lado a lado.
 * Colunas: Entidade (select), Operação (select validado), Ativo (input + autocomplete).
 * Funcionalidades: Add linha, Excluir linha, Copiar, Undo (Ctrl+Z), Redo (Ctrl+Y).
 */
document.addEventListener('DOMContentLoaded', () => {
    // TASK-241: Resumo aberto na janela nativa do Mapa («Abrir em nova janela»): mostra só o Mapa, com uma cópia da obra sincronizada
    // com a janela principal (ela grava o projeto; esta cópia nunca grava os dados da obra no navegador).
    const MODO_JANELA_MAPA = new URLSearchParams(location.search).has('janela_transferida');
    window.__modoJanelaMapa = MODO_JANELA_MAPA;
    if (MODO_JANELA_MAPA) document.body.classList.add('janela-mapa');

    /* ── Helpers ── */
    function deepClone(obj) { return JSON.parse(JSON.stringify(obj)); }

    /* ═══════════════════════════════════════
       CONECTIVIDADE COM A NUVEM (TASK-040)
       Escuta o evento {"type":"connectivity"} do WebSocket /ws (emitido por
       services/connectivity_monitor.py) e mostra um aviso só quando offline —
       sem ruído quando tudo está normal.
    ═══════════════════════════════════════ */
    (function conectividade() {
        const selo = document.getElementById('indicador-conectividade');
        const bola = document.getElementById('indicador-conectividade-bola');
        const texto = document.getElementById('indicador-conectividade-texto');
        if (!selo || !bola || !texto) return;

        // TASK-121: no desktop o evento traz a data da cópia local ("2026-11-09 10:30:00"): o aviso diz com que cópia está trabalhando
        const dataCurta = c => { const m = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}:\d{2})/.exec(c || ''); return m ? `${m[3]}/${m[2]} ${m[4]}` : ''; };
        function atualizar(online, extra) {
            if (online) {
                selo.style.display = 'none';
                return;
            }
            selo.style.display = 'inline-flex';
            selo.style.background = 'rgba(248,81,73,0.15)';
            selo.style.color = '#f85149';
            bola.style.background = '#f85149';
            const cp = dataCurta(extra && extra.copia);
            texto.textContent = cp ? `Sem conexão com a nuvem — modo local (cópia de ${cp})` : 'Sem conexão com a nuvem — trabalhando localmente';
        }
        // o aviso também aparece se a tela abrir quando o programa já está offline (o evento do WebSocket só vem na troca de estado)
        fetch('/api/rede/estado').then(r => (r.ok ? r.json() : null)).then(d => { if (d && d.ativa && d.online === false) atualizar(false, d.copia ? { copia: d.copia.ultima } : null); }).catch(() => {});

        try {
            const wsProtocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
            const socket = new WebSocket(`${wsProtocol}//${window.location.host}/ws`);
            socket.onmessage = (event) => {
                try {
                    const data = JSON.parse(event.data);
                    if (data.type === 'connectivity') atualizar(data.status === 'online', data);
                } catch (e) { /* mensagem não era JSON de conectividade: ignora */ }
            };
            socket.onerror = () => { /* sem WebSocket: o indicador só não aparece, nada quebra */ };
        } catch (e) { /* idem */ }
    })();

    /* ═══════════════════════════════════════
       DADOS GLOBAIS
    ═══════════════════════════════════════ */
    const ENTIDADES = ['0', 'CABO', 'CHAVE', 'TRAFO', 'ESTRUTURA', 'APOIO', 'IP', 'POSTE', 'RAMAIS', 'CERCA'];
    const OPERACOES = ['I', '*I', 'R', '*R', 'M', '*M'];
    // Estado de desempenho das tabelas (declarado aqui no topo porque renderTable já roda na montagem inicial, antes das seções abaixo):
    const _medidasTabela = { cabos: null, outros: null };      // altura da área/cabeçalho/linha, para contar as linhas em branco sem forçar layout
    const _idsObj = new WeakMap(); let _seqId = 0;             // identidade de objetos (base, configuração) para as impressões digitais dos caches
    let _cacheCores = null, _tmCoresRender = 0, _tmResumoRede = 0, _cacheAssin = null;
    // Tabelas grandes (TASK-185): acima de LIMIAR_LAZY linhas, só as linhas perto da área visível são montadas na hora; as demais entram como linha vazia de altura fixa (_lazy)
    // e são montadas por pedaços em segundo plano (hidratação), ao rolar, ou antes de qualquer ação do usuário que leia as células. Resultado: criar/excluir/duplicar/desfazer
    // com 400+ linhas custa o desenho de ~50 linhas, não de todas.
    // Estado da seleção em intervalo da grade: declarado AQUI, antes do 1º renderTable (que pode repintar a seleção ao montar linhas) — `const` mais abaixo dava «Cannot access 'grade' before initialization» ao abrir a tela com muitas linhas
    const grade = { tipo: null, ancora: null, foco: null, arrastando: false, ultimaLinha: null };
    const LIMIAR_LAZY = 60, LIMIAR_PARCIAL = 300;
    const _lazy = { cabos: [], outros: [] }, _lazyTm = { cabos: 0, outros: 0 }, _hLinhaReal = { cabos: 0, outros: 0 };                // cores por linha (verde/laranja/vermelho) e o agendamento depois do desenho
    /* TASK-182 C — Operação digitada, validada como uma célula do Excel: maiúsculas e sem espaços (*r → *R); só aceita OPERACOES (fonte única); inválida = não grava,
       a célula fica vermelha com o aviso e o valor volta ao anterior. O campo tem a lista (datalist) para quem prefere escolher. */
    const MSG_OPERACAO_INVALIDA = 'Operação inválida: use I, *I, R, *R, M ou *M';
    const normalizarOperacao = t => String(t == null ? '' : t).replace(/\s+/g, '').toUpperCase();
    // TASK-235: F = fictício (só nas linhas de Cabos e Demais Ativos, não na Totalizadora): a linha fica congelada (fora do orçamento) e vale só para o vínculo
    const OPERACOES_LINHA = OPERACOES.concat(['F']);
    const DETALHE_OPERACAO = { 'F': 'fictício (fora do orçamento, só para o vínculo)', 'I': 'instalar', '*I': 'instalar em linha viva', 'R': 'remover', '*R': 'remover em linha viva', 'M': 'manter', '*M': 'manter em linha viva' };
    const ehCampoOperacao = el => !!(el && el.matches && el.matches('input.inp-operacao, input.tot-op-input'));
    /* Lista de operações no estilo do VS Code (static/ide_lista.js): ao entrar no campo aparecem TODAS (a atual selecionada); digitar filtra; ↑↓ navegam; Enter/Tab escolhem; Esc fecha. */
    let _listaOp = null;
    /** TASK-235: troca a operação da linha. F congela a linha (fora do orçamento); sair de F descongela o que o F congelou. */
    function aplicarOperacaoNaLinha(row, v) {
        const eraF = String(row.operacao || '').trim().toUpperCase() === 'F';
        row.operacao = v;
        if (v === 'F') { if (!row.congelada) row._congeladaPorF = true; row.congelada = true; }
        else if (eraF && row._congeladaPorF) { delete row.congelada; delete row._congeladaPorF; }
        if (typeof aplicarCoresJa === 'function') aplicarCoresJa();          // TASK-247: R/*R saem das cores na hora
    }
    function listaOperacao() {
        return _listaOp || (_listaOp = window.IdeLista.criar({ aoEscolher: (item, campo) => { campo.value = item.valor; campo.dispatchEvent(new Event('change', { bubbles: true })); } }));
    }
    function abrirListaOperacao(inp, filtrar) {
        if (!window.IdeLista || inp.readOnly || inp.disabled) return;
        const q = filtrar ? normalizarOperacao(inp.value) : '';
        const itens = (inp.dataset.aceitaF ? OPERACOES_LINHA : OPERACOES).filter(v => !q || v.startsWith(q)).map(v => ({ valor: v, detalhe: DETALHE_OPERACAO[v] }));
        const atual = normalizarOperacao(inp.dataset.valorOk || inp.value);
        listaOperacao().abrir(inp, itens, q ? (itens.length ? 0 : -1) : itens.findIndex(x => x.valor === atual));
    }
    document.addEventListener('focusin', e => { if (ehCampoOperacao(e.target) && !e.target.readOnly) { e.target.select(); abrirListaOperacao(e.target, false); } });
    document.addEventListener('click', e => { if (ehCampoOperacao(e.target) && window.IdeLista && !listaOperacao().aberta()) abrirListaOperacao(e.target, false); });
    document.addEventListener('input', e => { if (ehCampoOperacao(e.target)) abrirListaOperacao(e.target, true); });
    document.addEventListener('keydown', e => {
        if (!ehCampoOperacao(e.target) || !window.IdeLista) return;
        const l = listaOperacao();
        if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
            e.preventDefault(); e.stopPropagation();
            if (!l.aberta()) abrirListaOperacao(e.target, false); else l.mover(e.key === 'ArrowDown' ? 1 : -1);
        } else if ((e.key === 'Enter' || e.key === 'Tab') && l.aberta() && l.indice() >= 0 && !e.shiftKey) {
            e.preventDefault(); e.stopPropagation();
            l.escolherAtual();
        } else if (e.key === 'Enter' || e.key === 'Escape') l.fechar();
        if ((e.key === 'Enter' || e.key === 'Tab') && !e.shiftKey && !e.ctrlKey && !e.altKey && !e.metaKey) {                  // Enter/Tab na operação (escolhida ou digitada) → vai para o ativo da mesma linha
            const tr = e.target.closest('tr'), alvo = tr && (tr.querySelector('.inp-ativo') || tr.querySelector('.inp-ativo-vazio') || tr.querySelector('.tot-ativo-input'));
            if (alvo) { e.preventDefault(); e.stopPropagation(); alvo.focus(); const f = alvo.value.length; try { alvo.setSelectionRange(f, f); } catch (x) { /* sem seleção */ } }
        }
    }, true);
    // ← no começo do ativo volta para a Operação da mesma linha (no meio do texto, a seta continua movendo o cursor)
    document.addEventListener('keydown', e => {
        if (e.key !== 'ArrowLeft' || e.shiftKey || e.ctrlKey || e.metaKey || e.altKey || e.isComposing) return;
        const t = e.target;
        if (!(t && t.matches && t.matches('#body-cabos .inp-ativo, #body-outros .inp-ativo, #body-cabos-vazio .inp-ativo-vazio, #body-outros-vazio .inp-ativo-vazio'))) return;
        if (t.selectionStart !== 0 || t.selectionEnd !== 0) return;
        const op = t.closest('tr') && t.closest('tr').querySelector('input.inp-operacao');
        if (!op || op.readOnly) return;
        e.preventDefault(); e.stopPropagation();
        op.focus();                                                                    // o focusin do campo seleciona tudo e abre a lista de operações
    }, true);
    document.addEventListener('focusout', e => { if (ehCampoOperacao(e.target)) setTimeout(() => { if (window.IdeLista && document.activeElement !== listaOperacao().doCampo()) listaOperacao().fechar(); }, 120); });
    /** Marca a célula como inválida por alguns segundos (vermelho + aviso no título) e devolve o valor anterior. */
    function recusarOperacao(inp, anterior) {
        inp.value = anterior;
        inp.classList.add('op-invalida');
        const tituloAntes = inp.dataset.tituloOp || inp.title || '';
        inp.dataset.tituloOp = tituloAntes;
        inp.title = MSG_OPERACAO_INVALIDA;
        showToast(MSG_OPERACAO_INVALIDA);
        clearTimeout(inp._tmOpInvalida);
        inp._tmOpInvalida = setTimeout(() => { inp.classList.remove('op-invalida'); inp.title = tituloAntes; }, 3000);
    }
    /** Campo de operação (texto + lista). `aoMudar(v)` só roda com valor VÁLIDO e diferente do atual. */
    function criarCampoOperacao(valor, classe, aoMudar) {
        const inp = document.createElement('input');
        inp.type = 'text'; inp.className = classe; inp.value = valor; inp.dataset.valorOk = valor;
        inp.autocomplete = 'off'; inp.spellcheck = false; inp.maxLength = 4;
        inp.setAttribute('aria-label', 'Operação: I, *I, R, *R, M ou *M');
        inp.title = 'Operação: I, *I, R, *R, M ou *M — digite ou escolha na lista';
        inp.addEventListener('change', () => {
            const v = normalizarOperacao(inp.value), ok = inp.dataset.valorOk;
            if (v === ok) { inp.value = ok; return; }
            if (!(inp.dataset.aceitaF ? OPERACOES_LINHA : OPERACOES).includes(v)) { recusarOperacao(inp, ok); return; }
            inp.value = v; inp.dataset.valorOk = v; inp.classList.remove('op-invalida');
            aoMudar(v);
        });
        inp.addEventListener('keydown', e => { if (e.key === 'Escape') { inp.value = inp.dataset.valorOk; inp.blur(); } });
        return inp;
    }

    let extractedDataCache = [];

    // ── Regras do Leitor (TASK-006) — reusadas aqui via a Tabela de Classificação ──────────────
    window.__regrasLeitorProcessamento = [];
    window.__regrasLeitorClassificacao = [];
    window.carregarRegrasLeitorResumo = async function (projetoCodigo) {
        if (!projetoCodigo) return;
        try {
            const [resProc, resCls] = await Promise.all([
                fetch(`/api/regras-leitor/processamento?projeto_codigo=${encodeURIComponent(projetoCodigo)}`),
                fetch(`/api/regras-leitor/classificacao?projeto_codigo=${encodeURIComponent(projetoCodigo)}`)
            ]);
            const dataProc = resProc.ok ? await resProc.json() : { regras: [] };
            const dataCls = resCls.ok ? await resCls.json() : { regras: [] };
            window.__regrasLeitorProcessamento = dataProc.regras || [];
            window.__regrasLeitorClassificacao = dataCls.regras || [];
        } catch (e) {
            console.error('Erro ao carregar regras do leitor:', e);
        }
    };

    /** Linha CONGELADA: fica na tela (e pode ser editada), mas sai de Totalizadora, orçamento, validação, ajustes, cores, contadores e algoritmos — serve de comentário. Só um vínculo feito à mão pode usar um cabo congelado como fonte.
     *  Linha BLOQUEADA: não pode ser editada nem excluída; continua entrando nos cálculos (a menos que também esteja congelada). */
    const congelada = r => !!(r && r.congelada), bloqueada = r => !!(r && r.bloqueada);
    /** TASK-182: texto do ativo começando por "#" (depois de espaços) = comentário. `sincronizarHash(linha, textoAnterior)`: começou com "#" → vira comentário (congelada);
     *  tinha "#" e deixou de ter → volta ao normal (só se ainda é comentário). Linha bloqueada não muda. Devolve true se mudou a linha. */
    const comecaComHash = t => /^\s*#/.test(String(t == null ? '' : t));
    function sincronizarHash(row, anterior) {
        if (!row || row.bloqueada) return false;
        let mudou = false;
        const GL = window.GruposLinhas;
        if (GL) {                                                 // TASK-195: «@» no começo = declaração de grupo (congelada, como o comentário)
            const g = GL.sincronizar(row, anterior);
            if (g) { mudou = true; if (g.tipo === 'deixou') ['outros', 'cabos'].forEach(t => GL.liberar(tableStates[t].data, g.gid)); }
        }
        if (comecaComHash(row.ativo)) {
            if (row.comentario) return mudou;
            row.comentario = true; row.congelada = true; return true;
        }
        if (comecaComHash(anterior) && row.comentario) { delete row.comentario; delete row.congelada; return true; }
        return mudou;
    }
    /** Marca o projeto como alterado a cada ação do usuário que entra no histórico (congelar/bloquear/comentar pelo menu não disparam input/change). Não vale ao abrir/carregar um projeto. */
    let _historicoSemSujo = false;
    function marcarProjetoAlterado() {
        if (!_historicoSemSujo && window.GridRProjeto && window.GridRProjeto.marcarSujo) window.GridRProjeto.marcarSujo(false);
        window.VinculacaoUI?.atualizarObra();
    }
    const tableStates = {
        cabos: { bodyId: 'body-cabos', data: [] },
        outros: { bodyId: 'body-outros', data: [] },
        totalizadora: { bodyId: 'body-totalizadora', data: [] },
        regras: { bodyId: 'body-regras', data: [] }
    };

    /* ─── Column filters ─── */
    // Selections: type → field → Set of selected values
    const filterSelections = {
        cabos:  { operacao: new Set(), ativo: new Set() },
        outros: { entidade: new Set(), operacao: new Set(), ativo: new Set() }
    };
    const FILTER_FIELDS_POR_TIPO = { cabos: ['operacao', 'ativo'], outros: ['entidade', 'operacao', 'ativo'] };   // Cabos não tem coluna Entidade (TASK-064); Outros voltou a ter (TASK-069)
    const camposFiltro = type => FILTER_FIELDS_POR_TIPO[type];
    // TASK-069: estado de classificação por coluna (critérios aplicados) e agrupamento visual por tabela
    const ordemAtual = { cabos: [], outros: [] };
    const agrupamento = { cabos: null, outros: null };
    const FILTER_CONTAINERS = {
        cabos:  { operacao: 'rfc-cabos-operacao', ativo: 'rfc-cabos-ativo' },
        outros: { entidade: 'rfc-outros-entidade', operacao: 'rfc-outros-operacao', ativo: 'rfc-outros-ativo' }
    };

    /* ─── Undo / Redo history ─── */
    const MAX_HISTORY = 80;
    let history = [];      // array de snapshots
    let historyIdx = -1;   // ponteiro atual
    let historyDirty = false;  // pode haver mudança ainda não gravada no histórico (ver pushHistory)

    /* ─── Autocomplete state ─── */
    let acList = null;          // elemento DOM do dropdown ativo
    let acInput = null;         // input ativo
    let acItems = [];           // itens filtrados
    let acSelected = -1;        // índice selecionado
    const ativoSets = { cabos: new Set(), outros: new Set() };

    /* ─── Context menu state ─── */
    const ctxMenu = document.getElementById('ctx-menu-resumo');
    let ctxRow = null, ctxType = null, ctxFoco = null;          // ctxFoco: a célula que tinha o foco quando o menu abriu (volta para ela ao fechar)

    /* ═══════════════════════════════════════
       INICIALIZAÇÃO
    ═══════════════════════════════════════ */
    let dataRaw = localStorage.getItem('processar_dados');
    if (!dataRaw) {
        dataRaw = '[]';
    }

    extractedDataCache = JSON.parse(dataRaw);

    // Pré-processa lógica de negócio para cada item
    const allProcessed = extractedDataCache.map(item => {
        const result = computeRowLogic(item);
        return result;
    });

    tableStates.cabos.data = allProcessed
        .filter(r => r.entidade === 'CABO')
        .map(deepClone);

    tableStates.outros.data = allProcessed
        .filter(r => r.entidade !== 'CABO' && r.entidade !== '0' && r.entidade !== 'RAMAIS')
        .map(deepClone);

    // Popula dados para o modal RAMAIS (RAMAIS, IP e APOIO condicional)
    window._ramaisData = allProcessed
        .filter(r => {
            if (r.entidade === 'RAMAIS' || r.entidade === 'IP') return true;
            if (r.entidade === 'APOIO') {
                const txt = ((r._raw && r._raw.texto) || r.ativo || '').toUpperCase();
                return /REC.*CAL[CÇ]ADA/i.test(txt) || /CONC.*BASE/i.test(txt);
            }
            return false;
        })
        .map(r => ({
            entidade: r.entidade,
            texto: (r._raw && r._raw.texto) || r.ativo || '',
            _textoOriginal: (r._raw && r._raw.texto) || '',
            _raw: r._raw || null,
            pagina: r._raw && r._raw.pagina != null ? r._raw.pagina : '-'
        }));

    // Garante que ao menos os campos necessários existam
    ['cabos', 'outros'].forEach(type => {
        tableStates[type].data.forEach(r => {
            r.entidade = r.entidade || '0';
            r.operacao = r.operacao || 'M';
            r.ativo    = r.ativo    || '';
        });
    });

    // TASK-059 (aprendizado, trabalho manual): foto das tabelas como saíram do Leitor, antes de qualquer edição. Vale enquanto ninguém carregar,
    // adicionar ou subtrair uma obra (aí as tabelas deixam de ser "o que o programa entregou"). Só os campos necessários, sem o desenho.
    let sessaoManual = (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(36).slice(2);
    let baselineManual = extractedDataCache.length ? {
        cabos: tableStates.cabos.data.map(r => ({ entidade: r.entidade, operacao: r.operacao, ativo: r.ativo })),
        outros: tableStates.outros.data.map(r => ({ entidade: r.entidade, operacao: r.operacao, ativo: r.ativo }))
    } : null;

    buildAtivoSets();
    recalcAllQtdAtivos();  // calcula qtdAtivos antes do primeiro render
    pushHistory();   // estado inicial
    historyDirty = false;  // nada a desfazer ainda
    updateHistoryUI();

    renderTable('cabos');
    renderTable('outros');
    updateCounters();
    updateHistoryUI();
    buildDataLists();
    refreshAllFilters('cabos');
    refreshAllFilters('outros');

    // Fecha dropdowns de filtro ao clicar fora
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.r-filter-container')) {
            const hadActive = document.querySelectorAll('.r-filter-dropdown.active').length > 0;
            document.querySelectorAll('.r-filter-dropdown.active').forEach(d => d.classList.remove('active'));
            if (hadActive) { applyFilters('cabos'); applyFilters('outros'); }
        }
    });

    // isGray/processAtivoFormula foram substituidas pelo motor de regras do leitor
    // (TASK-006) -- ver static/regras_leitor_engine.js.

    // computeRowLogic foi religada ao motor de regras do leitor (TASK-006) -- a cascata embutida
    // antiga foi removida (ver static/regras_leitor_engine.js e window.__regrasLeitorProcessamento/
    // Classificacao, carregadas por projeto em carregarRegrasLeitorResumo()).
    function computeRowLogic(item) {
        // Se a entidade e ativo já vieram definidos da aba principal, confiar neles diretamente
        // (evita re-derivação que descartaria a classificação manual do usuário)
        if (item.entidade && item.entidade !== '0' && item.ativo && item.ativo.trim() !== '') {
            return {
                entidade: item.entidade,
                operacao: item.operacao || 'M',
                ativo: item.ativo,
                _raw: item
            };
        }

        const globalIdx = extractedDataCache.indexOf(item);
        const engineItem = {
            texto: item.texto || '',
            cor: item.cor || '#000000',
            layer: item.layer || '',
            pagina: item.pagina,
            index: globalIdx >= 0 ? globalIdx : 0,
            allItems: extractedDataCache
        };
        const resultado = RegrasLeitorEngine.processarEClassificar(
            engineItem,
            window.__regrasLeitorProcessamento || [],
            window.__regrasLeitorClassificacao || []
        );
        return { entidade: resultado.entidade, operacao: resultado.operacao, ativo: resultado.ativo, _raw: item };
    }

    /* ═══════════════════════════════════════
       RENDER TABLE
    ═══════════════════════════════════════ */
    /** Select com só a opção atual no DOM; a lista completa entra ao focar/clicar (milhares de linhas ficam bem mais leves). */
    function preencherSelectPreguicoso(sel, valores, atual) {
        const uma = document.createElement('option');
        uma.value = atual; uma.textContent = atual; uma.selected = true;
        sel.appendChild(uma);
        const completar = () => {
            if (sel.dataset.completo) return;
            sel.dataset.completo = '1';
            const agora = sel.value;
            sel.innerHTML = '';
            valores.forEach(v => {
                const opt = document.createElement('option');
                opt.value = v; opt.textContent = v;
                if (v === agora) opt.selected = true;
                sel.appendChild(opt);
            });
        };
        sel.definirValor = v => {          // troca o valor sem exigir a lista completa
            if (sel.dataset.completo) { sel.value = v; return; }
            sel.innerHTML = '';
            const o = document.createElement('option');
            o.value = v; o.textContent = v; o.selected = true;
            sel.appendChild(o);
        };
        sel.addEventListener('focus', completar);
        sel.addEventListener('mousedown', completar);
        sel.addEventListener('keydown', completar);
    }

    /* ── Hidratação das linhas (TASK-185) ── */
    function hidratarLinha(tr) {
        const m = tr && tr._montar;
        if (!m) return false;
        tr._montar = null;
        tr.classList.remove('lazy-row'); tr.style.height = ''; tr.textContent = '';
        m();
        return true;
    }
    function _terminouHidratar() {
        agendarCoresAposDesenhar();
        if (typeof intervaloGrade === 'function' && intervaloGrade()) pintarGrade();           // seleção em intervalo: as linhas recém-montadas entram pintadas
    }
    /** Monta as linhas que estão (ou estarão logo) na área visível da tabela. */
    function hidratarJanela(type, folga) {
        const pend = _lazy[type];
        if (!pend.length) return 0;
        const body = document.getElementById(tableStates[type].bodyId), wrap = body && body.closest('.table-wrap');
        let alvo;
        if (wrap && wrap.clientHeight > 0 && wrap.offsetParent !== null) {
            const r = wrap.getBoundingClientRect(), y0 = r.top - folga, y1 = r.bottom + folga;
            alvo = pend.filter(tr => { const q = tr.getBoundingClientRect(); return q.height > 0 && q.bottom >= y0 && q.top <= y1; });
        } else alvo = pend.slice(0, 40);                    // área escondida (outra aba/painel): só o começo; o resto vem ao mostrar/rolar
        let n = 0;
        alvo.forEach(tr => { if (hidratarLinha(tr)) n++; });
        if (n) { _lazy[type] = pend.filter(tr => tr._montar); _terminouHidratar(); }
        return n;
    }
    function hidratarTudo(type) {
        const pend = _lazy[type];
        if (!pend.length) return 0;
        _lazy[type] = [];
        if (_lazyTm[type]) { clearTimeout(_lazyTm[type]); _lazyTm[type] = 0; }
        let n = 0;
        pend.forEach(tr => { if (hidratarLinha(tr)) n++; });
        if (n) _terminouHidratar();
        return n;
    }
    function agendarHidratacao(type) {
        if (_lazyTm[type] || !_lazy[type].length) return;
        _lazyTm[type] = setTimeout(function passo() {
            _lazyTm[type] = 0;
            if (tableStates[type].data.length > LIMIAR_PARCIAL) { hidratarJanela(type, 2500); return; }          // tabelas muito grandes: só ~2500 px em volta da área visível; o resto monta ao rolar
            const pend = _lazy[type];
            const t0 = performance.now();
            let n = 0;
            while (pend.length && performance.now() - t0 < 7) { if (hidratarLinha(pend.shift())) n++; }
            if (pend.length) _lazyTm[type] = setTimeout(passo, 16);
            else if (n) _terminouHidratar();
        }, 120);
    }
    // Só colar/copiar/cortar/arrastar (e Ctrl+A) montam tudo antes. Clique, Enter, excluir, duplicar, congelar, desfazer etc. trabalham sobre os DADOS e sobre as classes da linha
    // (`row-selected`), que existem também nas linhas ainda vazias; a seleção em intervalo é repintada quando as linhas são montadas (`_terminouHidratar`). Assim, uma
    // sequência rápida de ações não paga a montagem de todas as linhas a cada vez.
    (function ligarGatilhosHidratacao() {
        const todas = () => { hidratarTudo('cabos'); hidratarTudo('outros'); };
        ['copy', 'cut', 'paste', 'dragstart'].forEach(ev => document.addEventListener(ev, () => { if (_lazy.cabos.length || _lazy.outros.length) todas(); }, true));
        document.addEventListener('keydown', e => {
            if ((e.ctrlKey || e.metaKey) && String(e.key).toLowerCase() === 'a' && (_lazy.cabos.length || _lazy.outros.length)) todas();
        }, true);
        let raf = 0;
        document.addEventListener('scroll', e => {
            if (!(e.target && e.target.classList && e.target.classList.contains('table-wrap')) || raf) return;
            raf = requestAnimationFrame(() => { raf = 0; hidratarJanela('cabos', 500); hidratarJanela('outros', 500); });
        }, { capture: true, passive: true });
    })();

    function contarColunasVisiveis(tabela) {
        const cabecalho = tabela.tHead.rows[0];
        const ocultas = tabela.classList.contains('sem-entidade') ? '.col-entidade'
            : tabela.classList.contains('sem-cabo-extra') ? '.col-cabo-extra' : null;
        return cabecalho.cells.length - (ocultas ? cabecalho.querySelectorAll(ocultas).length : 0);
    }

    function ajustarCelulasDeLarguraTotal(tabela) {
        const colunasVisiveis = contarColunasVisiveis(tabela);
        tabela.querySelectorAll('.lazy-row > td, .grupo-sep > td').forEach(td => {
            td.colSpan = colunasVisiveis;
        });
    }

    function renderTable(type) {
        window.VinculacaoUI?.atualizarObra();
        if (MODO_JANELA_MAPA && (type === 'cabos' || type === 'outros')) window.__agendarEnvioMapa?.();
        const state = tableStates[type];
        const body = document.getElementById(state.bodyId);
        body.innerHTML = '';
        _lazy[type] = []; if (_lazyTm[type]) { clearTimeout(_lazyTm[type]); _lazyTm[type] = 0; }
        const lazy = state.data.length > LIMIAR_LAZY, hLazy = (_medidasTabela[type] && _medidasTabela[type].linha) || _hLinhaReal[type] || 30;
        const colunasVisiveis = contarColunasVisiveis(body.closest('table'));
        const frag = document.createDocumentFragment();   // uma única inserção no DOM (tabelas grandes)
        const avulsoVelho = type === 'outros' ? _indicesAvulsoVelho() : new Set();      // TASK-171

        state.data.forEach((row, idx) => {
            const tr = document.createElement('tr');
            tr.dataset.index = idx;
            tr.dataset.type = type;
            let btnDel = null;
            const montar = () => {

            /* ── Nº da linha (seleciona a linha) ── */
            const tdNum = document.createElement('td');
            tdNum.className = 'col-num';
            tdNum.textContent = idx + 1;
            if (row.comentario) { const m = document.createElement('span'); m.className = 'marca-linha'; m.textContent = '💬'; m.title = 'Comentário: anotação da tabela, fora de todos os cálculos (botão direito para tirar o comentário)'; tdNum.appendChild(m); }
            else if (row.congelada && !row.grupoDecl) { const m = document.createElement('span'); m.className = 'marca-linha'; m.textContent = '❄'; m.title = 'Linha congelada: fora da Totalizadora, do orçamento e da validação (botão direito para descongelar)'; tdNum.appendChild(m); }
            if (row.grupoDecl) tdNum.appendChild(criarBotaoGrupo(type, row));            // TASK-195: +/− do grupo e quantas linhas ele tem
            if (row.bloqueada) { const m = document.createElement('span'); m.className = 'marca-linha'; m.textContent = '🔒'; m.title = 'Linha bloqueada: não pode ser editada nem excluída (botão direito para desbloquear)'; tdNum.appendChild(m); }
            tr.appendChild(tdNum);

            /* ── Entidade (só Outros — TASK-069): texto leve; clicar abre o seletor. Escolher à mão trava a linha (entidadeManual) ── */
            let atualizarCelulaEntidade = () => {};
            if (type === 'outros') {
                const tdEnt = document.createElement('td');
                tdEnt.className = 'cel-ent';                     // sem data-col: fora da grade (copiar/colar/intervalo), só o seletor
                const view = document.createElement('div');
                view.className = 'ent-view';
                view.tabIndex = 0;
                const lockEnt = document.createElement('span');
                lockEnt.className = 'ent-lock';
                atualizarCelulaEntidade = () => {
                    view.textContent = row.entidade || '0';
                    lockEnt.textContent = row.entidadeManual ? '🔒' : '';
                    view.title = row.entidadeManual ? 'Definida manualmente (travada). Clique para trocar; botão direito → Definir entidade… → Automática destrava.' : 'Automática: acompanha o ativo. Clique para escolher à mão.';
                };
                atualizarCelulaEntidade();
                const abrirSeletor = () => {
                    if (tdEnt.querySelector('select') || bloqueada(row)) return;
                    const sel = document.createElement('select');
                    sel.className = 'sel-entidade';
                    sel.dataset.field = 'entidade';
                    const lista = ENTIDADES.filter(e => e !== 'CABO');
                    if (row.entidade && !lista.includes(row.entidade)) lista.push(row.entidade);
                    lista.forEach(e => {
                        const opt = document.createElement('option');
                        opt.value = e; opt.textContent = e;
                        if (e === (row.entidade || '0')) opt.selected = true;
                        sel.appendChild(opt);
                    });
                    let fechado = false;
                    const fechar = () => { if (fechado) return; fechado = true; if (sel.parentNode) sel.remove(); view.style.display = ''; atualizarCelulaEntidade(); };
                    sel.addEventListener('change', () => {
                        if (sel.value !== row.entidade) {
                            pushHistory();                          // snapshot ANTES da mudança
                            row.entidade = sel.value;
                            row.entidadeManual = true;              // escolha manual: não é recalculada ao editar o ativo
                            refreshAllFilters(type);
                            atualizarResumoRedeUI();
                        }
                        fechar();
                    });
                    sel.addEventListener('blur', fechar);
                    sel.addEventListener('keydown', e => { if (e.key === 'Escape') { e.stopPropagation(); fechar(); } });
                    view.style.display = 'none';
                    tdEnt.insertBefore(sel, view);
                    sel.focus();
                    try { if (sel.showPicker) sel.showPicker(); } catch (err) { /* sem showPicker: o usuário abre com um clique */ }
                };
                view.addEventListener('click', abrirSeletor);
                view.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); abrirSeletor(); } });
                tdEnt.appendChild(view);
                tdEnt.appendChild(lockEnt);
                tr.appendChild(tdEnt);
            }

            /* ── Operação ── */
            const tdOp = document.createElement('td');
            const selOp = criarCampoOperacao(row.operacao, 'sel-operacao inp-operacao', v => {
                pushHistory();  // snapshot ANTES da mudança
                aplicarOperacaoNaLinha(row, v);
                refreshAllFilters(type);
                atualizarResumoRedeUI();
            });
            selOp.dataset.field = 'operacao';
            if (!row.grupoDecl && !row.comentario) selOp.dataset.aceitaF = '1';
            if (type === 'outros') selOp.addEventListener('keydown', e => { if (e.key === 'Tab' && e.shiftKey && !e.ctrlKey && !e.altKey && !e.metaKey && idx > 0) { e.preventDefault(); irParaCelula(type, idx - 1, 1, -1); } });
            tdOp.appendChild(selOp);
            tdOp.dataset.col = '0';
            tr.appendChild(tdOp);

            /* ── Ativo ── */
            const tdAt = document.createElement('td');
            tdAt.style.position = 'relative';
            const inpAt = document.createElement('input');
            inpAt.type = 'text';
            inpAt.className = 'inp-ativo';
            inpAt.dataset.field = 'ativo';
            // «Ocultar parâmetros» ligado (só Outros): o campo mostra e edita só os NOMES ("12600D, N1, N4"); row.ativo guarda o texto cheio e os parâmetros ficam presos a cada item
            const oculto = type === 'outros' && !!window.AtivoItens && ocultarParametrosLigado();
            let itensDoCampo = oculto ? window.AtivoItens.itens(row.ativo) : null, compactoDoCampo = '';
            inpAt.value = oculto ? window.AtivoItens.compacto(row.ativo) : row.ativo;
            compactoDoCampo = inpAt.value;
            const textoDoAtivo = () => (oculto ? row.ativo : inpAt.value);
            inpAt._sincronizarCheio = () => { if (oculto) { itensDoCampo = window.AtivoItens.itens(row.ativo); compactoDoCampo = inpAt.value; } else inpAt.value = row.ativo; };
            inpAt.autocomplete = 'off';
            inpAt.setAttribute('list', '');  // desabilita datalist nativo, usamos o nosso
            if (type === 'cabos') { inpAt.dataset.parTipo = 'cabos'; inpAt.dataset.parIdx = idx; }
            if (type === 'outros') { inpAt.dataset.parTipo = 'outros'; inpAt.dataset.parIdx = idx; if (avulsoVelho.has(idx)) { inpAt.classList.add('avulso-velho'); inpAt.title = 'avulso desatualizado: o cabo ou a estrutura fictícia mudou (veja Vinculação)'; } }        // TASK-167: o balão (parâmetros preenchidos e os que faltam) lê estes dados ao passar o mouse
            let prevAtivo = row.ativo;

            inpAt.addEventListener('input', () => {
                const oldVal = inpAt.value;
                const upperVal = oldVal.toUpperCase();
                if (oldVal !== upperVal) {
                    const cursor = inpAt.selectionStart;
                    inpAt.value = upperVal;
                    if (cursor !== null) inpAt.setSelectionRange(cursor, cursor);
                }
                if (oculto) {
                    const r = window.AtivoItens.aplicarEdicao(itensDoCampo, compactoDoCampo, inpAt.value, inpAt.selectionStart);
                    itensDoCampo = r.itens; row.ativo = r.texto;
                    if (inpAt.value.indexOf('(') >= 0) { inpAt.value = window.AtivoItens.compacto(r.texto); const f = inpAt.value.length; inpAt.setSelectionRange(f, f); }      // colou com parênteses: os parâmetros vão para o item e o campo continua só com os nomes
                    compactoDoCampo = inpAt.value;
                } else row.ativo = inpAt.value;
                agendarCoresParametros();
                if (type !== 'outros') showAutocomplete(inpAt, type);                // TASK-136: em Outros quem sugere é AtivoSugestao (ativos/apelidos do projeto, com botão liga/desliga)
            });
            inpAt.addEventListener('keydown', e => {
                if (type === 'outros' && window.AtivoBalao) {
                    if (oculto && e.key === '(' && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); window.AtivoBalao.abrirNoCursor(inpAt); return; }       // "(" = atalho para os parâmetros do ativo
                    if (e.key === 'Enter' && e.ctrlKey) { e.preventDefault(); window.AtivoBalao.abrir(inpAt, { foco: true }); return; }                              // Ctrl+Enter = balão dos ativos
                    if (e.key === 'Tab' && !e.ctrlKey && !e.altKey && !e.metaKey && !window.AtivoBalao.aberto()) {                                                    // Tab: Operação → Ativo → Operação da linha de baixo
                        e.preventDefault();
                        if (e.shiftKey) { irParaCelula(type, idx, 0, -1); return; }
                        if (idx + 1 < state.data.length) { irParaCelula(type, idx + 1, 0, 1); return; }
                        if (String(row.ativo == null ? '' : row.ativo).trim() !== '') { addRow(type, idx); setTimeout(() => irParaCelula(type, idx + 1, 0, 1), 80); }
                        return;
                    }
                }
                if (e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {          // TASK-195: Alt+↑/↓ move a linha (ou as selecionadas; numa linha de grupo, o grupo todo)
                    e.preventDefault(); e.stopPropagation();
                    moverLinhas(type, linhasParaMover(type, idx), e.key === 'ArrowUp' ? -1 : 1);
                    return;
                }
                if (e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && (e.key === 'ArrowLeft' || e.key === 'ArrowRight') && row.grupoDecl) {    // TASK-195: Alt+← recolhe, Alt+→ expande o grupo
                    e.preventDefault(); e.stopPropagation();
                    alternarRecolhido(type, row, e.key === 'ArrowLeft');
                    return;
                }
                if (e.key === 'Delete' && inpAt.value.trim() === '') {
                    e.preventDefault();
                    e.stopPropagation();          // TASK-087: o tratador global da grade (Delete = limpar a seleção) também limparia a linha de cima
                    deleteRow(type, idx);
                    setTimeout(() => irParaCelula(type, idx - 1, 1, -1), 0);       // o foco volta para o Ativo da linha de cima
                    return;
                }
                const consumed = handleAcKeydown(e, inpAt, type);
                if (e.key === 'Enter' && !consumed) {
                    e.preventDefault();
                    // Enter cria uma linha NOVA logo abaixo (como numa IDE); numa linha ainda em branco só desce, para não empilhar linhas vazias
                    if (String(row.ativo == null ? '' : row.ativo).trim() === '') irParaCelula(type, idx + 1, 1, 1);
                    else if (!enterEmLinhaDeGrupo(type, idx, row, e.shiftKey)) addRow(type, idx);          // TASK-195: linha de grupo → Enter cria membro no fim do grupo; Shift+Enter, linha solta logo abaixo do grupo
                }
                if (!consumed && !e.shiftKey && !e.ctrlKey && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
                    e.preventDefault();
                    irParaCelula(type, idx + (e.key === 'ArrowDown' ? 1 : -1), 1, e.key === 'ArrowDown' ? 1 : -1);
                }
            });
            // depois que o valor da linha mudou (blur ou gravação pelo balão): classificação, índices, filtros, quantidades
            const confirmarMudanca = snapAtivo => {
                const mudouHash = sincronizarHash(row, snapAtivo);                    // TASK-182: "#" no começo = comentário (no mesmo passo de Ctrl+Z)
                if (type === 'outros') {                              // TASK-065
                    reclassificarEntidade(row);
                    atualizarCelulaEntidade();
                }
                buildAtivoSets();
                buildDataLists();
                if (type === 'cabos') {
                    const currentIdx = parseInt(tr.dataset.index);
                    recalcRowAndAbove(currentIdx);
                    // Atualiza visualmente os inputs de qtdAtivos
                    const body = document.getElementById(state.bodyId);
                    if (body) {
                        Array.from(body.children).forEach(trLoop => {
                            const loopIdx = parseInt(trLoop.dataset.index);
                            if (loopIdx === currentIdx || loopIdx === currentIdx - 1) {
                                const qtdInp = trLoop.querySelector('.inp-qtd-ativos');
                                if (qtdInp && state.data[loopIdx]) {
                                    qtdInp.value = state.data[loopIdx].qtdAtivos !== undefined ? state.data[loopIdx].qtdAtivos : '';
                                }
                            }
                        });
                    }
                }
                refreshAllFilters(type);
                atualizarResumoRedeUI();
                if (mudouHash) renderPreservandoFoco(type);
            };
            inpAt._gravarCheio = novo => {                                              // o balão de parâmetros grava o texto cheio da linha (um passo do Ctrl+Z)
                const snap = row.ativo;
                if (novo === snap) return;
                row.ativo = snap; pushHistory(); row.ativo = novo; prevAtivo = novo;
                inpAt._sincronizarCheio();
                confirmarMudanca(snap);
                aplicarCoresJa();                                                         // TASK-247: preencheu pelo balão → a cor muda na hora (laranja → verde)
            };
            inpAt.addEventListener('blur', () => {
                setTimeout(() => {
                    hideAutocomplete();
                    const atual = textoDoAtivo();
                    if (atual !== prevAtivo) {
                        const snapAtivo = prevAtivo;  // guarda o valor antes
                        prevAtivo = atual;
                        // Restaura temporariamente o valor antigo para capturar o snapshot correto
                        row.ativo = snapAtivo;
                        pushHistory();  // snapshot com valor ANTES
                        row.ativo = atual;  // aplica o novo valor
                        confirmarMudanca(snapAtivo);
                    }
                }, 150);
            });
            inpAt.addEventListener('focus', () => {
                prevAtivo = textoDoAtivo();
                if (inpAt.value.length === 0) showAutocomplete(inpAt, type);
            });

            tdAt.appendChild(inpAt);
            tdAt.dataset.col = '1';
            tr.appendChild(tdAt);

            /* ── Fases / Tensão / Nível (só Cabos — TASK-092): vêm do texto original; o que for digitado vale mais ── */
            if (type === 'cabos') {
                const VRg = window.VinculacaoRegras, ex = VRg ? VRg.extraCabo(row) : { fases: 1, tensao: '', nivel: 'BT' };
                [['fases', ex.fases, 'text', '3'], ['tensao', ex.tensao, 'text', '13,8kV'], ['nivel', ex.nivel, 'select', '']].forEach(([campo, valor, tp]) => {
                    const td = document.createElement('td');
                    td.className = 'col-cabo-extra';
                    const inp = tp === 'select' ? document.createElement('select') : document.createElement('input');
                    inp.className = 'inp-cabo-extra' + (row[campo] === undefined || row[campo] === '' ? ' derivado' : '');
                    inp.dataset.field = campo;
                    if (tp === 'select') { ['MT', 'BT'].forEach(v => { const o = document.createElement('option'); o.value = v; o.textContent = v; if (v === String(valor)) o.selected = true; inp.appendChild(o); }); }
                    else { inp.type = 'text'; inp.value = valor === '' || valor == null ? '' : String(valor); inp.autocomplete = 'off'; }
                    inp.title = row[campo] === undefined || row[campo] === '' ? 'Derivado do texto original (editável)' : 'Definido à mão';
                    inp.addEventListener('change', () => {
                        pushHistory();
                        const v = inp.value.trim();
                        if (v === '') delete row[campo]; else row[campo] = campo === 'tensao' ? v : (campo === 'nivel' ? v.toUpperCase() : v);
                        inp.classList.toggle('derivado', row[campo] === undefined);
                        atualizarResumoRedeUI();
                    });
                    inp.addEventListener('keydown', e => { if (e.key === 'Enter' && tp !== 'select') { e.preventDefault(); inp.blur(); } });
                    td.appendChild(inp);
                    tr.appendChild(td);
                });
            }

            /* ── Qtd Ativos (apenas para cabos) ── */
            if (type === 'cabos') {
                const tdQtd = document.createElement('td');
                tdQtd.style.textAlign = 'center';
                const inpQtd = document.createElement('input');
                inpQtd.type = 'text';
                inpQtd.className = 'inp-qtd-ativos';
                inpQtd.dataset.field = 'qtdAtivos';
                inpQtd.value = row.qtdAtivos !== undefined ? row.qtdAtivos : '';
                inpQtd.autocomplete = 'off';
                let prevQtd = inpQtd.value;

                inpQtd.addEventListener('blur', () => {
                    const val = inpQtd.value.trim();
                    const num = val === '' ? 0 : parseInt(val, 10);
                    if (inpQtd.value !== prevQtd) {
                        pushHistory();  // snapshot ANTES da mudança
                        prevQtd = inpQtd.value;
                        if (!isNaN(num)) {
                            row.qtdAtivos = num;
                        }
                    }
                });

                tdQtd.appendChild(inpQtd);
                tdQtd.dataset.col = '2';
                tr.appendChild(tdQtd);
            }

            /* ── Delete button ── */
            const tdDel = document.createElement('td');
            tdDel.className = 'col-del';
            btnDel = document.createElement('button');
            btnDel.className = 'btn-row-del';
            btnDel.title = 'Excluir linha';
            btnDel.innerHTML = '';
            btnDel.addEventListener('click', (e) => {
                e.stopPropagation();
                deleteRow(type, idx);
            });
            tdDel.appendChild(btnDel);
            tr.appendChild(tdDel);

            /* ── Context menu trigger ── */
            tr.addEventListener('contextmenu', e => {
                e.preventDefault();
                ctxRow = tr; ctxType = type;
                const a = document.activeElement, td = a && a.closest ? a.closest('td[data-col]') : null;
                ctxFoco = td && tr.contains(td) ? { tipo: type, idx: parseInt(tr.dataset.index), col: parseInt(td.dataset.col) } : null;
                showCtxMenu(e.clientX, e.clientY);
            });
            if (row.bloqueada) {
                tr.querySelectorAll('input').forEach(i => { i.readOnly = true; });
                tr.querySelectorAll('select').forEach(s => { s.disabled = true; });
                btnDel.disabled = true;
            }
            };          // fim de montar(): a linha completa (células, campos e eventos)

            // TASK-069: agrupamento visual (separador antes de cada trecho de valores iguais da coluna agrupada)
            const campoGrupo = agrupamento[type];
            if (campoGrupo) {
                const chave = String(row[campoGrupo] == null || row[campoGrupo] === '' ? '(vazio)' : row[campoGrupo]);
                const ant = idx > 0 ? state.data[idx - 1] : null;
                const chaveAnt = ant ? String(ant[campoGrupo] == null || ant[campoGrupo] === '' ? '(vazio)' : ant[campoGrupo]) : null;
                if (chave !== chaveAnt) {
                    let n = 0;
                    for (let k = idx; k < state.data.length; k++) {
                        const v = state.data[k][campoGrupo];
                        if (String(v == null || v === '' ? '(vazio)' : v) !== chave) break;
                        n++;
                    }
                    const sep = document.createElement('tr');
                    sep.className = 'grupo-sep';
                    sep.dataset.grupoCampo = campoGrupo;
                    const tdSep = document.createElement('td');
                    tdSep.colSpan = colunasVisiveis;
                    tdSep.textContent = `▸ ${campoGrupo.toUpperCase()}: ${chave}  (${n})`;
                    sep.appendChild(tdSep);
                    frag.appendChild(sep);
                }
            }
            if (row.grupoDecl) { tr.classList.add('linha-grupo-decl'); tr.dataset.grupo = row.gid; tr.dataset.gcor = String(Math.max(0, window.GruposLinhas.ordemDoGrupo(state.data, row.gid)) % 6); }      // TASK-195
            else if (row.grupoDe) { tr.classList.add('linha-grupo-membro'); tr.dataset.grupo = row.grupoDe; tr.dataset.gcor = String(Math.max(0, window.GruposLinhas.ordemDoGrupo(state.data, row.grupoDe)) % 6); }
            if (row.comentario) tr.classList.add('linha-comentario');
            else if (row.congelada && !row.grupoDecl) tr.classList.add('linha-congelada');
            if (row.bloqueada) tr.classList.add('linha-bloqueada');
            if (lazy) {
                tr.classList.add('lazy-row'); tr.style.height = hLazy + 'px'; tr._montar = montar;
                // Exceder o cabeçalho cria colunas vazias no layout fixo e comprime o Ativo.
                const td = document.createElement('td');
                td.colSpan = colunasVisiveis;
                tr.appendChild(td);
                _lazy[type].push(tr);
            } else montar();
            frag.appendChild(tr);
        });
        body.appendChild(frag);
        if (lazy) { hidratarJanela(type, 500); agendarHidratacao(type); }
        agendarCoresAposDesenhar();                                     // TASK-167 (só com a opção ligada); Outros e Cabos — depois do desenho, uma vez só
        atualizarIndicadoresOrdem(type);

        updateCounters();
        agendarResumoRede();                                            // painel de resumo da rede: depois do desenho (uma vez só)

        // Após render, aplica filtros ativos sem rebuildar dropdowns
        applyFilters(type);
        ajustarLinhasVazias(type);                    // TASK-111: linhas em branco até o fim da área visível
    }

    /* ═══════════════════════════════════════
       CRUD
    ═══════════════════════════════════════ */
    function deleteRow(type, idx) {
        if (bloqueada(tableStates[type].data[idx])) { showToast('Linha bloqueada: desbloqueie antes de excluir.'); return; }
        pushHistory();  // snapshot ANTES da exclusão
        limparGrade();  // TASK-087: a seleção apontava para posições que mudaram
        tableStates[type].data.splice(idx, 1);
        limparOrfasDeGrupo(type);
        // TASK-032: ao remover uma linha de Outros, os vínculos cabo<->estrutura são baseados em
        // índice — reajusta as referências pra não apontarem para a linha errada depois do shift.
        if (type === 'outros') _reajustarVinculosAposRemoverOutros(idx);
        renderTable(type);
        buildAtivoSets(); buildDataLists();
        refreshAllFilters(type);
    }

    function _reajustarVinculosAposRemoverOutros(idxRemovido) {
        tableStates.cabos.data.forEach(cabo => {
            if (Array.isArray(cabo.vinculoPostes)) { cabo.vinculoPostes = cabo.vinculoPostes.filter(l => l !== idxRemovido).map(l => (l > idxRemovido ? l - 1 : l)); if (!cabo.vinculoPostes.length && !cabo._mapaTopologia) delete cabo.vinculoPostes; }
            if (!Array.isArray(cabo.vinculoEstruturas)) return;
            const VR = window.VinculacaoRegras;       // TASK-094: o id da unidade é linha (+ posição/100); só a linha muda
            cabo.vinculoEstruturas = cabo.vinculoEstruturas
                .filter(i => VR.linhaDe(i) !== idxRemovido)
                .map(i => VR.linhaDe(i) > idxRemovido ? VR.idUnidade(VR.linhaDe(i) - 1, VR.slotDe(i)) : i);
        });
    }

    function addRow(type, afterIdx = -1) {
        pushHistory();  // snapshot ANTES da inserção
        // TASK-111: operação da linha PREENCHIDA mais próxima acima da nova (tabela vazia ou sem nada preenchido acima = I)
        const op = window.GradeTabela.operacaoHerdada(tableStates[type].data, afterIdx === -1 ? tableStates[type].data.length : afterIdx + 1);

        const newRow = { entidade: '0', operacao: op, ativo: '' };
        if (afterIdx === -1) tableStates[type].data.push(newRow);
        else {
            const dd = tableStates[type].data, acima = dd[afterIdx], abaixo = dd[afterIdx + 1], gu = acima && (acima.grupoDecl ? acima.gid : acima.grupoDe);
            if (gu && abaixo && abaixo.grupoDe === gu) { newRow.grupoDe = gu; const dc = dd.find(r => r.grupoDecl && r.gid === gu); if (dc) delete dc.grupoRecolhido; }          // TASK-195: entre dois itens do mesmo grupo (ou logo abaixo da declaração) a linha nova é do grupo
            dd.splice(afterIdx + 1, 0, newRow);
        }
        limparGrade();  // TASK-087: a seleção antiga (linha de cima) não pode sobreviver ao foco na linha nova
        renderTable(type);
        // Foca no input Ativo da nova linha
        const body = document.getElementById(tableStates[type].bodyId);
        const target = afterIdx === -1 ? body.lastElementChild : body.children[afterIdx + 1];
        hidratarLinha(target);
        if (target) {
            const inp = target.querySelector('.inp-ativo');
            if (inp) { setTimeout(() => inp.focus(), 30); }
        }
        refreshAllFilters(type);
    }


    /* ═══════════════════════════════════════
       GRUPOS DE LINHAS (TASK-195) — static/grupos_linhas.js tem a lógica pura
       «@P1 Rua X» declara um grupo (linha congelada); os membros ficam logo abaixo e podem ser recolhidos (+/−). Sem nenhuma linha «@» nada disto entra em ação.
    ═══════════════════════════════════════ */
    const temGrupoNaTabela = type => tableStates[type].data.some(r => r && (r.grupoDecl || r.grupoDe));
    /** Vínculos cabo↔estrutura são por índice de Outros: depois de reordenar, `mapa[antigo] = novo` (linha que saiu: undefined → vínculo descartado). */
    function _remapearVinculosOutros(mapa, cabos = tableStates.cabos.data) {
        const VR = window.VinculacaoRegras;
        cabos.forEach(c => {
            if (!c) return;
            if (Array.isArray(c.vinculoPostes)) { c.vinculoPostes = c.vinculoPostes.map(l => mapa[l]).filter(l => l !== undefined); if (!c.vinculoPostes.length && !c._mapaTopologia) delete c.vinculoPostes; }
            if (Array.isArray(c.vinculoEstruturas)) c.vinculoEstruturas = c.vinculoEstruturas.map(i => { const l = mapa[VR.linhaDe(i)]; return l !== undefined ? VR.idUnidade(l, VR.slotDe(i)) : null; }).filter(i => i !== null);
        });
    }
    /** Marca de grupo que ficou sem declaração (linha excluída) some; os membros viram linhas comuns. */
    function limparOrfasDeGrupo(type) {
        const GL = window.GruposLinhas; if (!GL || !temGrupoNaTabela(type)) return;
        tableStates[type].data = GL.normalizar(tableStates[type].data);
    }
    /** Depois de inserir linhas: os membros voltam para junto da declaração (e, em Outros, os vínculos acompanham). */
    function reorganizarGrupos(type) {
        const GL = window.GruposLinhas; if (!GL || !temGrupoNaTabela(type)) return;
        const antiga = tableStates[type].data, nova = GL.normalizar(antiga);
        if (nova.length === antiga.length && nova.every((r, k) => r === antiga[k])) return;
        if (type === 'outros') _remapearVinculosOutros(GL.mapa(antiga, nova));
        tableStates[type].data = nova;
    }
    /** Troca a lista da tabela (já reordenada por GruposLinhas), acerta os vínculos, redesenha e põe o cursor no Ativo da linha `foco`. */
    function aplicarLista(type, nova, foco) {
        const GL = window.GruposLinhas, antiga = tableStates[type].data;
        if (type === 'outros') _remapearVinculosOutros(GL.mapa(antiga, nova));
        tableStates[type].data = nova;
        limparGrade();
        document.querySelectorAll('tr.row-selected').forEach(x => x.classList.remove('row-selected'));
        renderTable(type);
        buildAtivoSets(); buildDataLists();
        refreshAllFilters(type);
        atualizarResumoRedeUI();
        if (!foco) return;
        const i = nova.indexOf(foco), body = document.getElementById(tableStates[type].bodyId);
        const tr = Array.from(body.children).find(x => parseInt(x.dataset.index) === i);
        if (!tr) return;
        hidratarLinha(tr);
        try { tr.scrollIntoView({ block: 'nearest' }); } catch (e) { /* sem scroll */ }
        setTimeout(() => { const inp = tr.querySelector('.inp-ativo'); if (inp && document.body.contains(inp)) inp.focus(); }, 30);
    }
    /** Enter / Shift+Enter numa linha de grupo. Devolve false quando a linha não é de grupo (aí vale o Enter de sempre). */
    function addRowGrupo(type, idx, solta) {
        const GL = window.GruposLinhas, dados = tableStates[type].data, row = dados[idx];
        if (!GL || !row) return false;
        const marcar = GL.ehDeclaracao(row.ativo) && !row.grupoDecl && !row.bloqueada;       // acabou de digitar «@P1» e deu Enter sem sair da célula
        if (!marcar && !GL.bloco(dados, idx)) return false;
        pushHistory();
        if (marcar) sincronizarHash(row, '');
        const b = GL.bloco(dados, idx); if (!b) return false;
        const nova = { entidade: '0', operacao: window.GradeTabela.operacaoHerdada(dados, b.fim + 1), ativo: '' };          // igual ao addRow de sempre
        if (!solta) { const dc = dados.find(x => x.grupoDecl && x.gid === b.gid); if (dc) delete dc.grupoRecolhido; }                 // a linha nova do grupo não pode nascer escondida
        const r = solta ? GL.novaLinhaSolta(dados, idx, nova) : GL.novaLinhaNoGrupo(dados, idx, nova);
        aplicarLista(type, r.lista, r.foco);
        return true;
    }
    /** true quando a linha é de grupo (ou acabou de ser digitada como «@…») e o Enter já foi tratado aqui. */
    function enterEmLinhaDeGrupo(type, idx, row, shift) {
        const GL = window.GruposLinhas; if (!GL) return false;
        if (!(GL.ehDeclaracao(row.ativo) || GL.bloco(tableStates[type].data, idx))) return false;
        return addRowGrupo(type, idx, shift);
    }
    function criarBotaoGrupo(type, row) {
        const GL = window.GruposLinhas, w = document.createElement('span');
        w.className = 'grupo-ctl';
        const b = document.createElement('button');
        b.type = 'button'; b.className = 'grupo-toggle'; b.tabIndex = -1;
        const n = GL.contar(tableStates[type].data, row.gid);
        const pinta = () => { b.textContent = row.grupoRecolhido ? '+' : '−'; b.title = (row.grupoRecolhido ? 'Expandir' : 'Recolher') + ' o grupo (Alt+→ / Alt+←)'; };
        pinta();
        b.addEventListener('mousedown', e => e.preventDefault());
        b.addEventListener('click', e => { e.stopPropagation(); alternarRecolhido(type, row, !row.grupoRecolhido); });
        const c = document.createElement('span'); c.className = 'grupo-conta'; c.textContent = n; c.title = n + ' linha(s) no grupo';
        w.appendChild(b); w.appendChild(c);
        return w;
    }
    /** Recolhe/expande o grupo (só esconde as linhas; vale também depois de salvar). */
    function alternarRecolhido(type, row, recolher) {
        if (!row || !row.grupoDecl) return;
        if (!!row.grupoRecolhido === !!recolher) return;
        if (recolher) row.grupoRecolhido = true; else delete row.grupoRecolhido;
        const body = document.getElementById(tableStates[type].bodyId), tr = Array.from(body.children).find(x => x.dataset.index === String(tableStates[type].data.indexOf(row)));
        const btn = tr && tr.querySelector('.grupo-toggle');
        if (btn) { btn.textContent = row.grupoRecolhido ? '+' : '−'; btn.title = (row.grupoRecolhido ? 'Expandir' : 'Recolher') + ' o grupo (Alt+→ / Alt+←)'; }
        applyFilters(type);
        marcarProjetoAlterado();
    }
    /** Linhas que Alt+↑/↓ move: as selecionadas (se a atual está entre elas) ou só a atual. */
    function linhasParaMover(type, idx) {
        const body = document.getElementById(tableStates[type].bodyId);
        const sel = Array.from(body.querySelectorAll('tr.row-selected')).map(tr => parseInt(tr.dataset.index)).filter(i => !isNaN(i));
        return sel.includes(idx) ? sel : [idx];
    }
    /** Move linhas uma posição (Alt+↑/↓ e menu). Numa linha de grupo move o grupo inteiro; um item sai do grupo ao passar da borda e uma linha solta entra ao descer sobre a declaração. */
    function moverLinhas(type, idxs, dir) {
        const GL = window.GruposLinhas; if (!GL) return;
        const dados = tableStates[type].data;
        const alvos = Array.from(new Set(idxs)).filter(i => dados[i]).sort((a, b) => dir < 0 ? a - b : b - a);
        if (!alvos.length) return;
        const executar = lista => {                                // quem se mexe: as linhas (por objeto), na ordem que não as atropela
            const objs = alvos.map(i => lista[i]); let cur = lista, movido = false;
            objs.forEach(r => { const res = GL.mover(cur, cur.indexOf(r), dir); if (res) { cur = res.lista; movido = true; } });
            return { lista: cur, movido, foco: objs[objs.length - 1] };
        };
        if (!executar(dados.map(r => Object.assign({}, r))).movido) return;                  // ensaio em cópias: nada a mover não gasta passo de Ctrl+Z
        pushHistory();
        const r = executar(dados);
        aplicarLista(type, r.lista, r.foco);
        const tr = Array.from(document.getElementById(tableStates[type].bodyId).children).find(x => parseInt(x.dataset.index) === r.lista.indexOf(r.foco));
        if (tr) restaurarSelecaoLinhas(type, alvos.length > 1 ? alvos.map(i => r.lista.indexOf(dados[i])).filter(i => i >= 0) : []);
    }

    /* ═══════════════════════════════════════
       LINHAS VAZIAS ESTILO EXCEL (TASK-111)
       Abaixo das linhas de dados a tabela mostra linhas em branco até preencher a área visível. Elas NÃO são dados (ficam num <tbody> à parte):
       digitar no Ativo de uma delas cria a linha de dados (operação herdada da linha preenchida acima, ou I) e passa o foco para ela; colar várias
       linhas numa delas cola a partir da primeira linha em branco. Sem o botão Add.
    ═══════════════════════════════════════ */
    function tbodyVazio(type) {
        const body = document.getElementById(tableStates[type].bodyId);
        const id = tableStates[type].bodyId + '-vazio';
        let tb = document.getElementById(id);
        if (!tb) {
            tb = document.createElement('tbody');
            tb.id = id; tb.className = 'tbody-vazio';
            body.after(tb);
        }
        return tb;
    }

    /** Redesenha a tabela mantendo o foco na mesma célula (a edição do "#" acontece quando o foco já foi para outra célula). */
    function renderPreservandoFoco(type) {
        const body = document.getElementById(tableStates[type].bodyId), a = document.activeElement;
        const tr = a && a.closest ? a.closest('tr[data-index]') : null;
        const ref = tr && body.contains(tr) ? { idx: tr.dataset.index, campo: a.dataset ? a.dataset.field : '', cls: a.classList && a.classList[0] } : null;
        renderTable(type);
        if (!ref) return;
        const novo = Array.from(body.children).find(x => x.dataset.index === ref.idx);
        hidratarLinha(novo);
        const el = novo && ((ref.campo && novo.querySelector(`[data-field="${ref.campo}"]`)) || (ref.cls && novo.querySelector('.' + ref.cls)));
        if (el) el.focus();
    }

    function materializarLinhaVazia(type, texto, operacao) {
        const dados = tableStates[type].data;
        pushHistory();                                         // um passo de Ctrl+Z para criar a linha
        limparGrade();
        const GT = window.GradeTabela;
        const row = GT.novaLinha(type, operacao || GT.operacaoHerdada(dados, dados.length));          // operação escolhida na linha em branco, ou a herdada
        row.ativo = String(texto).toUpperCase();
        sincronizarHash(row, '');                                // TASK-182: linha nova que começa com "#" já nasce comentário
        dados.push(row);
        if (type === 'outros') reclassificarEntidade(row);
        else recalcRowAndAbove(dados.length - 1);
        renderTable(type);
        buildAtivoSets(); buildDataLists();
        refreshAllFilters(type);
        const body = document.getElementById(tableStates[type].bodyId);
        const tr = Array.from(body.children).find(x => parseInt(x.dataset.index) === dados.length - 1);
        hidratarLinha(tr);
        const inp = tr && tr.querySelector('.inp-ativo');
        if (inp) { inp.focus(); inp.setSelectionRange(inp.value.length, inp.value.length); }
        atualizarResumoRedeUI();
    }

    function focarLinhaVazia(type, i) {
        const tb = document.getElementById(tableStates[type].bodyId + '-vazio');
        const tr = tb && tb.children[i];
        const inp = tr && tr.querySelector('.inp-ativo-vazio');
        if (inp) { inp.focus(); return true; }
        return false;
    }

    function ajustarLinhasVazias(type) {
        if (type !== 'cabos' && type !== 'outros') return;
        const body = document.getElementById(tableStates[type].bodyId);
        if (!body) return;
        const tb = tbodyVazio(type);
        const dados = tableStates[type].data;
        const wrap = body.closest('.table-wrap');
        const MIN = 4, MAX = 300;                                                  // (constantes locais: renderTable roda antes desta seção na inicialização)
        let n = MIN;
        // Quantas linhas em branco cabem na área visível: conta com as medidas guardadas (altura da área, do cabeçalho e de uma linha) em vez de ler offsetHeight/clientHeight a cada
        // desenho — isso forçava o navegador a calcular o layout das centenas de linhas toda vez. As medidas são refeitas só quando a área muda de tamanho.
        let m = _medidasTabela[type];
        if (!m && wrap && wrap.offsetParent !== null && wrap.clientHeight > 0) {
            const tabela = body.closest('table'), amostra = body.querySelector('tr:not(.grupo-sep):not(.lazy-row)');
            m = _medidasTabela[type] = { h: wrap.clientHeight, cab: tabela.tHead ? tabela.tHead.offsetHeight : 0, linha: (amostra && amostra.offsetHeight) || tb.querySelector('tr')?.offsetHeight || 30 };
        }
        if (m) { _hLinhaReal[type] = m.linha; n = Math.max(MIN, Math.min(MAX, Math.floor((m.h - m.cab - dados.length * m.linha) / m.linha))); }
        const dica = window.GradeTabela.operacaoHerdada(dados, dados.length);
        const assinatura = [n, dados.length, dica, type === 'outros' ? 'o' : 'c'].join('|');
        if (tb.dataset.assinatura === assinatura && tb.children.length === n) return;
        tb.dataset.assinatura = assinatura;
        const frag = document.createDocumentFragment();
        for (let i = 0; i < n; i++) {
            const tr = document.createElement('tr');
            tr.className = 'linha-vazia';
            const cel = (cls, txt) => { const td = document.createElement('td'); if (cls) td.className = cls; if (txt) td.textContent = txt; tr.appendChild(td); return td; };
            cel('col-num', String(dados.length + i + 1));
            if (type === 'outros') cel('cel-ent', '');
            const tdOpVz = cel('vz-op', '');                                           // operação da nova linha: já vem a que ela receberia (herdada); dá para escolher ANTES de digitar o ativo
            const selOpVz = criarCampoOperacao(OPERACOES.includes(dica) ? dica : 'I', 'sel-operacao inp-operacao vz-sel', () => { selOpVz.dataset.tocada = '1'; });
            selOpVz.setAttribute('aria-label', 'Nova linha: operação (opcional; escolha antes de digitar o ativo)');
            tdOpVz.appendChild(selOpVz);
            const tdAt = cel('', '');
            const inp = document.createElement('input');
            inp.type = 'text'; inp.className = 'inp-ativo-vazio'; inp.autocomplete = 'off';
            inp.setAttribute('aria-label', 'Nova linha: digite o ativo');
            inp.addEventListener('input', () => { if (inp.value.trim() !== '') materializarLinhaVazia(type, inp.value, selOpVz.dataset.tocada ? selOpVz.value : null); });
            inp.addEventListener('paste', e => {
                const texto = (e.clipboardData || window.clipboardData).getData('text');
                if (!window.GradeTabela.ehColagemMultipla(texto)) return;               // valor único: colagem nativa (cai no 'input')
                e.preventDefault();
                colarNaGrade(type, dados.length, 1, texto);
            });
            inp.addEventListener('keydown', e => {
                if (e.key === 'ArrowDown' || e.key === 'Enter') { e.preventDefault(); focarLinhaVazia(type, i + 1); }
                else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    if (i > 0) focarLinhaVazia(type, i - 1);
                    else if (dados.length) irParaCelula(type, dados.length - 1, 1, -1);
                }
            });
            tdAt.appendChild(inp);
            if (type === 'cabos') { cel('col-cabo-extra', ''); cel('col-cabo-extra', ''); cel('col-cabo-extra', ''); cel('', ''); }
            cel('col-del', '');
            frag.appendChild(tr);
        }
        tb.innerHTML = '';
        tb.appendChild(frag);
    }

    /* mantém as linhas em branco do tamanho da área visível (redimensionar a janela, abrir/fixar painéis, mudar o layout) */
    (function observarAlturaDasTabelas() {
        let pendente = false;
        const agendar = () => {
            if (pendente) return;
            pendente = true;
            requestAnimationFrame(() => { pendente = false; _medidasTabela.cabos = _medidasTabela.outros = null; ['cabos', 'outros'].forEach(ajustarLinhasVazias); });
        };
        window.addEventListener('resize', agendar);
        if (window.ResizeObserver) {
            const ro = new ResizeObserver(agendar);
            ['table-cabos', 'table-outros'].forEach(id => { const w = document.getElementById(id)?.closest('.table-wrap'); if (w) ro.observe(w); });
        }
    })();

    /* ═══════════════════════════════════════
       BOTÕES DE AÇÃO
    ═══════════════════════════════════════ */
    document.getElementById('btn-add-cabos').addEventListener('click', () => addRow('cabos'));
    document.getElementById('btn-add-outros').addEventListener('click', () => addRow('outros'));

    document.getElementById('btn-del-cabos').addEventListener('click', () => deleteSelectedOrLast('cabos'));
    document.getElementById('btn-del-outros').addEventListener('click', () => deleteSelectedOrLast('outros'));

    document.getElementById('btn-copy-cabos').addEventListener('click', () => copyTable('cabos'));
    document.getElementById('btn-copy-outros').addEventListener('click', () => copyTable('outros'));
    
    document.getElementById('btn-clear-cabos').addEventListener('click', () => {
        if(confirm('Tem certeza que deseja limpar toda a tabela Cabos?')) {
            pushHistory();  // snapshot ANTES de limpar
            tableStates['cabos'].data = [];
            renderTable('cabos');
            buildAtivoSets(); buildDataLists();
            refreshAllFilters('cabos');
        }
    });
    document.getElementById('btn-clear-outros').addEventListener('click', () => {
        if(confirm('Tem certeza que deseja limpar toda a tabela Outros?')) {
            pushHistory();  // snapshot ANTES de limpar
            tableStates['outros'].data = [];
            renderTable('outros');
            buildAtivoSets(); buildDataLists();
            refreshAllFilters('outros');
        }
    });

    /* ═══════════════════════════════════════
       FILTROS POR COLUNA (estilo Excel)
    ═══════════════════════════════════════ */
    function getUniqueValues(type, field) {
        return [...new Set(tableStates[type].data.map(r => r[field] || ''))].sort();
    }

    function buildFilterDropdown(containerId, type, field) {
        const container = document.getElementById(containerId);
        if (!container) return;
        const values = getUniqueValues(type, field);
        const selections = filterSelections[type][field];
        const isAllSelected = selections.size === 0 || selections.size === values.length;

        container.innerHTML = `
            <div class="r-filter-trigger" title="Filtrar ${field}">
                <span>${selections.size > 0 ? selections.size + ' sel' : 'Todos'}</span>
                <div class="r-filter-indicator ${selections.size > 0 ? 'active' : ''}"></div>
            </div>
            <div class="r-filter-dropdown">
                <input type="text" class="r-filter-search" placeholder="Pesquisar...">
                <div class="r-filter-options-list">
                    <label class="r-filter-option r-select-all-opt">
                        <input type="checkbox" ${isAllSelected ? 'checked' : ''}>
                        <span>(Selecionar Tudo)</span>
                    </label>
                    <div class="r-options-inner">
                        ${values.map(val => `
                            <label class="r-filter-option" data-value="${val}">
                                <input type="checkbox" ${(selections.has(val) || selections.size === 0) ? 'checked' : ''}>
                                <span title="${val}">${val === '' ? '(Vazio)' : val}</span>
                            </label>
                        `).join('')}
                    </div>
                </div>
                <div class="r-filter-actions">
                    <button class="r-btn-filter-action r-clear-btn">Limpar</button>
                    <button class="r-btn-filter-action r-all-btn">Todos</button>
                </div>
            </div>
        `;

        const trigger    = container.querySelector('.r-filter-trigger');
        const dropdown   = container.querySelector('.r-filter-dropdown');
        const searchInp  = container.querySelector('.r-filter-search');
        const inner      = container.querySelector('.r-options-inner');
        const mainChk    = container.querySelector('.r-select-all-opt input');
        const clearBtn   = container.querySelector('.r-clear-btn');
        const allBtn     = container.querySelector('.r-all-btn');

        trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            // Fecha outros
            document.querySelectorAll('.r-filter-dropdown.active').forEach(d => {
                if (d !== dropdown) d.classList.remove('active');
            });
            dropdown.classList.toggle('active');
        });
        dropdown.addEventListener('click', (e) => e.stopPropagation());

        const onSelectionChange = () => {
            const all   = inner.querySelectorAll('input');
            const chkd  = inner.querySelectorAll('input:checked');
            mainChk.checked = chkd.length === all.length;
            selections.clear();
            if (chkd.length < all.length) {
                chkd.forEach(cb => selections.add(cb.parentElement.dataset.value));
            }
            // Atualiza indicador visualmente sem fechar
            const trigger2 = container.querySelector('.r-filter-trigger span');
            const ind = container.querySelector('.r-filter-indicator');
            if (trigger2) trigger2.textContent = selections.size > 0 ? selections.size + ' sel' : 'Todos';
            if (ind) ind.classList.toggle('active', selections.size > 0);
            applyFilters(type);
        };

        inner.querySelectorAll('input').forEach(cb => cb.addEventListener('change', onSelectionChange));

        mainChk.addEventListener('change', (e) => {
            inner.querySelectorAll('.r-filter-option').forEach(opt => {
                if (opt.style.display !== 'none') opt.querySelector('input').checked = e.target.checked;
            });
            onSelectionChange();
        });

        clearBtn.addEventListener('click', () => {
            inner.querySelectorAll('input').forEach(cb => cb.checked = false);
            mainChk.checked = false;
            onSelectionChange();
        });

        allBtn.addEventListener('click', () => {
            inner.querySelectorAll('input').forEach(cb => cb.checked = true);
            mainChk.checked = true;
            onSelectionChange();
        });

        searchInp.addEventListener('input', (e) => {
            const term = e.target.value.toLowerCase();
            inner.querySelectorAll('.r-filter-option').forEach(opt => {
                const val = (opt.dataset.value || '').toLowerCase();
                opt.style.display = val.includes(term) ? '' : 'none';
            });
        });
    }

    function refreshAllFilters(type) {
        camposFiltro(type).forEach(field => {
            buildFilterDropdown(FILTER_CONTAINERS[type][field], type, field);
        });
    }

    function applyFilters(type) {
        const state = tableStates[type];
        const body  = document.getElementById(state.bodyId);
        if (!body) return;
        const sels = filterSelections[type];
        const recolhidos = new Set(state.data.filter(r => r && r.grupoDecl && r.grupoRecolhido).map(r => r.gid));      // TASK-195: grupo recolhido esconde os membros (mesmo mecanismo do filtro)
        Array.from(body.children).forEach(tr => {
            const idx = parseInt(tr.dataset.index);
            if (isNaN(idx)) return;
            const row = state.data[idx];
            if (!row) return;
            let visible = !(recolhidos.size && row.grupoDe && recolhidos.has(row.grupoDe));
            for (const field of camposFiltro(type)) {
                const sel = sels[field];
                if (sel.size > 0 && !sel.has(row[field] || '')) {
                    visible = false; break;
                }
            }
            tr.style.display = visible ? '' : 'none';
        });
        // separador de grupo só aparece se sobrou alguma linha visível depois dele
        const filhos = Array.from(body.children);
        filhos.forEach((tr, i) => {
            if (!tr.classList.contains('grupo-sep')) return;
            let algum = false;
            for (let k = i + 1; k < filhos.length && !filhos[k].classList.contains('grupo-sep'); k++) if (filhos[k].style.display !== 'none') { algum = true; break; }
            tr.style.display = algum ? '' : 'none';
        });
        updateCounters();
    }


    function deleteSelectedOrLast(type) {
        const body = document.getElementById(tableStates[type].bodyId);
        const selected = Array.from(body.querySelectorAll('tr.row-selected'));
        const doIntervalo = selected.length ? [] : linhasDoIntervaloSelecionado(type);
        if (selected.length > 0 || doIntervalo.length > 0) {
            const idxs = (selected.length ? selected.map(tr => parseInt(tr.dataset.index)) : doIntervalo).sort((a,b) => b-a).filter(i => !bloqueada(tableStates[type].data[i]));          // linha bloqueada não é excluída
            if (!idxs.length) { showToast('Linha bloqueada: desbloqueie antes de excluir.'); return; }
            pushHistory();  // snapshot ANTES da exclusão em lote
            // Remove de trás pra frente para não deslocar índices
            idxs.forEach(i => {
                tableStates[type].data.splice(i, 1);
                if (type === 'outros') _reajustarVinculosAposRemoverOutros(i);
            });
            limparOrfasDeGrupo(type);
            limparGrade();
            renderTable(type);
            buildAtivoSets(); buildDataLists();
            refreshAllFilters(type);
        } else {
            // Remove a última linha
            if (tableStates[type].data.length > 0 && bloqueada(tableStates[type].data[tableStates[type].data.length - 1])) { showToast('Linha bloqueada: desbloqueie antes de excluir.'); return; }
            if (tableStates[type].data.length > 0) {
                pushHistory();  // snapshot ANTES da exclusão
                tableStates[type].data.pop();
                limparOrfasDeGrupo(type);
                renderTable(type);
                buildAtivoSets(); buildDataLists();
                refreshAllFilters(type);
            }
        }
    }

    /* ═══════════════════════════════════════
       COPIAR TABELA
    ═══════════════════════════════════════ */
    function copyTable(type) {
        const body = document.getElementById(tableStates[type].bodyId);
        let text = '';
        Array.from(body.children).forEach(tr => {
            if (tr.style.display === 'none') return; // Respeita o filtro
            const row = tableStates[type].data[parseInt(tr.dataset.index)];
            if (!row) return;
            text += `${row.operacao}\t${row.ativo}\n`;
        });
        if (!text.trim()) { showToast('Tabela vazia ou tudo oculto.'); return; }
        copyToClipboard(text, () => {
            const btn = document.getElementById(`btn-copy-${type}`);
            const orig = btn.innerHTML;
            btn.innerHTML = '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg> Copiado!';
            setTimeout(() => btn.innerHTML = orig, 1200);
        });
    }

    function copyToClipboard(text, cb) {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(text).then(() => { if(cb) cb(); }).catch(() => fallbackCopy(text, cb));
        } else fallbackCopy(text, cb);
    }

    function fallbackCopy(text, cb) {
        const ta = document.createElement('textarea');
        ta.value = text; ta.style.position = 'fixed'; ta.style.left = '-9999px';
        document.body.appendChild(ta); ta.focus(); ta.select();
        try { document.execCommand('copy'); if(cb) cb(); } catch(e) {}
        document.body.removeChild(ta);
    }

    /* ═══════════════════════════════════════
       CONTEXT MENU
    ═══════════════════════════════════════ */
    /** Linhas que o menu de contexto afeta: as selecionadas (se a clicada está entre elas) ou só a clicada. */
    function _linhasDoMenu() {
        if (!ctxRow || !ctxType) return [];
        const idx = parseInt(ctxRow.dataset.index);
        const body = document.getElementById(tableStates[ctxType].bodyId);
        const sel = Array.from(body.querySelectorAll('tr.row-selected')).map(tr => parseInt(tr.dataset.index));
        return sel.includes(idx) ? sel : [idx];
    }
    /** Congelar/descongelar e bloquear/desbloquear linhas (um passo de Ctrl+Z). Se todas já têm a marca, tira; senão, põe em todas. Linha bloqueada só aceita ser desbloqueada. */
    /** TASK-196: marcar a linha do grupo («@») vale para o grupo todo — junta os membros aos índices. */
    function expandirGrupos(tipo, idxs) {
        const dados = tableStates[tipo].data, saida = new Set(idxs);
        idxs.forEach(i => { const r = dados[i]; if (r && r.grupoDecl) dados.forEach((m, k) => { if (m.grupoDe === r.gid) saida.add(k); }); });
        return Array.from(saida).sort((a, b) => a - b);
    }
    function alternarMarcaLinhas(tipo, idxs, campo) {
        const dados = tableStates[tipo].data;
        idxs = expandirGrupos(tipo, idxs);
        let alvos = idxs.filter(i => dados[i] && !dados[i]._mapaPosteFicticio);
        if (campo === 'congelada') {
            const comentarios = alvos.filter(i => dados[i].comentario).length;
            const decls = alvos.filter(i => dados[i].grupoDecl).length;
            alvos = alvos.filter(i => !bloqueada(dados[i]) && !dados[i].comentario && !dados[i].grupoDecl);          // a linha «@» já é sempre congelada: congelar o grupo age nos membros
            if (!alvos.length) { showToast(decls ? 'O grupo não tem linhas para congelar/descongelar (ou estão bloqueadas).' : comentarios ? 'Comentário já fica fora de todos os cálculos: tire o comentário para mudar o congelamento.' : 'Linha bloqueada: desbloqueie antes de congelar ou descongelar.'); return; }
        } else if (!alvos.length) return;
        const poe = !alvos.every(i => dados[i][campo]);
        pushHistory();
        alvos.forEach(i => { if (poe) dados[i][campo] = true; else delete dados[i][campo]; });
        limparGrade();
        renderTable(tipo);
        restaurarSelecaoLinhas(tipo, alvos);
        buildAtivoSets(); buildDataLists();
        refreshAllFilters(tipo);
        atualizarResumoRedeUI();
        const nome = campo === 'congelada' ? ['congelada', 'descongelada'] : ['bloqueada', 'desbloqueada'];
        showToast(`${alvos.length} linha(s) ${poe ? nome[0] : nome[1]}. Ctrl+Z desfaz.`);
    }
    /** Comentário (menu de contexto): a linha vira uma anotação — congelada (fora de todos os cálculos), com outra aparência. Tirar o comentário devolve a linha ao normal. Um passo de Ctrl+Z. */
    function alternarComentario(tipo, idxs) {
        const dados = tableStates[tipo].data;
        const alvos = idxs.filter(i => dados[i] && !bloqueada(dados[i]) && !dados[i].grupoDecl);
        if (!alvos.length) { showToast(idxs.some(i => dados[i] && dados[i].grupoDecl) ? 'Linha de grupo (@) não vira comentário.' : 'Linha bloqueada: desbloqueie antes de mudar o comentário.'); return; }
        const poe = !alvos.every(i => dados[i].comentario);
        pushHistory();
        alvos.forEach(i => { if (poe) { dados[i].comentario = true; dados[i].congelada = true; } else { delete dados[i].comentario; delete dados[i].congelada; if (comecaComHash(dados[i].ativo)) dados[i].ativo = String(dados[i].ativo).replace(/^\s*#\s?/, ''); } });
        limparGrade();
        renderTable(tipo);
        restaurarSelecaoLinhas(tipo, alvos);
        buildAtivoSets(); buildDataLists();
        refreshAllFilters(tipo);
        atualizarResumoRedeUI();
        showToast(`${alvos.length} linha(s) ${poe ? 'viraram comentário' : 'voltaram ao normal'}. Ctrl+Z desfaz.`);
    }
    /** Duplicar linha: copia as linhas (a clicada ou as selecionadas) e coloca as cópias LOGO ABAIXO (num bloco, na mesma ordem). A cópia vira uma linha comum: sem bloqueio e sem o vínculo/geometria do original. Um passo de Ctrl+Z. */
    function duplicarLinhas(tipo, idxs) {
        const dados = tableStates[tipo].data;
        const alvos = Array.from(new Set(idxs)).filter(i => dados[i] && !dados[i]._mapaPosteFicticio).sort((a, b) => a - b);
        if (!alvos.length) return;
        pushHistory();
        const copias = alvos.map(i => {
            const c = deepClone(dados[i]);
            ['grupoDecl', 'gid', 'grupoRecolhido', 'bloqueada', '_x', '_y', '_geo', '_geoTipo', '_poste', '_posteAprox', '_mapaPoste', '_mapaCabo', '_mapaPoste500', '_mapaCabo500', '_mapaId', '_mapaPosteRef', '_mapaTopologia', '_mapaAncoras', 'vinculoEstruturas', 'vinculoOrigem', 'vinculoPostes', 'estruturasFicticias', 'cabosFicticios', 'cabosFicticiosSlots', 'origemAvulso'].forEach(k => delete c[k]);
            return c;
        });
        const depois = alvos[alvos.length - 1] + 1;
        dados.splice(depois, 0, ...copias);
        copias.forEach(c => { if (window.GruposLinhas && window.GruposLinhas.ehDeclaracao(c.ativo)) { c.ativo = String(c.ativo).replace(/^(\s*)@+\s*/, '$1'); delete c.congelada; } });          // TASK-195: a cópia de «@P1» não vira um 2º grupo
        reorganizarGrupos(tipo);
        if (tipo === 'cabos') recalcAllQtdAtivos();
        limparGrade();
        renderTable(tipo);
        buildAtivoSets(); buildDataLists();
        refreshAllFilters(tipo);
        atualizarResumoRedeUI();
        restaurarSelecaoLinhas(tipo, copias.map((_, k) => depois + k));
        showToast(`${copias.length} linha(s) duplicada(s) logo abaixo. Ctrl+Z desfaz.`);
    }

    /** Depois de redesenhar a tabela, mantém selecionadas as linhas que o usuário marcou (assim ❄/💬/🔒 podem ser desfeitos no mesmo grupo). */
    function restaurarSelecaoLinhas(tipo, idxs) {
        const body = document.getElementById(tableStates[tipo].bodyId);
        Array.from(body.querySelectorAll('tr[data-index]')).forEach(tr => { if (idxs.includes(parseInt(tr.dataset.index))) tr.classList.add('row-selected'); });
    }
    /** Botões ❄ 💬 🔒 do cabeçalho da tabela: agem nas linhas selecionadas (clique no número da linha) ou, sem seleção, na linha onde está o cursor. */
    function linhasParaMarcar(tipo) {
        const body = document.getElementById(tableStates[tipo].bodyId);
        const sel = Array.from(body.querySelectorAll('tr.row-selected')).map(tr => parseInt(tr.dataset.index)).filter(i => !isNaN(i));
        if (sel.length) return sel;
        const a = document.activeElement, tr = a && a.closest ? a.closest('tr[data-index]') : null;
        return tr && body.contains(tr) ? [parseInt(tr.dataset.index)] : [];
    }
    ['cabos', 'outros'].forEach(tipo => {
        const ligar = (id, fn) => { const b = document.getElementById(id); if (b) b.addEventListener('mousedown', e => e.preventDefault()); if (b) b.addEventListener('click', () => { const idxs = linhasParaMarcar(tipo); if (!idxs.length) { showToast('Selecione linhas (clique no número) ou posicione o cursor numa linha.'); return; } fn(tipo, idxs); }); };
        ligar('btn-congelar-' + tipo, (t, i) => alternarMarcaLinhas(t, i, 'congelada'));
        ligar('btn-bloquear-' + tipo, (t, i) => alternarMarcaLinhas(t, i, 'bloqueada'));
        ligar('btn-comentar-' + tipo, (t, i) => alternarComentario(t, i));
    });
    function showCtxMenu(x, y) {
        const itemEnt = document.getElementById('ctx-entidade');
        if (itemEnt) itemEnt.style.display = ctxType === 'outros' ? '' : 'none';
        const alvos = expandirGrupos(ctxType, _linhasDoMenu()).map(i => tableStates[ctxType].data[i]).filter(Boolean);
        const temGrupo = alvos.some(r => r.grupoDecl), alvosCong = alvos.filter(r => !r.grupoDecl);
        const tc = document.getElementById('ctx-congelar-txt'), tb = document.getElementById('ctx-bloquear-txt');
        if (tc) tc.textContent = (alvosCong.length && alvosCong.every(r => r.congelada) ? 'Descongelar ' : 'Congelar ') + (temGrupo ? 'grupo' : 'linha');
        const tcm = document.getElementById('ctx-comentario-txt');
        if (tcm) tcm.textContent = alvos.length && alvos.every(r => r.comentario) ? 'Tirar comentário' : 'Marcar como comentário';
        if (tb) tb.textContent = (alvos.length && alvos.every(r => r.bloqueada) ? 'Desbloquear ' : 'Bloquear ') + (temGrupo ? 'grupo' : 'linha');
        ctxView = 'main'; ctxMenu.classList.remove('sub'); ctxPos = { x, y };
        ctxMenu.classList.add('visible');
        posicionarMenu(ctxMenu, x, y);                                            // TASK-195: mede o menu aberto — vira para cima/esquerda e rola se não couber (antes cortava nas linhas de baixo)
    }
    /** Põe um menu flutuante (position:fixed) no ponto (x, y) sem deixá-lo sair da tela: sem espaço embaixo/à direita ele abre para cima/esquerda; mais alto que a tela, ganha rolagem. `yAcima` = onde acaba o menu quando abre para cima. */
    function posicionarMenu(el, x, y, yAcima) {
        const M = 8, W = window.innerWidth, H = window.innerHeight;
        el.style.maxHeight = (H - 2 * M) + 'px'; el.style.overflowY = 'auto';
        const w = el.offsetWidth, h = el.offsetHeight;
        let left = x, top = y;
        if (left + w > W - M) left = (x - w >= M) ? x - w : Math.max(M, W - M - w);
        if (top + h > H - M) { const ac = (yAcima === undefined ? y : yAcima) - h; top = ac >= M ? ac : Math.max(M, H - M - h); }
        el.style.left = Math.max(M, left) + 'px'; el.style.top = Math.max(M, top) + 'px';
    }
    /* Menu do botão direito pelo teclado: ↑↓ (Home/End) navegam pelas opções, Enter executa, Esc fecha */
    let ctxKb = -1, ctxView = 'main', ctxPos = { x: 0, y: 0 };
    const itensDoMenu = () => Array.from(ctxMenu.querySelectorAll('.ctx-r-item')).filter(el => el.offsetParent !== null);
    function marcarItemDoMenu(i) { ctxKb = i; itensDoMenu().forEach((el, k) => el.classList.toggle('kb', k === i)); }
    function hideCtxMenu(naoFocar) {
        ctxMenu.classList.remove('visible', 'sub'); ctxView = 'main'; ctxMenu.querySelectorAll('.kb').forEach(el => el.classList.remove('kb')); ctxKb = -1; ctxRow = null; ctxType = null;
        const f = ctxFoco; ctxFoco = null;
        if (f && !naoFocar) setTimeout(() => {                                          // depois da ação (que pode redesenhar a tabela): o cursor volta para a célula em que estava (Tab continua dali)
            const n = tableStates[f.tipo].data.length;
            if (n) focarCelula(f.tipo, Math.min(f.idx, n - 1), f.col);
        }, 0);
    }
    document.addEventListener('keydown', e => {
        if (!ctxMenu.classList.contains('visible')) return;
        const n = itensDoMenu().length;
        if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); e.stopPropagation(); if (n) marcarItemDoMenu(ctxKb < 0 ? (e.key === 'ArrowDown' ? 0 : n - 1) : (ctxKb + (e.key === 'ArrowDown' ? 1 : -1) + n) % n); }
        else if (e.key === 'Home' || e.key === 'End') { e.preventDefault(); e.stopPropagation(); if (n) marcarItemDoMenu(e.key === 'Home' ? 0 : n - 1); }
        else if ((e.key === 'Enter' || e.key === ' ') && ctxKb >= 0) { e.preventDefault(); e.stopPropagation(); const it = itensDoMenu()[ctxKb]; if (it) it.click(); }
        else if (e.key === 'ArrowLeft' && ctxView !== 'main' && !(e.target && e.target.tagName === 'INPUT')) { e.preventDefault(); e.stopPropagation(); abrirVistaGrupo(ctxView === 'lista' ? 'grupo' : 'main', true); }
        else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); hideCtxMenu(); }
    }, true);
    document.addEventListener('mousedown', e => { if (!ctxMenu.contains(e.target)) hideCtxMenu(true); });          // clicou em outro lugar: o foco vai para onde ele clicou


    /* ── Menu «Grupo ▸» (TASK-195): as vistas trocam o conteúdo do MESMO menu (sem janelas soltas que possam sair da tela): Grupo → Adicionar ao grupo → lista de grupos ── */
    let ctxSub = null;
    function _subEl() {
        if (!ctxSub) { ctxSub = document.createElement('div'); ctxSub.className = 'ctx-sub'; ctxMenu.appendChild(ctxSub); }
        return ctxSub;
    }
    function abrirVistaGrupo(vista, porTeclado) {
        const GL = window.GruposLinhas, tipo = ctxType; if (!GL || !tipo) return;
        ctxView = vista;
        if (vista === 'main') { ctxMenu.classList.remove('sub'); ctxKb = -1; ctxMenu.querySelectorAll('.kb').forEach(el => el.classList.remove('kb')); posicionarMenu(ctxMenu, ctxPos.x, ctxPos.y); if (porTeclado) { const it = document.getElementById('ctx-grupo'); marcarItemDoMenu(itensDoMenu().indexOf(it)); } return; }
        const dados = tableStates[tipo].data, idxs = _linhasDoMenu(), linhas = idxs.map(i => dados[i]).filter(Boolean);
        const sub = _subEl(); sub.innerHTML = '';
        const item = (txt, acao, extra) => { const d = document.createElement('div'); d.className = 'ctx-r-item'; d.textContent = txt; if (extra) { const c = document.createElement('span'); c.className = 'ctx-cont'; c.textContent = extra; d.appendChild(c); } d.addEventListener('click', acao); sub.appendChild(d); return d; };
        const sep = () => { const d = document.createElement('div'); d.className = 'ctx-r-sep'; sub.appendChild(d); };
        const titulo = txt => { const d = document.createElement('div'); d.className = 'ctx-r-titulo'; d.textContent = txt; sub.appendChild(d); };
        item('‹ Voltar', () => abrirVistaGrupo(vista === 'lista' || vista === 'criar' ? 'grupo' : 'main', true));
        sep();
        if (vista === 'grupo') {
            const comuns = linhas.filter(r => !r.grupoDecl), gids = new Set(linhas.map(r => r.grupoDecl ? r.gid : r.grupoDe).filter(Boolean));
            const gid = gids.size === 1 ? Array.from(gids)[0] : null, decl = gid ? dados.find(r => r.grupoDecl && r.gid === gid) : null;
            item('Adicionar ao grupo ▸', () => abrirVistaGrupo('lista', true)).style.display = comuns.length ? '' : 'none';
            item('Criar grupo…', () => abrirVistaGrupo('criar')).style.display = comuns.length ? '' : 'none';
            if (linhas.some(r => r.grupoDe)) item('Tirar do grupo', () => { const ix = idxs; hideCtxMenu(); acaoTirarDoGrupo(tipo, ix); });
            if (decl) { item(decl.grupoRecolhido ? 'Expandir o grupo' : 'Recolher o grupo', () => { hideCtxMenu(); alternarRecolhido(tipo, decl, !decl.grupoRecolhido); }); item('Desfazer o grupo', () => { hideCtxMenu(); acaoDesfazerGrupo(tipo, gid); }); }
            sep();
            item('Mover para cima', () => { const ix = idxs; hideCtxMenu(true); moverLinhas(tipo, ix, -1); }, 'Alt+↑');
            item('Mover para baixo', () => { const ix = idxs; hideCtxMenu(true); moverLinhas(tipo, ix, 1); }, 'Alt+↓');
        } else if (vista === 'lista') {
            titulo('Adicionar ao grupo');
            const gs = GL.grupos(dados);
            if (!gs.length) { const d = document.createElement('div'); d.className = 'ctx-r-info'; d.textContent = 'Nenhum grupo nesta tabela ainda. Use «Criar grupo…» ou digite uma linha começando com @ (ex.: @P1).'; sub.appendChild(d); }
            gs.forEach(g => item(GL.rotuloDe(g) || '(sem nome)', () => { const ix = idxs; hideCtxMenu(true); acaoAdicionarAoGrupo(tipo, ix, g.gid); }, GL.contar(dados, g.gid) + ' lin.'));
        } else if (vista === 'criar') {
            titulo('Nome do novo grupo');
            const inp = document.createElement('input'); inp.type = 'text'; inp.className = 'ctx-r-input'; inp.placeholder = 'ex.: P1 Rua das Flores — Enter cria'; inp.autocomplete = 'off';
            inp.addEventListener('input', () => { inp.value = inp.value.toUpperCase(); });
            inp.addEventListener('keydown', e => {
                if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); const nome = inp.value.trim(); if (!nome) return; const ix = idxs; hideCtxMenu(true); acaoCriarGrupo(tipo, ix, nome); }
                else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); hideCtxMenu(); }
                else e.stopPropagation();                                          // as demais teclas são do campo (o menu e a grade não as interceptam)
            });
            sub.appendChild(inp);
            const ok = item('Criar grupo', () => { const nome = inp.value.trim(); if (!nome) { inp.focus(); return; } const ix = idxs; hideCtxMenu(true); acaoCriarGrupo(tipo, ix, nome); });
            ctxMenu.classList.add('sub'); ctxKb = -1; ctxMenu.querySelectorAll('.kb').forEach(el => el.classList.remove('kb'));
            posicionarMenu(ctxMenu, ctxPos.x, ctxPos.y);
            setTimeout(() => inp.focus(), 0);
            return;
        }
        ctxMenu.classList.add('sub');
        posicionarMenu(ctxMenu, ctxPos.x, ctxPos.y);
        ctxKb = -1; ctxMenu.querySelectorAll('.kb').forEach(el => el.classList.remove('kb'));
        if (porTeclado) marcarItemDoMenu(1);                                      // 0 = Voltar; o cursor já cai na 1ª opção
    }
    function acaoAdicionarAoGrupo(tipo, idxs, gid) {
        const GL = window.GruposLinhas, dados = tableStates[tipo].data;
        if (!GL.selecionadas(dados, idxs).length) { showToast('Selecione linhas comuns (a linha «@» é o próprio grupo).'); return; }
        pushHistory();
        const r = GL.adicionarAoGrupo(dados, idxs, gid); if (!r) return;
        aplicarLista(tipo, r.lista, r.foco);
        showToast(`${r.n} linha(s) adicionada(s) ao grupo. Ctrl+Z desfaz.`);
    }
    function acaoCriarGrupo(tipo, idxs, nome) {
        const GL = window.GruposLinhas, dados = tableStates[tipo].data;
        if (!GL.selecionadas(dados, idxs).length) { showToast('Selecione linhas comuns (a linha «@» é o próprio grupo).'); return; }
        pushHistory();
        const op = window.GradeTabela.operacaoHerdada(dados, Math.min.apply(null, idxs));
        const r = GL.criarGrupo(dados, idxs, nome, texto => { const l = window.GradeTabela.novaLinha(tipo, op); l.entidade = '0'; l.ativo = texto.toUpperCase(); return l; });
        if (!r) return;
        aplicarLista(tipo, r.lista, r.foco);
        showToast(`Grupo criado com ${r.n} linha(s). Ctrl+Z desfaz.`);
    }
    function acaoTirarDoGrupo(tipo, idxs) {
        const GL = window.GruposLinhas, dados = tableStates[tipo].data;
        if (!idxs.some(i => dados[i] && dados[i].grupoDe)) return;
        pushHistory();
        const r = GL.tirarDoGrupo(dados, idxs); if (!r) return;
        aplicarLista(tipo, r.lista, r.foco);
        showToast(`${r.n} linha(s) fora do grupo. Ctrl+Z desfaz.`);
    }
    function acaoDesfazerGrupo(tipo, gid) {
        const GL = window.GruposLinhas, dados = tableStates[tipo].data;
        if (!dados.some(r => r.grupoDecl && r.gid === gid)) return;
        pushHistory();
        const r = GL.desfazerGrupo(dados, gid); if (!r) return;
        aplicarLista(tipo, r.lista, r.foco);
        showToast('Grupo desfeito: as linhas continuam na tabela. Ctrl+Z desfaz.');
    }
    document.getElementById('ctx-grupo').addEventListener('click', () => { if (ctxRow && ctxType) abrirVistaGrupo('grupo', ctxKb >= 0); });

    document.getElementById('ctx-add-above').addEventListener('click', () => {
        if (!ctxRow || !ctxType) return;
        const idx = parseInt(ctxRow.dataset.index);
        addRow(ctxType, idx - 1);
        hideCtxMenu();
    });
    document.getElementById('ctx-add-below').addEventListener('click', () => {
        if (!ctxRow || !ctxType) return;
        const idx = parseInt(ctxRow.dataset.index);
        addRow(ctxType, idx);
        hideCtxMenu();
    });
    document.getElementById('ctx-duplicar').addEventListener('click', () => {
        if (!ctxRow || !ctxType) return;
        const tipo = ctxType, idxs = _linhasDoMenu();
        hideCtxMenu();
        duplicarLinhas(tipo, idxs);
    });
    document.getElementById('ctx-congelar').addEventListener('click', () => {
        if (!ctxRow || !ctxType) return;
        const tipo = ctxType, idxs = _linhasDoMenu();
        hideCtxMenu();
        alternarMarcaLinhas(tipo, idxs, 'congelada');
    });
    document.getElementById('ctx-comentario').addEventListener('click', () => {
        if (!ctxRow || !ctxType) return;
        const tipo = ctxType, idxs = _linhasDoMenu();
        hideCtxMenu();
        alternarComentario(tipo, idxs);
    });
    document.getElementById('ctx-bloquear').addEventListener('click', () => {
        if (!ctxRow || !ctxType) return;
        const tipo = ctxType, idxs = _linhasDoMenu();
        hideCtxMenu();
        alternarMarcaLinhas(tipo, idxs, 'bloqueada');
    });
    document.getElementById('ctx-del').addEventListener('click', () => {
        if (!ctxRow || !ctxType) return;
        deleteRow(ctxType, parseInt(ctxRow.dataset.index));
        hideCtxMenu();
    });

    /* ═══════════════════════════════════════
       CLASSIFICAR / AGRUPAR POR COLUNA (TASK-069)
       Ordena de verdade (estável, natural no ativo, ordem lógica nas operações), com um passo de Ctrl+Z. Em Outros os vínculos
       cabo<->estrutura (por índice) são remapeados; em Cabos a Qtd Ativos NÃO é recalculada (fica com a linha).
    ═══════════════════════════════════════ */
    function ordenarTabela(tipo, criterios) {
        const GT = window.GradeTabela;
        const dados = tableStates[tipo].data;
        if (dados.length < 2) { showToast('Nada para ordenar.'); return; }
        pushHistory();
        const ordem = GT.ordenarLinhas(dados, criterios);
        let mapa = GT.mapaDeOrdem(ordem);
        tableStates[tipo].data = ordem.map(i => dados[i]);
        if (window.GruposLinhas && dados.some(r => r.grupoDecl || r.grupoDe)) {                 // TASK-195: com grupos, os membros continuam colados na declaração
            tableStates[tipo].data = window.GruposLinhas.normalizar(tableStates[tipo].data);
            mapa = window.GruposLinhas.mapa(dados, tableStates[tipo].data);
        }
        if (tipo === 'outros') {
            tableStates.cabos.data.forEach(c => {
                if (Array.isArray(c.vinculoPostes)) c.vinculoPostes = c.vinculoPostes.map(l => (mapa[l] !== undefined ? mapa[l] : l));
                if (Array.isArray(c.vinculoEstruturas)) c.vinculoEstruturas = c.vinculoEstruturas.map(i => { const l = window.VinculacaoRegras.linhaDe(i); return mapa[l] !== undefined ? window.VinculacaoRegras.idUnidade(mapa[l], window.VinculacaoRegras.slotDe(i)) : i; });
            });
        }
        ordemAtual[tipo] = criterios.map(c => Object.assign({}, c));
        limparGrade();
        document.querySelectorAll('tr.row-selected').forEach(x => x.classList.remove('row-selected'));
        renderTable(tipo);
        buildAtivoSets(); buildDataLists();
        refreshAllFilters(tipo);
        updateHistoryUI();
        showToast(`${tipo === 'cabos' ? 'Cabos' : 'Outros'} classificado por ${criterios.map(c => c.campo + (c.dir === 'desc' ? ' ↓' : ' ↑')).join(', ')}. Ctrl+Z desfaz.`
            + (tipo === 'cabos' ? ' Qtd Ativos não foi recalculada.' : ''));
    }

    function atualizarIndicadoresOrdem(tipo) {
        const GT = window.GradeTabela;
        if (ordemAtual[tipo].length && !(GT && GT.estaOrdenado(tableStates[tipo].data, ordemAtual[tipo]))) ordemAtual[tipo] = [];   // edição/undo desfez a ordem
        document.querySelectorAll(`.th-sort[data-tipo="${tipo}"]`).forEach(th => {
            const i = ordemAtual[tipo].findIndex(c => c.campo === th.dataset.campo);
            const ind = th.querySelector('.sort-ind');
            if (!ind) return;
            if (i >= 0) ind.textContent = (ordemAtual[tipo][i].dir === 'desc' ? '▼' : '▲') + (ordemAtual[tipo].length > 1 ? (i + 1) : '');
            else ind.textContent = agrupamento[tipo] === th.dataset.campo ? '▤' : '';
        });
    }

    /** Clique no título da coluna: 1º clique crescente, de novo decrescente, e assim por diante. Shift = critério adicional. */
    function alternarOrdem(tipo, campo, shift) {
        const atual = ordemAtual[tipo].map(c => Object.assign({}, c));
        const i = atual.findIndex(c => c.campo === campo);
        const dir = i >= 0 && atual[i].dir === 'asc' ? 'desc' : 'asc';
        let crit;
        if (shift && atual.length) { if (i >= 0) { atual[i].dir = dir; crit = atual; } else crit = atual.concat([{ campo, dir }]); }
        else crit = [{ campo, dir }];
        ordenarTabela(tipo, crit);
    }

    (function ligarCabecalhosClassificar() {
        const menu = document.createElement('div');
        menu.id = 'menu-classificar';
        document.body.appendChild(menu);
        const fechar = () => menu.classList.remove('visible');
        document.addEventListener('click', e => {
            const th = e.target.closest ? e.target.closest('.th-sort') : null;
            if (th) { fechar(); alternarOrdem(th.dataset.tipo, th.dataset.campo, e.shiftKey); return; }
            if (!menu.contains(e.target)) fechar();
        });
        // botão direito no título: agrupar (visual) / limpar agrupamento
        document.addEventListener('contextmenu', e => {
            const th = e.target.closest ? e.target.closest('.th-sort') : null;
            if (!th) return;
            e.preventDefault();
            const tipo = th.dataset.tipo, campo = th.dataset.campo;
            menu.innerHTML = '';
            const item = (txt, acao) => { const d = document.createElement('div'); d.className = 'mc-item'; d.textContent = txt; d.addEventListener('click', () => { fechar(); acao(); }); menu.appendChild(d); };
            if (campo !== 'qtdAtivos') item('Agrupar por esta coluna', () => { agrupamento[tipo] = campo; ordenarTabela(tipo, [{ campo, dir: 'asc' }]); });
            if (agrupamento[tipo]) item('Limpar agrupamento', () => { agrupamento[tipo] = null; renderTable(tipo); });
            if (!menu.children.length) { fechar(); return; }
            menu.classList.add('visible');
            posicionarMenu(menu, e.clientX, e.clientY + 4, e.clientY - 4);
        });
        document.addEventListener('keydown', e => { if (e.key === 'Escape') fechar(); });
    })();

    // TASK-092: Fases/Tensão/Nível dos Cabos aparecem por padrão e podem ser ocultadas (preferência lembrada)
    (function ligarOcultarCaboExtra() {
        const btn = document.getElementById('btn-toggle-cabo-extra'), tabela = document.getElementById('table-cabos');
        if (!btn || !tabela) return;
        const aplicar = (oculto, persistir) => {
            tabela.classList.toggle('sem-cabo-extra', oculto);
            ajustarCelulasDeLarguraTotal(tabela);
            btn.textContent = oculto ? 'Mostrar Fases/Nível' : 'Ocultar Fases/Nível';
            btn.title = oculto ? 'Mostrar as colunas Fases / Tensão / Nível' : 'Ocultar as colunas Fases / Tensão / Nível';
            if (persistir) { try { localStorage.setItem('resumo_ocultar_cabo_extra', oculto ? '1' : '0'); } catch (e) { /* vale só nesta sessão */ } }
        };
        let inicial = false;
        try { inicial = localStorage.getItem('resumo_ocultar_cabo_extra') === '1'; } catch (e) { /* ignora */ }
        aplicar(inicial, false);
        btn.addEventListener('click', () => aplicar(!tabela.classList.contains('sem-cabo-extra'), true));
    })();

    // TASK-088: a coluna Entidade de Outros começa OCULTA (a classificação continua; só não aparece). Menu "Exibir" reativa a coluna e classifica/agrupa por entidade.
    (function ligarOcultarEntidade() {
        const btn = document.getElementById('btn-toggle-entidade'), tabela = document.getElementById('table-outros');
        if (!btn || !tabela) return;
        const aplicar = (oculto, persistir) => {
            tabela.classList.toggle('sem-entidade', oculto);
            btn.textContent = 'Exibir ▾';
            btn.title = 'Exibir: mostrar a coluna Entidade, classificar ou agrupar por entidade';
            if (oculto && filterSelections.outros.entidade.size) { filterSelections.outros.entidade.clear(); refreshAllFilters('outros'); applyFilters('outros'); }   // filtro invisível não pode ficar ativo
            ajustarCelulasDeLarguraTotal(tabela);
            if (persistir) { try { localStorage.setItem('resumo_ocultar_entidade', oculto ? '1' : '0'); } catch (e) { /* sem localStorage: vale só nesta sessão */ } }
        };
        let inicial = true;                                           // padrão: oculta; quem reativou (preferência '0') continua vendo
        try { inicial = localStorage.getItem('resumo_ocultar_entidade') !== '0'; } catch (e) { /* ignora */ }
        aplicar(inicial, false);
        const menu = document.createElement('div');
        menu.id = 'menu-exibir-outros'; menu.className = '';
        menu.style.cssText = 'position:fixed;z-index:9000;display:none;background:#161b22;border:1px solid #30363d;border-radius:8px;padding:5px;min-width:230px;box-shadow:0 8px 24px rgba(0,0,0,.5);font-size:.8rem;';
        document.body.appendChild(menu);
        const fechar = () => { menu.style.display = 'none'; };
        const item = (txt, acao) => { const d = document.createElement('div'); d.textContent = txt; d.style.cssText = 'padding:6px 10px;border-radius:5px;cursor:pointer;color:#c9d1d9;'; d.onmouseenter = () => { d.style.background = '#21262d'; }; d.onmouseleave = () => { d.style.background = ''; }; d.addEventListener('click', () => { fechar(); acao(); }); menu.appendChild(d); };
        btn.addEventListener('click', e => {
            e.stopPropagation();
            menu.innerHTML = '';
            const oculto = tabela.classList.contains('sem-entidade');
            item(oculto ? 'Mostrar a coluna Entidade' : 'Ocultar a coluna Entidade', () => aplicar(!oculto, true));
            item('Classificar por Entidade (A→Z)', () => ordenarTabela('outros', [{ campo: 'entidade', dir: 'asc' }]));
            item('Classificar por Entidade (Z→A)', () => ordenarTabela('outros', [{ campo: 'entidade', dir: 'desc' }]));
            item('Agrupar por Entidade', () => { agrupamento.outros = 'entidade'; ordenarTabela('outros', [{ campo: 'entidade', dir: 'asc' }]); });
            if (agrupamento.outros) item('Limpar agrupamento', () => { agrupamento.outros = null; renderTable('outros'); });
            const r = btn.getBoundingClientRect();
            menu.style.display = 'block'; posicionarMenu(menu, r.left, r.bottom + 4, r.top - 4);
        });
        document.addEventListener('click', e => { if (!menu.contains(e.target)) fechar(); });
        document.addEventListener('keydown', e => { if (e.key === 'Escape') fechar(); });
    })();

    /* ═══════════════════════════════════════
       GRADE ESTILO EXCEL (TASK-067) — Cabos e Outros
       Seleção de intervalo (arrastar, Shift, nº da linha, Ctrl+A), copiar/colar em TSV, Delete.
       A lógica pura está em static/grade_tabela.js; aqui só a ligação com a tela e o histórico.
    ═══════════════════════════════════════ */
    const G = window.GradeTabela;

    function colunasGrade(tipo) { return G.colunasDe(tipo); }
    function bodyDe(tipo) { return document.getElementById(tableStates[tipo].bodyId); }
    function linhaVisivel(tr) { return tr && tr.style.display !== 'none'; }

    function intervaloGrade() {
        if (!grade.tipo || !grade.ancora || !grade.foco) return null;
        return {
            tipo: grade.tipo,
            r1: Math.min(grade.ancora.r, grade.foco.r), r2: Math.max(grade.ancora.r, grade.foco.r),
            c1: Math.min(grade.ancora.c, grade.foco.c), c2: Math.max(grade.ancora.c, grade.foco.c)
        };
    }
    function multiCelulas(iv) { return !!iv && (iv.r2 > iv.r1 || iv.c2 > iv.c1); }

    function pintarGrade() {
        document.querySelectorAll('td.cel-sel').forEach(td => td.classList.remove('cel-sel'));
        const iv = intervaloGrade();
        if (!iv || !multiCelulas(iv)) return;
        Array.from(bodyDe(iv.tipo).children).forEach(tr => {
            const r = parseInt(tr.dataset.index);
            if (isNaN(r) || r < iv.r1 || r > iv.r2 || !linhaVisivel(tr)) return;
            tr.querySelectorAll('td[data-col]').forEach(td => {
                const c = parseInt(td.dataset.col);
                if (c >= iv.c1 && c <= iv.c2) td.classList.add('cel-sel');
            });
        });
    }
    function limparGrade() {
        if (!grade.tipo) return;
        grade.tipo = grade.ancora = grade.foco = null;
        pintarGrade();
    }

    function celulaDe(el) {
        const td = el && el.closest ? el.closest('td[data-col]') : null;
        if (!td) return null;
        const tr = td.parentElement;
        if (!tr || tr.dataset.index === undefined || !tableStates[tr.dataset.type]) return null;
        return { tipo: tr.dataset.type, r: parseInt(tr.dataset.index), c: parseInt(td.dataset.col) };
    }

    function ehCampoDeTexto(el) {
        return !!el && (el.tagName === 'TEXTAREA' || (el.tagName === 'INPUT' && /^(text|search|number|)$/.test(el.type)));
    }
    function dentroDaGrade(el) { return !!(el && el.closest && el.closest('#body-cabos, #body-outros')); }

    function focarCelula(tipo, r, c) {
        const body = bodyDe(tipo), n = tableStates[tipo].data.length;
        if (r < 0 || r >= n) return false;
        const tr = Array.from(body.children).find(x => parseInt(x.dataset.index) === r);
        if (!tr) return false;
        hidratarLinha(tr);
        const td = tr.querySelector(`td[data-col="${c}"]`);
        const campo = td && td.querySelector('input,select');
        if (campo) { campo.focus(); if (campo.select && campo.tagName === 'INPUT') campo.select(); return true; }
        return false;
    }
    function irParaCelula(tipo, r, c, passo) {
        const n = tableStates[tipo].data.length;
        const body = bodyDe(tipo);
        if (passo > 0 && r >= n && c === 1) return focarLinhaVazia(tipo, 0);              // depois da última linha preenchida: a primeira linha em branco
        while (r >= 0 && r < n) {      // pula linhas escondidas por filtro
            const tr = Array.from(body.children).find(x => parseInt(x.dataset.index) === r);
            if (linhaVisivel(tr)) return focarCelula(tipo, r, c);
            r += passo;
        }
        return false;
    }

    function textoDaSelecao() {
        const iv = intervaloGrade();
        if (!iv) return '';
        const body = bodyDe(iv.tipo), partes = [];
        for (let r = iv.r1; r <= iv.r2; r++) {
            const tr = Array.from(body.children).find(x => parseInt(x.dataset.index) === r);
            if (tr && !linhaVisivel(tr)) continue;
            partes.push(G.textoDoIntervalo(tableStates[iv.tipo].data, iv.tipo, r, iv.c1, r, iv.c2));
        }
        return partes.join('\n');
    }

    function finalizarEdicaoEmLote(tipo) {
        if (tipo === 'cabos') recalcAllQtdAtivos();
        renderTable(tipo);
        buildAtivoSets(); buildDataLists();
        refreshAllFilters(tipo);
        updateHistoryUI();
    }

    function colarNaGrade(tipo, linha, coluna, texto) {
        const dados = tableStates[tipo].data;
        pushHistory();      // um único passo de Ctrl+Z para a colagem inteira
        const r = G.aplicarColagem(dados, texto, { tipo, linha, coluna });
        if (!r.linhasColadas) { showToast('Nada para colar.'); return; }
        r.alteradas.forEach(i => { if (tipo === 'outros') reclassificarEntidade(dados[i]); else dados[i].entidade = 'CABO'; sincronizarHash(dados[i], ''); });
        finalizarEdicaoEmLote(tipo);
        const cols = colunasGrade(tipo);
        grade.tipo = tipo;
        grade.ancora = { r: linha, c: Math.min(coluna, cols.length - 1) };
        grade.foco = { r: r.alteradas[r.alteradas.length - 1], c: Math.min(cols.length - 1, Math.max(grade.ancora.c, coluna)) };
        pintarGrade();
        let msg = `${r.linhasColadas} linha(s) coladas` + (r.criadas ? `, ${r.criadas} nova(s)` : '') + '. Ctrl+Z desfaz.';
        if (r.opsInvalidas) msg += ` ${r.opsInvalidas} operação(ões) inválida(s) mantida(s) como estavam.`;
        if (r.truncado) msg += ` Limite de ${G.MAX_LINHAS} linhas: o restante foi ignorado.`;
        showToast(msg);
    }

    function limparSelecaoGrade() {
        const iv = intervaloGrade();
        if (!iv) return;
        pushHistory();
        const alt = G.limparIntervalo(tableStates[iv.tipo].data, iv.tipo, iv.r1, iv.c1, iv.r2, iv.c2);
        if (!alt.length) { showToast('Nada para limpar (só o Ativo é esvaziado).'); return; }
        finalizarEdicaoEmLote(iv.tipo);
        pintarGrade();
    }

    function selecionarTudoGrade(tipo) {
        const n = tableStates[tipo].data.length;
        if (!n) return;
        grade.tipo = tipo;
        grade.ancora = { r: 0, c: 0 };
        grade.foco = { r: n - 1, c: colunasGrade(tipo).length - 1 };
        pintarGrade();
    }

    // Nº da linha: seleciona a linha toda (Ctrl alterna, Shift estende) e marca para o botão Excluir
    function clicarNumeroLinha(td, e) {
        const tr = td.parentElement, tipo = tr.dataset.type, r = parseInt(tr.dataset.index);
        const body = bodyDe(tipo);
        const marcar = (a, b) => Array.from(body.children).forEach(x => {
            const i = parseInt(x.dataset.index);
            if (i >= Math.min(a, b) && i <= Math.max(a, b) && linhaVisivel(x)) x.classList.add('row-selected');
        });
        if (e.shiftKey && grade.ultimaLinha && grade.ultimaLinha.tipo === tipo) {
            document.querySelectorAll('tr.row-selected').forEach(x => x.classList.remove('row-selected'));
            marcar(grade.ultimaLinha.r, r);
        } else if (e.ctrlKey || e.metaKey) {
            tr.classList.toggle('row-selected');
            grade.ultimaLinha = { tipo, r };
        } else {
            document.querySelectorAll('tr.row-selected').forEach(x => x.classList.remove('row-selected'));
            tr.classList.add('row-selected');
            grade.ultimaLinha = { tipo, r };
        }
        // a seleção de células acompanha: linhas inteiras
        const sel = Array.from(body.querySelectorAll('tr.row-selected')).map(x => parseInt(x.dataset.index));
        grade.tipo = tipo;
        grade.ancora = { r: Math.min(...sel), c: 0 };
        grade.foco = { r: Math.max(...sel), c: colunasGrade(tipo).length - 1 };
        pintarGrade();
    }

    document.addEventListener('mousedown', e => {
        if (e.button !== 0) return;
        const num = e.target.closest ? e.target.closest('td.col-num') : null;
        if (num && num.parentElement.dataset.index !== undefined) { e.preventDefault(); clicarNumeroLinha(num, e); return; }
        const cel = celulaDe(e.target);
        if (!cel) {
            // clique fora da grade: some o intervalo (as linhas marcadas pelo nº seguem marcadas até o próximo clique numa célula/nº)
            if (!(e.target.closest && e.target.closest('#ctx-menu-resumo, #modal-ler-projeto, #modal-definir-entidade, .modal-overlay, .section-actions, button'))) limparGrade();
            return;
        }
        document.querySelectorAll('tr.row-selected').forEach(x => x.classList.remove('row-selected'));
        if (e.shiftKey && grade.tipo === cel.tipo && grade.ancora) {
            e.preventDefault();
            grade.foco = { r: cel.r, c: cel.c };
            if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
            pintarGrade();
            return;
        }
        grade.tipo = cel.tipo;
        grade.ancora = { r: cel.r, c: cel.c };
        grade.foco = { r: cel.r, c: cel.c };
        grade.arrastando = true;
        pintarGrade();
    });
    document.addEventListener('mousemove', e => {
        if (!grade.arrastando || !(e.buttons & 1)) { grade.arrastando = grade.arrastando && !!(e.buttons & 1); return; }
        const cel = celulaDe(e.target);
        if (!cel || cel.tipo !== grade.tipo) return;
        if (cel.r === grade.foco.r && cel.c === grade.foco.c) return;
        grade.foco = { r: cel.r, c: cel.c };
        if (multiCelulas(intervaloGrade())) {
            e.preventDefault();
            document.body.classList.add('grade-arrastando');
            const sel = window.getSelection && window.getSelection();
            if (sel) sel.removeAllRanges();
            if (document.activeElement && dentroDaGrade(document.activeElement)) document.activeElement.blur();
        }
        pintarGrade();
    });
    document.addEventListener('mouseup', () => { grade.arrastando = false; document.body.classList.remove('grade-arrastando'); });

    const modalRamaisAberto = () => !!document.querySelector('.modal-overlay:not(.hidden)');   // qualquer modal aberto tem as suas próprias teclas

    document.addEventListener('copy', e => {
        if (modalRamaisAberto()) return;
        const iv = intervaloGrade();
        if (!iv) return;
        const ativo = document.activeElement;
        if (!multiCelulas(iv) && ehCampoDeTexto(ativo)) return;             // uma célula em edição: cópia nativa do texto
        if (ativo && !dentroDaGrade(ativo) && ativo !== document.body && ehCampoDeTexto(ativo)) return;
        const t = textoDaSelecao();
        if (!t) return;
        e.preventDefault();
        e.clipboardData.setData('text/plain', t);
        showToast('Copiado (' + t.split('\n').length + ' linha(s)).');
    });

    document.addEventListener('paste', e => {
        if (modalRamaisAberto()) return;
        const alvo = e.target;
        const cd = e.clipboardData || window.clipboardData;
        const texto = cd ? cd.getData('text') : '';
        if (!texto) return;
        const iv = intervaloGrade();
        const naGrade = dentroDaGrade(alvo);
        const corpo = alvo === document.body || alvo === document.documentElement;
        if (!naGrade && !(corpo && iv)) return;
        if (!G.ehColagemMultipla(texto) && !(corpo && iv)) return;          // valor único num campo: colagem nativa
        let tipo, linha, coluna;
        const cel = naGrade ? celulaDe(alvo) : null;
        const dentroDoIntervalo = !!(iv && cel && cel.tipo === iv.tipo && cel.r >= iv.r1 && cel.r <= iv.r2 && cel.c >= iv.c1 && cel.c <= iv.c2);
        if (iv && (!cel || (multiCelulas(iv) && dentroDoIntervalo))) { tipo = iv.tipo; linha = iv.r1; coluna = iv.c1; }
        else if (cel) { tipo = cel.tipo; linha = cel.r; coluna = cel.c; }
        else return;
        e.preventDefault();
        colarNaGrade(tipo, linha, coluna, texto);
    });

    document.addEventListener('keydown', e => {
        if (modalRamaisAberto()) return;
        const ativo = document.activeElement;
        const noCampo = ehCampoDeTexto(ativo);
        const emGrade = dentroDaGrade(ativo) || ativo === document.body;
        if (!emGrade) return;
        const iv = intervaloGrade();
        const mod = e.ctrlKey || e.metaKey;

        if (mod && (e.key === 'a' || e.key === 'A')) {
            let tipo = iv ? iv.tipo : null;
            const cel = celulaDe(ativo);
            if (!tipo && cel) tipo = cel.tipo;
            if (!tipo) return;
            if (noCampo && ativo.value && !(ativo.selectionStart === 0 && ativo.selectionEnd === ativo.value.length)) return;  // 1º Ctrl+A: texto do campo
            e.preventDefault();
            selecionarTudoGrade(tipo);
            if (ativo && ativo.blur && dentroDaGrade(ativo)) ativo.blur();
            return;
        }
        if (e.key === 'Escape' && iv) { limparGrade(); return; }

        if ((e.key === 'Delete' || e.key === 'Backspace') && iv && (multiCelulas(iv) || !noCampo) && !(ativo && ativo.tagName === 'SELECT')) {
            e.preventDefault();
            limparSelecaoGrade();
            return;
        }

        if (e.shiftKey && !mod && /^Arrow(Up|Down|Left|Right)$/.test(e.key)) {
            if (noCampo && /Left|Right/.test(e.key)) return;                    // Shift+←/→ dentro do campo seleciona texto
            if (ativo && ativo.tagName === 'SELECT') return;
            let base = iv;
            if (!base) {
                const cel = celulaDe(ativo);
                if (!cel) return;
                grade.tipo = cel.tipo; grade.ancora = { r: cel.r, c: cel.c }; grade.foco = { r: cel.r, c: cel.c };
                base = intervaloGrade();
            }
            e.preventDefault();
            const n = tableStates[grade.tipo].data.length, nc = colunasGrade(grade.tipo).length;
            const f = grade.foco;
            if (e.key === 'ArrowUp') f.r = Math.max(0, f.r - 1);
            if (e.key === 'ArrowDown') f.r = Math.min(n - 1, f.r + 1);
            if (e.key === 'ArrowLeft') f.c = Math.max(0, f.c - 1);
            if (e.key === 'ArrowRight') f.c = Math.min(nc - 1, f.c + 1);
            if (ativo && ativo.blur && dentroDaGrade(ativo)) ativo.blur();
            pintarGrade();
            return;
        }
        // sem foco em campo (depois de selecionar por arraste): setas movem para a célula
        if (ativo === document.body && iv && !mod && /^Arrow/.test(e.key)) {
            e.preventDefault();
            const f = grade.foco, n = tableStates[iv.tipo].data.length, nc = colunasGrade(iv.tipo).length;
            let r = f.r, c = f.c;
            if (e.key === 'ArrowUp') r = Math.max(0, r - 1);
            if (e.key === 'ArrowDown') r = Math.min(n - 1, r + 1);
            if (e.key === 'ArrowLeft') c = Math.max(0, c - 1);
            if (e.key === 'ArrowRight') c = Math.min(nc - 1, c + 1);
            grade.ancora = { r, c }; grade.foco = { r, c };
            pintarGrade();
            focarCelula(iv.tipo, r, c);
        }
    });

    // Excluir (botão): linhas marcadas pelo nº da linha, senão as linhas do intervalo selecionado, senão a última
    function linhasDoIntervaloSelecionado(tipo) {
        const iv = intervaloGrade();
        if (!iv || iv.tipo !== tipo || !multiCelulas(iv)) return [];
        const body = bodyDe(tipo), out = [];
        Array.from(body.children).forEach(tr => {
            const r = parseInt(tr.dataset.index);
            if (r >= iv.r1 && r <= iv.r2 && linhaVisivel(tr)) out.push(r);
        });
        return out;
    }

    // Entidade em Outros: correção manual pelo menu de contexto (TASK-065)
    function pedirEntidade(qtd, atual) {
        return new Promise(resolve => {
            const fundo = document.createElement('div');
            fundo.id = 'modal-definir-entidade';
            fundo.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:10000;display:flex;align-items:center;justify-content:center;';
            const caixa = document.createElement('div');
            caixa.style.cssText = 'background:var(--bg-canvas,#0d1117);color:var(--text-primary,#e6edf3);border:1px solid var(--border-default,#30363d);border-radius:10px;padding:20px;max-width:420px;width:92%;font-family:Inter,sans-serif;';
            const t = document.createElement('h3'); t.textContent = 'Definir entidade'; t.style.margin = '0 0 8px';
            const p = document.createElement('p'); p.style.cssText = 'margin:0 0 12px;font-size:.85rem;color:#8b949e;line-height:1.4;';
            p.textContent = `${qtd} linha(s). A entidade escolhida fica travada (não é recalculada ao editar o ativo). "Automática" destrava e recalcula pelas regras do Leitor.`;
            const sel = document.createElement('select');
            sel.style.cssText = 'width:100%;padding:6px;margin-bottom:14px;background:#1c1f26;color:#c9d1d9;border:1px solid #30363d;border-radius:6px;';
            const auto = document.createElement('option'); auto.value = '__auto__'; auto.textContent = 'Automática (destravar)'; sel.appendChild(auto);
            ENTIDADES.filter(x => x !== 'CABO').forEach(x => { const o = document.createElement('option'); o.value = x; o.textContent = x === '0' ? '0 (sem classificação)' : x; if (x === atual) o.selected = true; sel.appendChild(o); });
            const linha = document.createElement('div'); linha.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;';
            const fim = v => { fundo.remove(); resolve(v); };
            const bc = document.createElement('button'); bc.className = 'btn-secondary'; bc.textContent = 'Cancelar'; bc.onclick = () => fim(null);
            const bo = document.createElement('button'); bo.className = 'btn-primary'; bo.textContent = 'Aplicar'; bo.onclick = () => fim(sel.value);
            linha.append(bc, bo);
            caixa.append(t, p, sel, linha); fundo.appendChild(caixa); document.body.appendChild(fundo);
            sel.focus();
        });
    }
    document.getElementById('ctx-entidade').addEventListener('click', async () => {
        if (!ctxRow || ctxType !== 'outros') return;
        const idxClicado = parseInt(ctxRow.dataset.index);
        let alvos = linhasDoIntervaloSelecionado('outros');
        const marcadas = Array.from(bodyDe('outros').querySelectorAll('tr.row-selected')).map(x => parseInt(x.dataset.index));
        if (marcadas.length) alvos = marcadas;
        if (!alvos.includes(idxClicado)) alvos = [idxClicado];
        hideCtxMenu();
        const dados = tableStates.outros.data;
        const v = await pedirEntidade(alvos.length, dados[alvos[0]] && dados[alvos[0]].entidade);
        if (v === null) return;
        pushHistory();
        alvos.forEach(i => {
            const row = dados[i];
            if (!row) return;
            if (v === '__auto__') { delete row.entidadeManual; reclassificarEntidade(row); }
            else { row.entidade = v; row.entidadeManual = true; }
        });
        renderTable('outros');
        refreshAllFilters('outros');
        updateHistoryUI();
        showToast(v === '__auto__' ? 'Entidade automática.' : `Entidade ${v} definida em ${alvos.length} linha(s). Ctrl+Z desfaz.`);
    });

    /* ═══════════════════════════════════════
       AUTOCOMPLETE
    ═══════════════════════════════════════ */
    function buildAtivoSets() {
        ativoSets.cabos.clear();
        ativoSets.outros.clear();
        tableStates.cabos.data.forEach(r => { if (r.ativo && !r.congelada) ativoSets.cabos.add(r.ativo.trim()); });
        tableStates.outros.data.forEach(r => { if (r.ativo && !r.congelada) ativoSets.outros.add(r.ativo.trim()); });
    }

    function buildDataLists() {
        buildDl('dl-ativos-cabos', ativoSets.cabos);
        buildDl('dl-ativos-outros', ativoSets.outros);
    }

    function buildDl(id, set) {
        const dl = document.getElementById(id);
        dl.innerHTML = '';
        set.forEach(v => {
            const opt = document.createElement('option');
            opt.value = v; dl.appendChild(opt);
        });
    }

    function showAutocomplete(input, type) {
        hideAutocomplete();
        // Só opções de ATIVOS da base (nunca as linhas já digitadas na tabela, que confundiam) e só enquanto digita o nome (uma palavra); em Outros quem sugere é o AtivoSugestao
        const q = input.value.trim().toUpperCase();
        if (!q || type !== 'cabos' || !/^\S+$/.test(q) || !window.AtivoSugestao) return;
        const proj = localStorage.getItem('projeto_selecionado') || '';
        acItems = window.AtivoSugestao.sugerir(q, window.AtivoSugestao.indiceDeNomes(orcamentoBaseData, proj, 'CABOS'), 12).map(e => e.nome);
        if (acItems.length === 0) return;

        acList = document.createElement('div');
        acList.className = 'autocomplete-list';
        acSelected = -1;
        acInput = input;

        acItems.forEach((item, i) => {
            const div = document.createElement('div');
            div.className = 'autocomplete-item';
            div.textContent = item;
            div.addEventListener('mousedown', e => {
                e.preventDefault();
                input.value = item;
                input.dispatchEvent(new Event('input'));
                hideAutocomplete();
                input.focus();
            });
            div.addEventListener('mouseover', () => setAcSelected(i));
            acList.appendChild(div);
        });

        // Posiciona relativo ao input
        const rect = input.getBoundingClientRect();
        acList.style.position = 'fixed';
        acList.style.left = rect.left + 'px';
        acList.style.top  = (rect.bottom + 2) + 'px';
        acList.style.width = Math.max(160, rect.width) + 'px';
        document.body.appendChild(acList);
    }

    function hideAutocomplete() {
        if (acList) { acList.remove(); acList = null; acItems = []; acSelected = -1; acInput = null; }
    }

    function setAcSelected(idx) {
        if (!acList) return;
        acSelected = idx;
        Array.from(acList.children).forEach((el, i) => el.classList.toggle('active', i === idx));
    }

    function handleAcKeydown(e, input, type) {
        if (!acList) {
            if (e.key === 'ArrowDown' && type !== 'outros') { showAutocomplete(input, type); e.preventDefault(); }
            return;
        }
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            setAcSelected(Math.min(acSelected + 1, acItems.length - 1));
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            setAcSelected(Math.max(acSelected - 1, 0));
        } else if (e.key === 'Enter' || e.key === 'Tab') {
            if (acSelected >= 0 && acItems[acSelected]) {
                e.preventDefault();
                input.value = acItems[acSelected];
                input.dispatchEvent(new Event('input'));
                hideAutocomplete();
                return true;
            } else { hideAutocomplete(); }
        } else if (e.key === 'Escape') {
            hideAutocomplete();
        }
    }

    /* ═══════════════════════════════════════
       UNDO / REDO
    ═══════════════════════════════════════ */
    function snapshotState() {
        return {
            cabos: tableStates.cabos.data.map(deepClone),
            outros: tableStates.outros.data.map(deepClone)
        };
    }

    /* Histórico (TASK-021). Os chamadores empilham ANTES de mudar (a maioria) ou DEPOIS (poucos); como o estado ao vivo
       pode estar à frente do ponteiro, `historyDirty` marca "pode haver mudança não gravada" e o undo a grava antes
       de voltar — assim cada ação é exatamente um passo, qualquer que seja o estilo do chamador. */
    const mesmoEstado = (a, b) => JSON.stringify([a && a.cabos, a && a.outros]) === JSON.stringify([b && b.cabos, b && b.outros]);      // TASK-117: `modais` (extra do Limpar Tudo) não conta

    function pushHistory() {
        const snap = snapshotState();
        if (!(historyIdx >= 0 && mesmoEstado(history[historyIdx], snap))) {
            // Descarta redo futuro
            if (historyIdx < history.length - 1) history = history.slice(0, historyIdx + 1);
            history.push(snap);
            if (history.length > MAX_HISTORY) history.shift();
            historyIdx = history.length - 1;
        }
        historyDirty = true;
        updateHistoryUI();
        marcarProjetoAlterado();                                       // TASK-182: congelar/bloquear/comentar e demais ações do histórico também gravam no projeto
    }

    /** Grava o estado ao vivo se ele estiver à frente do ponteiro (mudança feita depois do último push). */
    function gravarEstadoAoVivo() {
        const vivo = snapshotState();
        if (historyIdx >= 0 && !mesmoEstado(history[historyIdx], vivo)) {
            history = history.slice(0, historyIdx + 1);
            history.push(vivo);
            if (history.length > MAX_HISTORY) history.shift();
            historyIdx = history.length - 1;
        }
        historyDirty = false;
    }

    function undo() {
        gravarEstadoAoVivo();
        if (historyIdx <= 0) { updateHistoryUI(); return; }
        if (history[historyIdx - 1].modais && !history[historyIdx].modais) history[historyIdx].modais = coletarModaisDaObra();      // TASK-117: o Refazer devolve o estado atual dos modais
        historyIdx--;
        applyUndoRedoSnapshot(history[historyIdx]);
        showToast('Desfeito');
    }

    function redo() {
        if (historyIdx >= history.length - 1) return;
        historyDirty = false;
        historyIdx++;
        applyUndoRedoSnapshot(history[historyIdx]);
        showToast('Refeito');
    }

    function applyUndoRedoSnapshot(snap) {
        tableStates.cabos.data  = snap.cabos.map(deepClone);
        tableStates.outros.data = snap.outros.map(deepClone);
        if (snap.modais && typeof restaurarModaisDaObra === 'function') restaurarModaisDaObra(snap.modais);      // TASK-117: Ramais/Conexões voltam com o Limpar Tudo
        renderTable('cabos');
        renderTable('outros');
        buildAtivoSets();
        buildDataLists();
        refreshAllFilters('cabos');
        refreshAllFilters('outros');
        updateHistoryUI();
        marcarProjetoAlterado();                                       // desfazer/refazer também são alterações do projeto
    }

    function updateHistoryUI() {
        const btnU = document.getElementById('btn-undo');
        const btnR = document.getElementById('btn-redo');
        const info = document.getElementById('history-info');
        btnU.disabled = historyIdx <= 0 && !historyDirty;
        btnR.disabled = historyIdx >= history.length - 1;
        const pos = historyIdx + 1, tot = history.length;
        info.textContent = tot > 1 ? `${pos}/${tot}` : '';
    }

    document.getElementById('btn-undo').addEventListener('click', undo);
    document.getElementById('btn-redo').addEventListener('click', redo);

    /* ═══════════════════════════════════════
       TECLADO GLOBAL
    ═══════════════════════════════════════ */
    document.addEventListener('keydown', e => {
        const tag = document.activeElement.tagName;
        const isEditing = (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || document.activeElement.contentEditable === 'true');
        if (document.querySelector('.modal-overlay:not(.hidden)')) return;   // com um modal aberto, Ctrl+Z/Y não mexem nas tabelas de trás

        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) {
            if (!isEditing) { e.preventDefault(); undo(); }
            return;
        }
        if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'y' || (e.key.toLowerCase() === 'z' && e.shiftKey))) {
            if (!isEditing) { e.preventDefault(); redo(); }
            return;
        }
        if (e.key === 'Escape') hideAutocomplete();
    });

    /* ═══════════════════════════════════════
       COUNTERS
    ═══════════════════════════════════════ */
    function updateCounters() {
        // Conta apenas as linhas visíveis (não filtradas)
        ['cabos', 'outros'].forEach(type => {
            const body = document.getElementById(tableStates[type].bodyId);
            const visible = body ? Array.from(body.children).filter(tr => tr.style.display !== 'none' && !tr.classList.contains('grupo-sep')).length : tableStates[type].data.length;
            const total   = tableStates[type].data.length;
            const badge   = document.getElementById(`count-${type}`);
            if (badge) badge.textContent = visible < total ? `${visible}/${total}` : total;
        });
    }

    /* ═══════════════════════════════════════
       RESUMO DA REDE (TASK-008)
    ═══════════════════════════════════════ */
    /**
     * Extrai os pares "qtd-ativo" de uma linha da tabela Outros. Uma linha pode conter mais de
     * um par (ex.: "3-DT9 2-CV8"). Mesma lógica de services/orcamento_calc.py:98-122: "-" separa
     * quantidade do ativo, "*" representa linha viva (vira "-" antes de virar espaço), e "DT"/"CV"
     * sem quantidade explícita ganham "1-" automaticamente.
     */
    function extrairParesQtdAtivoOutros(ativoRaw) {
        if (!ativoRaw) return [];
        let txt = ativoRaw.trim();
        if (!txt) return [];
        if (/^(DT|CV)/i.test(txt) && !/^\d/.test(txt)) {
            txt = '1-' + txt;
        }
        const protegidos = [];                                                   // TASK-108: o que está entre [ ] (nome literal) não é quebrado nem tem o "-" trocado
        txt = txt.replace(/\[[^\[\]]*\]/g, m => { protegidos.push(m); return '\x00' + (protegidos.length - 1) + '\x00'; });
        txt = txt.replace(/-/g, ' ').replace(/\*/g, '-');
        const tokens = txt.split(/\s+/).filter(Boolean).map(t => t.replace(/\x00(\d+)\x00/g, (_, i) => protegidos[+i]));
        const pares = [];
        let i = 0;
        while (i + 1 < tokens.length) {
            let qtd = parseFloat(tokens[i].replace(',', '.'));
            if (isNaN(qtd)) qtd = 1;
            pares.push({ ativo: tokens[i + 1].toUpperCase(), qtd });
            i += 2;
        }
        return pares;
    }

    /**
     * Calcula as 4 contagens da "Resumo da rede": postes, rede de média, rede de baixa e equipamentos — só linhas instalando (I e *I).
     * A regra mora em resumo_rede.js (TASK-192), que também reconhece o poste pela convenção de postes do projeto (Pernambuco: `12600D`).
     */
    function calcularResumoRede() {
        return ResumoRede.calcular({
            cabos: tableStates.cabos.data, outros: tableStates.outros.data, ctxOutros: _ctxOutros(),
            convencaoPostes: window.__vinculacaoConfig ? window.__vinculacaoConfig.postes : undefined,
            AnotacaoOutros, AnotacaoCabos, extrairPares: extrairParesQtdAtivoOutros,
            extrairPrefixoEComprimento: extrairPrefixoEComprimentoCabo, extrairPrefixoSemComprimento: extrairPrefixoCaboSemComprimento
        });
    }

    function formatarNumeroResumoRede(n) {
        const arredondado = Math.round(n * 100) / 100;
        return arredondado % 1 === 0 ? String(arredondado) : arredondado.toFixed(2).replace('.', ',');
    }

    function agendarResumoRede() {
        if (_tmResumoRede) return;
        _tmResumoRede = setTimeout(() => { _tmResumoRede = 0; atualizarResumoRedeUI(); }, 0);
    }
    function atualizarResumoRedeUI() {
        const resumo = calcularResumoRede();
        const elPostes = document.getElementById('resumo-rede-postes');
        const elMedia = document.getElementById('resumo-rede-media');
        const elBaixa = document.getElementById('resumo-rede-baixa');
        const elEquip = document.getElementById('resumo-rede-equipamentos');
        if (elPostes) elPostes.textContent = formatarNumeroResumoRede(resumo.postes);
        if (elMedia) elMedia.textContent = `${formatarNumeroResumoRede(resumo.redeMedia)} m`;
        if (elBaixa) elBaixa.textContent = `${formatarNumeroResumoRede(resumo.redeBaixa)} m`;
        if (elEquip) elEquip.textContent = formatarNumeroResumoRede(resumo.equipamentos);
    }

    /* ═══════════════════════════════════════
       VINCULAÇÃO CABO<->ESTRUTURA/POSTE (TASK-032)
    ═══════════════════════════════════════ */
    function _escapeHtmlVinculacao(u) {
        return String(u || '').replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[m]));
    }

    let pedidoConfigVinculo = null;
    window.__vinculacaoConfig = null;
    window.carregarRegrasVinculacaoResumo = function (projetoCodigo) {
        if (!projetoCodigo) return Promise.resolve(null);
        if (pedidoConfigVinculo && pedidoConfigVinculo.codigo === projetoCodigo) return pedidoConfigVinculo.resposta;
        const pedido = { codigo: projetoCodigo };
        pedidoConfigVinculo = pedido;
        pedido.resposta = buscarConfigVinculo(pedido);
        return pedido.resposta;
    };
    async function buscarConfigVinculo(pedido) {
        const projetoCodigo = pedido.codigo;
        try {                                                                  // TASK-084: Tabela de Cabos + Tabela de Estruturas
            const rc = await fetch(`/api/vinculacao-config?projeto_codigo=${encodeURIComponent(projetoCodigo)}`);
            const config = rc.ok ? (await rc.json()).config : null;
            if (pedido !== pedidoConfigVinculo) return;
            window.__vinculacaoConfig = config;
            AnotacaoCabos.definirDicionario(window.__vinculacaoConfig && window.__vinculacaoConfig.dicionario_cabos);          // TASK-173
            if (window.PerfilExportacao) PerfilExportacao.definir(window.__vinculacaoConfig && window.__vinculacaoConfig.exportacao);          // TASK-176
            return config;
        } catch (e) {
            if (pedido !== pedidoConfigVinculo) return;
            window.__vinculacaoConfig = null;
            AnotacaoCabos.definirDicionario(null);
            console.error('Erro ao carregar a configuração de vinculação:', e);
        } finally {
            if (pedido === pedidoConfigVinculo) pedidoConfigVinculo = null;
        }
    }

    const idPontoMapa = () => Array.from(crypto.getRandomValues(new Uint8Array(16)), n => n.toString(16).padStart(2, '0')).join('');
    function idPontoValido(id) { return typeof id === 'string' && /^[a-f0-9]{32}$/.test(id); }

    function postesDaLinhaVinculo(ativo) {
        const ctx = _ctxOutros(), conv = ctx.convencao_postes;
        const r = AnotacaoOutros.interpretarLinha(ativo || '', 'I', _classicoTotalizadora, ctx);
        return r.itens.map(i => AnotacaoOutros.posteParaCompacto(AtivoParametros.nomeBase(i.ativo), conv)).filter(Boolean);
    }

    function idPosteNoMapa(destino, outros) {
        if (!idPontoValido(destino._mapaId) || outros.some(r => r && r !== destino && r._mapaId === destino._mapaId))
            destino._mapaId = idPontoMapa();
        return destino._mapaId;
    }

    function ehPosteNoMapa(row) {
        return row && !row.comentario && (row.entidade === 'POSTE' || postesDaLinhaVinculo(row.ativo).length);
    }

    /* Todo ponto do mapa é um grupo em Demais Ativos com título, ativos e comentário opcional. */
    const GL_ = () => window.GruposLinhas;
    const nomeDePontoExiste = (outros, nome, exceto) => outros.some(r => r && r.grupoDecl && r !== exceto && GL_().nomeDoPonto(r).toUpperCase() === nome.toUpperCase());
    /** Nome escolhido pelo usuário (ou o automático «P7»); `erro` se já existe outro ponto com esse nome. */
    /** Número mínimo do nome automático: pontos ainda sem grupo aparecem no mapa como «P<linha>», e o nome novo não pode repetir um deles. */
    function pisoNomePonto(outros) {
        let m = 0;
        outros.forEach((r, i) => { if (r && !r.grupoDe && !r.grupoDecl && !r.comentario && (r.entidade === 'POSTE' || r.entidade === 'ESTRUTURA' || r._poste)) m = i + 1; });
        return m;
    }
    function nomeParaPonto(outros, texto, exceto) {
        const GL = GL_(), digitado = GL.nomeLimpo(texto);
        if (digitado && nomeDePontoExiste(outros, digitado, exceto)) return { erro: 'Já existe um ponto com o nome «' + digitado + '».' };
        return { nome: digitado || GL.proximoNome(outros, 'P', pisoNomePonto(outros)) };
    }
    function inserirPosteFicticioNoMapa(nome, estruturas) {
        const outros = tableStates.outros.data, GL = GL_();
        let ponto = null;
        if (GL) { const r = nomeParaPonto(outros, nome); ponto = GL.abrirPonto(r.erro ? GL.proximoNome(outros, 'P', pisoNomePonto(outros)) : r.nome, 'M'); }
        pushHistory();
        const linha = {
            ativo: nome ? 'Poste fictício · ' + nome : 'Poste fictício', entidade: 'POSTE', operacao: 'M', entidadeManual: true,
            congelada: true, bloqueada: true, _mapaPosteFicticio: true, _mapaId: idPontoMapa(),
            _mapaEstruturasFicticias: estruturas.slice()
        };
        if (ponto) { linha.grupoDe = ponto.gid; outros.push(ponto.titulo); }
        return outros.push(linha) - 1;
    }

    function removerPosteFicticioNoMapa(indice) {
        const row = tableStates.outros.data[indice];
        if (row?._mapaPosteFicticio !== true) return false;
        pushHistory();
        tableStates.outros.data.splice(indice, 1);
        _reajustarVinculosAposRemoverOutros(indice);
        tableStates.cabos.data.forEach(c => {
            if (c?._mapaAncoras) c._mapaAncoras = c._mapaAncoras.filter(([id]) => id !== row._mapaId);
        });
        buildAtivoSets(); buildDataLists();
        refreshAllFilters('outros');
        return true;
    }

    function validarAnotacaoOutrosNoMapa(tipo, nome) {
        const r = AnotacaoOutros.interpretarLinha(nome, 'I', _classicoTotalizadora, _ctxOutros());
        if (r.erros.length || !r.itens.length) return r.erros[0]?.msg || 'Anotação inválida.';
        const postes = postesDaLinhaVinculo(nome);
        if (tipo === 'poste' && postes.length !== 1) return 'Informe um poste reconhecido pela convenção do projeto. Exemplo: 12600D, N3.';
        if (tipo === 'estrutura' && postes.length) return 'Esta linha de estrutura deve conter apenas estruturas.';
        return '';
    }

    function atualizarLinhaDoMapa(tabela, indice, recalcular) {
        if (recalcular) recalcRowAndAbove(indice);
        renderTable(tabela);
        buildAtivoSets();
        buildDataLists();
        refreshAllFilters(tabela);
        atualizarResumoRedeUI();
    }

    function validarEdicaoCaboNoMapa(indice, antes, nova, cfg) {
        const VR = VinculacaoRegras;
        const ids = antes.vinculoEstruturas || [], ficticias = VR.estruturasFicticiasDe(antes);
        if (!ids.length && !ficticias.length) return '';
        if (nova.operacao !== antes.operacao && !VR.participaVinculo(nova)) return 'Remova os vínculos internos antes de trocar para retirada.';
        const info = row => VR.infoCabo(_prefixoCaboVinculo(row.ativo) || '', row.operacao, cfg, VR.extraCabo(row, cfg.dicionario_cabos));
        const anterior = info(antes), posterior = info(nova);
        if ((anterior.vincula && !posterior.vincula) ||
            (posterior.pontasMax < anterior.pontasMax && ids.length + ficticias.length > posterior.pontasMax))
            return 'A alteração reduz as pontas aceitas ou desativa o vínculo. Remova os vínculos internos que impedem a edição.';
        const cabos = tableStates.cabos.data, outros = tableStates.outros.data;
        const analisar = lista => VR.analisar(lista, outros, cfg, _prefixoCaboVinculo, _codigoEstruturaVinculo, _codigosEstruturaVinculo);
        const anteriorAnalise = analisar(cabos), copia = cabos.slice();
        copia[indice] = nova;
        const posteriorAnalise = analisar(copia), estruturas = new Set(ids);
        const chave = l => l.cabo + ':' + l.estrutura;
        const anteriores = new Set(anteriorAnalise.links.filter(l => !l.ok).map(chave));
        const conflito = posteriorAnalise.links.find(l => !l.ok && estruturas.has(l.estrutura) && !anteriores.has(chave(l)));
        const errosAnteriores = anteriorAnalise.cabos[indice]?.erros || [];
        const erro = conflito?.motivo || posteriorAnalise.cabos[indice]?.erros?.find(m => !errosAnteriores.includes(m));
        return erro ? 'A edição tornaria um vínculo incompatível. ' + erro : '';
    }

    function validarEdicaoEstruturaNoMapa(indice, antes, nova) {
        const VR = VinculacaoRegras;
        const codigosAntes = _codigosEstruturaVinculo(antes.ativo), codigosDepois = _codigosEstruturaVinculo(nova.ativo);
        const slots = new Set(Object.entries(antes.cabosFicticiosSlots || {}).filter(([, v]) => v.length).map(([k]) => Number(k)));
        if (antes.cabosFicticios?.length) slots.add(0);
        for (const c of tableStates.cabos.data) {
            for (const id of c?.vinculoEstruturas || []) if (VR.linhaDe(id) === indice) slots.add(VR.slotDe(id));
        }
        if (!slots.size) return '';
        if (VR.participaVinculo(antes) && !VR.participaVinculo(nova)) return 'Remova os vínculos e os cabos fictícios desta linha antes de trocar para retirada.';
        for (const k of slots) {
            if (codigosAntes[k] !== codigosDepois[k]) return 'A edição mudaria a estrutura vinculada na posição ' + (k + 1) + '. Remova esse vínculo antes.';
        }
        return '';
    }

    function editarAtivoNoMapa(tipo, indice, referencia, dados, configVinculo) {
        const tabela = tipo === 'cabo' ? 'cabos' : 'outros';
        const row = tableStates[tabela].data[indice];
        if (!row || row !== referencia) return 'A linha mudou. Reabra a edição no cartão.';
        if (tipo === 'poste-ficticio') {
            if (!row._mapaPosteFicticio) return 'O poste fictício já não está disponível.';
            const estruturas = dados.estruturas;
            const emUso = tableStates.cabos.data.some(c =>
                (c?._mapaAncoras || []).some(([id, codigo]) => id === row._mapaId && !estruturas.includes(codigo)));
            if (emUso) return 'Solte os cabos ancorados antes de remover ou renomear a estrutura fictícia.';
            const ativo = dados.ativo ? 'Poste fictício · ' + dados.ativo : 'Poste fictício';
            if (ativo === row.ativo && JSON.stringify(estruturas) === JSON.stringify(row._mapaEstruturasFicticias)) return '';
            pushHistory();
            row.ativo = ativo;
            row._mapaEstruturasFicticias = estruturas.slice();
            atualizarLinhaDoMapa(tabela, indice, false);
            return '';
        }
        if (row.bloqueada || row.comentario) return 'A linha está bloqueada para edição.';
        const entrada = String(dados.ativo || '').trim();
        const mudouAtivo = entrada.toUpperCase() !== row.ativo.trim().toUpperCase();
        const ativo = mudouAtivo ? entrada.toUpperCase() : row.ativo;
        if (!entrada || entrada.length > 500 || /^[#@]/.test(entrada)) return 'Informe uma anotação válida.';
        if (dados.operacao !== row.operacao && !/^(?:(?:\*)?[IMR]|F)$/.test(dados.operacao)) return 'Escolha a operação do ativo.';
        const nova = { ...row, ativo };
        aplicarOperacaoNaLinha(nova, dados.operacao);
        const ponto = tipo === 'poste' && dados.ponto ? prepararPontoNoMapa(dados.ponto.linhas, dados.ponto) : null;
        if (ponto?.erro) return ponto.erro;
        let erro;
        if (tipo === 'cabo') {
            if (dados.fases !== String(row.fases ?? '') && (!/^\d?$/.test(dados.fases) || dados.fases === '0')) return 'Informe um número de fases de 1 a 9 ou deixe em branco.';
            if (dados.nivel !== String(row.nivel ?? '') && dados.nivel && !['MT', 'BT'].includes(dados.nivel)) return 'Escolha MT, BT ou deixe o nível automático.';
            if (mudouAtivo) nova.textoOriginal = ativo;
            for (const campo of ['fases', 'tensao', 'nivel']) {
                if (dados[campo] === String(row[campo] ?? '')) continue;
                if (dados[campo]) nova[campo] = campo === 'fases' ? Number(dados[campo]) : dados[campo];
                else delete nova[campo];
            }
            if (['ativo', 'operacao', 'fases', 'tensao', 'nivel'].every(k => String(nova[k] ?? '') === String(row[k] ?? ''))) return '';
            erro = validarEdicaoCaboNoMapa(indice, row, nova, VinculacaoRegras.normalizarConfig(configVinculo));
        } else {
            if (!['poste', 'estrutura'].includes(tipo)) return 'Tipo de linha inválido.';
            if (mudouAtivo) {
                erro = validarAnotacaoOutrosNoMapa(tipo, ativo);
                if (erro) return erro;
                reclassificarEntidade(nova);
            }
            if (!mudouAtivo && nova.operacao === row.operacao && !ponto?.aplicar) return '';
            erro = validarEdicaoEstruturaNoMapa(indice, row, nova);
        }
        if (erro) return erro;
        pushHistory();
        Object.assign(row, nova);
        for (const k of ['congelada', '_congeladaPorF']) if (!(k in nova)) delete row[k];               // sair do F descongela
        const novaLista = ponto?.aplicar?.();
        if (tipo === 'cabo') for (const campo of ['fases', 'tensao', 'nivel']) if (!(campo in nova)) delete row[campo];
        if (novaLista) aplicarLista('outros', novaLista, null);
        else atualizarLinhaDoMapa(tabela, indice, tipo === 'cabo' && mudouAtivo);
        return '';
    }

    function inserirAtivoNoMapa(tipo, dados) {
        const nome = String(dados.ativo || '').trim().toUpperCase();
        const operacao = dados.operacao;
        const tabela = tipo === 'cabo' ? 'cabos' : 'outros';
        if (!nome || nome.length > 440 || /^[#@]/.test(nome)) return { erro: 'Informe um ativo válido.' };
        if (!/^(?:(?:\*)?[IMR]|F)$/.test(operacao)) return { erro: 'Escolha a operação do ativo.' };
        const row = { entidade: 'CABO', operacao, ativo: nome, entidadeManual: true };
        if (operacao === 'F') { row.congelada = true; row._congeladaPorF = true; }                  // TASK-235: cabo fictício (fora do orçamento, vale no vínculo)
        if (tipo === 'cabo') {
            const entrada = String(dados.metros || '').trim().replace(',', '.');
            const metros = Number(entrada), fase = String(dados.fase || '').trim().toUpperCase();
            if (!/^\d+(?:\.\d+)?$/.test(entrada) || !Number.isFinite(metros) || metros <= 0)
                return { erro: 'Informe um comprimento maior que zero, em metros.' };
            const rotulo = `${nome} ${entrada} m`;
            if (AnotacaoCabos.ehNotacao(rotulo)) {
                const r = AnotacaoCabos.interpretar(rotulo);
                if (r.erros.length || !r.itens.length) return { erro: r.erros[0] || 'Rótulo de cabo inválido.' };
                row.ativo = rotulo;
            } else {
                if (!GradeTabela.FASES.includes(fase)) return { erro: 'Escolha as fases para esta anotação de cabo.' };
                row.ativo = `${nome} ${fase} ${entrada} m`;
            }
            if (['MT', 'BT'].includes(dados.nivel)) row.nivel = dados.nivel;
        } else {
            if (!['poste', 'estrutura'].includes(tipo)) return { erro: 'Escolha o tipo de ativo.' };
            const erro = validarAnotacaoOutrosNoMapa(tipo, nome);
            if (erro) return { erro };
            row.entidade = tipo === 'poste' ? 'POSTE' : 'ESTRUTURA';
        }
        const outros = tableStates.outros.data;
        const poste = tipo === 'estrutura' && Number.isInteger(dados.poste) ? outros[dados.poste] : null;
        if (tipo === 'estrutura' && Number.isInteger(dados.poste) && !ehPosteNoMapa(poste))
            return { erro: 'O poste escolhido já não está disponível.' };
        if (tipo === 'cabo' && Number.isInteger(dados.poste)) {
            const ponto = outros[dados.poste];
            if (!ponto?.ativo || ponto.comentario) return { erro: 'O ponto escolhido já não está disponível.' };
            row._mapaTopologia = true;
            row.vinculoPostes = [dados.poste];
        }
        // Estrutura associada a um poste entra no grupo dele. Poste ou estrutura avulsa ganha grupo próprio.
        const GL = GL_(), novoPonto = !!GL && tabela === 'outros' && !poste;
        let ponto = null;
        if (novoPonto) {
            const r = nomeParaPonto(outros, dados.nome);
            if (r.erro) return { erro: r.erro };
            ponto = GL.abrirPonto(r.nome, operacao, dados.comentario);
        }
        pushHistory();
        if (poste) row._mapaPosteRef = idPosteNoMapa(poste, outros);
        const lista = tableStates[tabela].data;
        let indice;
        const grupoDoPoste = GL && poste ? GL.pontoDaLinha(outros, dados.poste) : null;
        if (grupoDoPoste) {                                          // entra no fim do grupo do poste; os índices seguintes mudam (vínculos acompanham)
            const nova = outros.slice();
            row.grupoDe = grupoDoPoste.gid;
            nova.splice(grupoDoPoste.fim + 1, 0, row);
            _remapearVinculosOutros(GL.mapa(outros, nova));
            tableStates.outros.data = nova;
            indice = grupoDoPoste.fim + 1;
        } else {
            if (ponto) {
                row.grupoDe = ponto.gid;
                lista.push(ponto.titulo);
                if (ponto.comentario) lista.push(ponto.comentario);
            }
            indice = lista.push(row) - 1;
        }
        if (tabela === 'cabos') recalcRowAndAbove(indice);
        renderTable(tabela);
        buildAtivoSets();
        buildDataLists();
        refreshAllFilters(tabela);
        return { indice };
    }

    function associarEstruturaNoMapa(indice, poste) {
        const outros = tableStates.outros.data, linha = outros[indice], destino = outros[poste];
        if (!linha || linha.comentario || ehPosteNoMapa(linha)) return false;
        if (poste !== null && (!ehPosteNoMapa(destino) || destino._mapaPosteFicticio)) return false;
        pushHistory();
        if (poste === null) delete linha._mapaPosteRef;
        else linha._mapaPosteRef = idPosteNoMapa(destino, outros);
        return true;
    }

    /** TASK-225: o grupo (ponto) a que pertencem estas linhas de Demais Ativos — título «@», comentário «#» e membros — com o estado de congelar/bloquear. */
    function pontoNoMapa(linhas) {
        const outros = tableStates.outros.data, GL = GL_();
        if (!GL) return null;
        for (const l of linhas || []) {
            const p = GL.pontoDaLinha(outros, l); if (!p) continue;
            const membros = p.membros.map(i => outros[i]), titulo = outros[p.h];
            return { ...p, tituloIdx: p.h, titulo: titulo.ativo, congelada: membros.length > 0 && membros.every(r => r.congelada), bloqueada: [titulo, ...membros].every(r => r.bloqueada),
                ficticio: membros.some(r => r._mapaPosteFicticio) };
        }
        return null;
    }
    /** Dá nome (ou o número automático) e comentário ao ponto. Sem grupo ainda: cria o grupo, o que leva as linhas do ponto para baixo do título (os vínculos por índice acompanham). */
    function prepararPontoNoMapa(linhas, dados) {
        const outros = tableStates.outros.data, GL = GL_();
        if (!GL) return { erro: 'Grupos indisponíveis.' };
        const atual = pontoNoMapa(linhas);
        const digitado = GL.nomeLimpo(dados.nome);
        if (!atual) {
            if (dados.manterSemGrupo && !digitado && !GL.limparComentario(dados.comentario)) return {};
            const objetos = (linhas || []).map(l => outros[l]).filter(r => r && !r.comentario);
            if (!objetos.length) return { erro: 'O ponto não tem linha na tabela.' };
            const r = nomeParaPonto(outros, dados.nome);
            if (r.erro) return r;
            return { aplicar() {
                return GL.agruparPontos(outros, [{ linhas: objetos, nome: r.nome, comentario: dados.comentario }]).lista;
            } };
        }
        const titulo = outros[atual.tituloIdx];
        let nome = atual.nome;
        if (digitado) {
            if (nomeDePontoExiste(outros, digitado, titulo)) return { erro: 'Já existe um ponto com o nome «' + digitado + '».' };
            nome = digitado;
        } else if (dados.nome !== undefined && !/^P\d+$/i.test(nome)) {
            nome = GL.proximoNome(outros, 'P', pisoNomePonto(outros));
        }
        // TASK-235: o comentário vai no título («@nome - comentário»); linhas «#» antigas saem (o texto delas já está em `atual.comentario`)
        const comentario = dados.comentario === undefined ? atual.comentario : dados.comentario;
        const novoTitulo = GL.tituloDoPonto(nome, comentario);
        if (novoTitulo === titulo.ativo && !atual.formatoAntigo) return {};
        return { aplicar() {
            titulo.ativo = novoTitulo;
            if (atual.formatoAntigo) {
                const nova = outros.slice();
                for (const k of atual.comentarios.slice().reverse()) nova.splice(k, 1);
                return nova;
            }
        } };
    }

    function definirPontoNoMapa(linhas, dados) {
        const ponto = prepararPontoNoMapa(linhas, dados);
        if (ponto.erro) return ponto.erro;
        if (ponto.aplicar) {
            pushHistory();
            const novaLista = ponto.aplicar();
            if (novaLista) aplicarLista('outros', novaLista, null);
            else atualizarLinhaDoMapa('outros', linhas[0], false);
        }
        return '';
    }
    /**
     * TASK-235: grava o ponto inteiro pelo mapa — nome, comentário (no título «@nome - comentário») e as linhas, cada uma com a operação (F = fictícia).
     * dados = {linhas: índices atuais do ponto | null (ponto novo), nome, comentario, itens: [{indice?, ativo, op}], forcar}.
     * Trocar a estrutura MANTÉM os vínculos dos cabos quando a nova aceita os mesmos cabos (regras do projeto); senão devolve {conflitos} e só aplica com `forcar`
     * (aí os vínculos que não servem caem). Um passo de Ctrl+Z. Devolve {indice} (1ª linha do ponto), {conflitos: [texto]} ou {erro}.
     */
    function salvarPontoNoMapa(dados, cfgVinculo) {
        const GL = GL_(), VR = VinculacaoRegras, outros = tableStates.outros.data, cabos = tableStates.cabos.data;
        if (!GL) return { erro: 'Grupos indisponíveis.' };
        const itens = (dados.itens || []).map(x => ({ indice: Number.isInteger(x.indice) ? x.indice : null, ativo: String(x.ativo || '').trim().toUpperCase(), op: String(x.op || '').trim().toUpperCase() }))
            .filter(x => x.ativo || x.indice !== null);
        if (!itens.some(x => x.ativo)) return { erro: 'Preencha ao menos uma linha do ponto.' };
        for (const x of itens) {
            if (!x.ativo) continue;
            if (x.ativo.length > 440 || /^[#@]/.test(x.ativo)) return { erro: 'Linha inválida: ' + x.ativo };
            if (!OPERACOES_LINHA.includes(x.op)) return { erro: 'Escolha a operação da linha ' + x.ativo + '.' };
            const r = AnotacaoOutros.interpretarLinha(x.ativo, 'I', _classicoTotalizadora, _ctxOutros());
            if (r.erros.length || !r.itens.length) return { erro: x.ativo + ': ' + (r.erros[0]?.msg || 'anotação inválida.') };
        }
        const montar = (base, x) => {
            const row = Object.assign({}, base || { entidade: '0', entidadeManual: true }, { ativo: x.ativo });
            aplicarOperacaoNaLinha(row, x.op);
            if (!base || base.ativo !== x.ativo) row.entidade = postesDaLinhaVinculo(x.ativo).length ? 'POSTE' : 'ESTRUTURA';
            return row;
        };
        // ── ponto novo ──
        if (!dados.linhas || !dados.linhas.length) {
            const r = nomeParaPonto(outros, dados.nome);
            if (r.erro) return { erro: r.erro };
            const linhas = itens.filter(x => x.ativo).map(x => montar(null, x));
            const ab = GL.abrirPonto(r.nome, linhas[0].operacao, dados.comentario);
            linhas.forEach(l => { l.grupoDe = ab.gid; });
            pushHistory();
            outros.push(ab.titulo);
            const indice = outros.length;
            outros.push(...linhas);
            atualizarLinhaDoMapa('outros', indice, false);
            return { indice, nome: r.nome };
        }
        // ── ponto existente ──
        const p = pontoNoMapa(dados.linhas);
        const atuais = p ? p.membros : dados.linhas.filter(l => outros[l] && !outros[l].comentario && !outros[l].grupoDecl);
        if (atuais.some(l => outros[l].bloqueada && itens.some(x => x.indice === l && (x.ativo !== outros[l].ativo || x.op !== String(outros[l].operacao || '').toUpperCase()))))
            return { erro: 'Há linha bloqueada no ponto: desbloqueie antes de editar.' };
        const novos = new Map(), removidas = atuais.filter(l => !itens.some(x => x.indice === l && x.ativo));
        for (const x of itens) if (x.indice !== null && x.ativo && atuais.includes(x.indice)) {
            const antes = outros[x.indice];
            if (antes.ativo !== x.ativo || String(antes.operacao || '').toUpperCase() !== x.op) novos.set(x.indice, montar(antes, x));
        }
        // vínculos dos cabos que deixariam de valer
        const cfg = VR.normalizarConfig(cfgVinculo || window.__vinculacaoConfig || {});
        const copia = outros.map((r, i) => novos.get(i) || r);
        const quebrados = [], chave = (ic, id) => ic + ':' + id;
        const antesA = VR.analisar(cabos, outros, cfg, _prefixoCaboVinculo, _codigoEstruturaVinculo, _codigosEstruturaVinculo);
        const depoisA = VR.analisar(cabos, copia, cfg, _prefixoCaboVinculo, _codigoEstruturaVinculo, _codigosEstruturaVinculo);
        const okAntes = new Set(antesA.links.filter(l => l.ok).map(l => chave(l.cabo, l.estrutura)));
        const okDepois = new Map(depoisA.links.map(l => [chave(l.cabo, l.estrutura), l]));
        cabos.forEach((c, ic) => (c?.vinculoEstruturas || []).forEach(id => {
            const l = VR.linhaDe(id), k = VR.slotDe(id);
            const nome = 'Cabo ' + (ic + 1) + ' (' + c.ativo + ')';
            if (removidas.includes(l)) { quebrados.push({ ic, id, texto: nome + ': a linha ' + outros[l].ativo + ' sai do ponto' }); return; }
            if (!novos.has(l)) return;
            const nova = novos.get(l), codigos = _codigosEstruturaVinculo(nova.ativo);
            if (!VR.participaVinculo(nova)) { quebrados.push({ ic, id, texto: nome + ': ' + nova.ativo + ' em retirada não recebe cabo' }); return; }
            if (k >= codigos.length) { quebrados.push({ ic, id, texto: nome + ': a estrutura da posição ' + (k + 1) + ' sai da linha' }); return; }
            const depois = okDepois.get(chave(ic, id));
            if (okAntes.has(chave(ic, id)) && depois && !depois.ok) quebrados.push({ ic, id, texto: nome + ': ' + (depois.motivo || codigos[k] + ' não aceita este cabo') });
        }));
        if (quebrados.length && !dados.forcar) return { conflitos: quebrados.map(q => q.texto) };
        const tituloAtual = p ? outros[p.tituloIdx] : null;
        const digitado = GL.nomeLimpo(dados.nome);
        if (digitado && nomeDePontoExiste(outros, digitado, tituloAtual)) return { erro: 'Já existe um ponto com o nome «' + digitado + '».' };
        pushHistory();
        quebrados.forEach(({ ic, id }) => { const c = cabos[ic]; c.vinculoEstruturas = (c.vinculoEstruturas || []).filter(x => x !== id); });
        novos.forEach((nova, i) => { const row = outros[i]; Object.assign(row, nova); for (const k of ['congelada', '_congeladaPorF']) if (!(k in nova)) delete row[k]; });
        let lista = outros.slice();
        const remover = new Set(removidas.map(l => outros[l]));
        lista = lista.filter(r => !remover.has(r));
        const adicionar = itens.filter(x => x.indice === null && x.ativo).map(x => montar(null, x));
        if (p) {
            const titulo = tituloAtual;
            titulo.ativo = GL.tituloDoPonto(digitado || p.nome, dados.comentario === undefined ? p.comentario : dados.comentario);
            if (p.formatoAntigo) { const fora = new Set(p.comentarios.map(k => outros[k])); lista = lista.filter(r => !fora.has(r)); }
            adicionar.forEach(r => { r.grupoDe = p.gid; });
            const fim = lista.reduce((f, r, i) => (r.grupoDe === p.gid || r === titulo ? i : f), -1);
            lista.splice(fim + 1, 0, ...adicionar);
        } else {
            const membros = atuais.filter(l => !removidas.includes(l)).map(l => outros[l]);
            lista.push(...adicionar);
            lista = GL.agruparPontos(lista, [{ linhas: membros.concat(adicionar), nome: dados.nome, comentario: dados.comentario }], pisoNomePonto(outros)).lista;
        }
        aplicarLista('outros', lista, null);
        const primeira = (p ? lista.findIndex(r => r.grupoDe === p.gid && !r.comentario) : lista.findIndex(r => r === (adicionar[0] || outros[atuais.find(l => !removidas.includes(l))])));
        return { indice: primeira, quebrados: quebrados.length };
    }
    /** TASK-235: converte pontos do formato antigo — grupos com linha «#» (o texto vai para o título, juntos por «; ») e, se pedidos, pontos sem grupo. Um passo de Ctrl+Z. */
    function converterPontosNoMapa(semGrupo) {
        const GL = GL_(), outros = tableStates.outros.data;
        if (!GL) return 0;
        let res = GL.converterPontos(outros), lista = res.lista, n = res.n;
        const pontos = (semGrupo || []).map(p => ({ nome: p.nome, linhas: (p.linhas || []).map(l => outros[l]).filter(r => r && !r.comentario && !r.grupoDe && !r.grupoDecl) })).filter(p => p.linhas.length);
        if (!n && !pontos.length) return 0;
        pushHistory();
        if (pontos.length) { const g = GL.agruparPontos(lista, pontos, pisoNomePonto(outros)); lista = g.lista; n += g.n; }
        aplicarLista('outros', lista, null);
        return n;
    }
    /** Congela / bloqueia o ponto inteiro (todas as linhas do grupo), como no Resumo: repetir desfaz. */
    function marcarPontoNoMapa(linhas, campo) {
        const p = pontoNoMapa(linhas);
        if (!p || p.ficticio) return false;
        alternarMarcaLinhas('outros', [p.tituloIdx], campo);
        return true;
    }
    /** «Agrupar pontos»: cada ponto vira «@Pn» + as suas linhas (um passo de Ctrl+Z). */
    function agruparPontosNoMapa(pontos) {
        const outros = tableStates.outros.data, GL = GL_();
        if (!GL) return 0;
        const lista = pontos.map(p => ({ nome: p.nome, linhas: (p.linhas || []).map(l => outros[l]).filter(r => r && !r.comentario && !r.grupoDe && !r.grupoDecl) })).filter(p => p.linhas.length);
        if (!lista.length) return 0;
        pushHistory();
        const res = GL.agruparPontos(outros, lista, pisoNomePonto(outros));
        if (!res.n) return 0;
        aplicarLista('outros', res.lista, null);
        return res.n;
    }

    /** Campos que NÃO vão para uma cópia: geometria do CAD, vínculos com cabos, posição no mapa, identidade e grupo (a cópia ganha os seus). */
    const CAMPOS_FORA_DA_COPIA = ['bloqueada', '_x', '_y', '_geo', '_geoTipo', '_poste', '_posteAprox', 'vinculoEstruturas', 'vinculoOrigem', 'vinculoPostes', 'estruturasFicticias', 'cabosFicticios',
        'cabosFicticiosSlots', 'origemAvulso', '_mapaPoste', '_mapaPoste500', '_mapaCabo', '_mapaCabo500', '_mapaId', '_mapaPosteRef', '_mapaTopologia', '_mapaAncoras', 'grupoDe', 'grupoDecl', 'gid', 'grupoRecolhido'];
    /** Copia o ponto: um grupo novo (nome automático, mesmo comentário) com cópias das linhas do ativo, sem vínculos nem posição. Devolve {indice, nome} (indice = 1ª linha copiada) ou {erro}. */
    function copiarPontoNoMapa(linhas) {
        const outros = tableStates.outros.data, GL = GL_();
        if (!GL) return { erro: 'Grupos indisponíveis.' };
        const p = pontoNoMapa(linhas);
        const origens = p ? p.membros.map(i => outros[i]).filter(r => r && !r.comentario) : (linhas || []).map(l => outros[l]).filter(r => r && !r.comentario && !r.grupoDecl);
        if (!origens.length) return { erro: 'O ponto não tem linhas para copiar.' };
        const nome = nomeParaPonto(outros, '').nome, ab = GL.abrirPonto(nome, origens[0].operacao, p ? p.comentario : '');
        const novosIds = new Map();
        const copias = origens.map(r => {
            const c = deepClone(r), ficticio = r._mapaPosteFicticio === true;
            CAMPOS_FORA_DA_COPIA.forEach(k => delete c[k]);
            if (ficticio) { c.congelada = true; c.bloqueada = true; c._mapaPosteFicticio = true; c._mapaEstruturasFicticias = (r._mapaEstruturasFicticias || []).slice(); }
            if (idPontoValido(r._mapaId)) { c._mapaId = idPontoMapa(); novosIds.set(r._mapaId, c._mapaId); }
            c._copiaDe = r._mapaPosteRef;                                                       // temporário: remapeia a ligação estrutura→poste
            c.grupoDe = ab.gid;
            return c;
        });
        copias.forEach(c => { const antigo = c._copiaDe; delete c._copiaDe; if (idPontoValido(antigo) && novosIds.has(antigo)) c._mapaPosteRef = novosIds.get(antigo); });
        pushHistory();
        outros.push(ab.titulo);
        if (ab.comentario) outros.push(ab.comentario);
        const indice = outros.length;
        outros.push(...copias);
        atualizarLinhaDoMapa('outros', indice, false);
        return { indice, nome };
    }
    /** Copia o cabo (mesmo texto, nível e fases) para o fim de Cabos, sem pontas ligadas: o usuário liga no mapa. Devolve {indice} ou {erro}. */
    function copiarCaboNoMapa(ic) {
        const cabos = tableStates.cabos.data, origem = cabos[ic];
        if (!origem || origem.comentario) return { erro: 'O cabo já não está disponível.' };
        pushHistory();
        const c = deepClone(origem);
        CAMPOS_FORA_DA_COPIA.forEach(k => delete c[k]);
        delete c.qtdAtivos;
        const indice = cabos.push(c) - 1;
        atualizarLinhaDoMapa('cabos', indice, true);
        return { indice };
    }

    /**
     * TASK-233: exclui pelo mapa (tecla Del) — cabos pelo índice e pontos pelas linhas de Demais Ativos. Um passo de histórico.
     * Grupo «@» que fica sem membros sai inteiro (declaração e comentário); vínculos dos cabos acompanham os índices novos.
     */
    function removerDoMapa(sel) {
        const remC = new Set((sel && sel.cabos || []).filter(Number.isInteger)), remO = new Set((sel && sel.outros || []).filter(Number.isInteger));
        const outros = tableStates.outros.data, cabos = tableStates.cabos.data, GL = GL_();
        for (const i of [...remC]) if (!cabos[i] || cabos[i].bloqueada) remC.delete(i);
        for (const i of [...remO]) if (!outros[i] || outros[i].bloqueada) remO.delete(i);
        if (GL) {                                                                   // grupo sem nenhum membro restante: some a declaração e o comentário
            const gids = new Set([...remO].map(i => outros[i].grupoDe).filter(Boolean));
            for (const gid of gids) {
                const restam = outros.some((r, i) => r && r.grupoDe === gid && !r.comentario && !remO.has(i));
                if (!restam) outros.forEach((r, i) => { if (r && (r.grupoDe === gid || (r.grupoDecl && r.gid === gid))) remO.add(i); });
            }
        }
        if (!remC.size && !remO.size) return { erro: 'Nada a excluir (linhas bloqueadas ficam).' };
        pushHistory();
        if (remC.size) {
            const novoIndice = new Map(); let k = 0;
            cabos.forEach((r, i) => { if (!remC.has(i)) novoIndice.set(i, k++); });
            outros.forEach(r => {                                                    // TASK-171: avulsos da fictícia guardam o índice do cabo
                if (!r || !r.origemAvulso) return;
                const m = /^(\d+)(:.*)$/.exec(r.origemAvulso);
                if (!m) return;
                const n = novoIndice.get(Number(m[1]));
                if (n === undefined) r.origemAvulso = m[1] + m[2] + (m[2].endsWith('!velho') ? '' : '!velho'); else r.origemAvulso = n + m[2];
            });
            tableStates.cabos.data = cabos.filter((r, i) => !remC.has(i));
        }
        if (remO.size) {
            let nova = outros.filter((r, i) => !remO.has(i));
            if (GL) nova = GL.normalizar(nova);
            const pos = new Map(); nova.forEach((r, k) => pos.set(r, k));
            _remapearVinculosOutros(outros.map(r => pos.get(r)));
            tableStates.outros.data = nova;
        }
        ['cabos', 'outros'].forEach(t => { renderTable(t); refreshAllFilters(t); });
        buildAtivoSets(); buildDataLists();
        return { cabos: remC.size, outros: remO.size };
    }
    // TASK-233: fundo do CAD para o mapa (rede e blocos de poste sem anotação) — vem do projeto do plugin ou da leitura do DXF
    let fundoCadAtual = null;
    function definirFundoCad(fundo) {
        fundoCadAtual = fundo && Array.isArray(fundo.redes) && Array.isArray(fundo.postes) ? fundo : null;
        try { if (fundoCadAtual) sessionStorage.setItem('orcacore_fundo_cad', JSON.stringify(fundoCadAtual)); else sessionStorage.removeItem('orcacore_fundo_cad'); } catch (e) { /* grande demais ou sem storage: vale só nesta tela */ }
    }
    function fundoCad() {
        if (fundoCadAtual) return fundoCadAtual;
        for (const [store, chave] of [[sessionStorage, 'orcacore_fundo_cad'], [localStorage, 'mapa_geometria_fundo']]) {
            try { const f = JSON.parse(store.getItem(chave) || 'null'); if (f && Array.isArray(f.redes) && Array.isArray(f.postes)) return f; } catch (e) { /* sem fundo */ }
        }
        return null;
    }
    window.OrcaFundoCad = { definir: definirFundoCad, atual: fundoCad };

    /** TASK-084: o que o modal de vinculação (vinculacao_ui.js) precisa saber/fazer na tela do Resumo. */
    window.VinculacaoHost = {
        copiarPontoNoMapa, copiarCaboNoMapa, removerDoMapa, fundoCad, salvarPontoNoMapa, converterPontosNoMapa,
        pontoNoMapa, definirPontoNoMapa, marcarPontoNoMapa, agruparPontosNoMapa,
        cabos: () => tableStates.cabos.data,
        outros: () => tableStates.outros.data,
        projeto: () => localStorage.getItem('projeto_selecionado_codigo') || 'DEFAULT',
        prefixoCabo: _prefixoCaboVinculo,
        codigoEstrutura: _codigoEstruturaVinculo,
        codigosEstrutura: _codigosEstruturaVinculo,            // TASK-094: todas as estruturas da linha (9400D, N1, N4 → N1, N4)
        avulsosDoCabo: ic => _avulsosDoCabo(ic),               // TASK-171: componentes avulsos que a estrutura fictícia do cabo usa (alças, conectores)
        adicionarAvulsos: (ic, ficticia) => _adicionarAvulsos(ic, ficticia),
        postesDaLinha: postesDaLinhaVinculo,
        inserirNoMapa: inserirAtivoNoMapa,
        editarNoMapa: editarAtivoNoMapa,
        inserirPosteFicticioNoMapa,
        removerPosteFicticioNoMapa,
        associarNoMapa: associarEstruturaNoMapa,
        linhaSoPoste(ativo) { return !AnotacaoOutros.codigoEstrutura(ativo || '', _classicoTotalizadora, _ctxOutros()); },
        consultarAssistente: MODO_JANELA_MAPA ? undefined : selecao => selecionarMapaAssistente(selecao),
        antes() { pushHistory(); },
        desfazer() { undo(); },
        depois() { try { renderTable('outros'); atualizarResumoRedeUI(); if (typeof updateHistoryUI === 'function') updateHistoryUI(); } catch (e) { /* tela ainda carregando */ } },
        instantaneo: instantaneoDoMapa,
        aplicarInstantaneo: aplicarInstantaneoDoMapa
    };

    /** TASK-241: cópia das linhas de Cabos e Demais Ativos (o que o Mapa lê e muda) para a outra janela do Mapa. */
    function instantaneoDoMapa() {
        return JSON.parse(JSON.stringify({ cabos: tableStates.cabos.data, outros: tableStates.outros.data }));
    }
    /** TASK-241: aplica as linhas vindas da outra janela do Mapa (um passo de Ctrl+Z). false = igual ao que já está aqui. */
    function aplicarInstantaneoDoMapa(dados) {
        if (!dados || !Array.isArray(dados.cabos) || !Array.isArray(dados.outros)) return false;
        if (JSON.stringify({ cabos: tableStates.cabos.data, outros: tableStates.outros.data }) === JSON.stringify({ cabos: dados.cabos, outros: dados.outros })) return false;
        pushHistory();
        tableStates.cabos.data = JSON.parse(JSON.stringify(dados.cabos));
        tableStates.outros.data = JSON.parse(JSON.stringify(dados.outros));
        try {
            buildAtivoSets();
            renderTable('cabos');
            renderTable('outros');
            buildDataLists();
            refreshAllFilters('cabos');
            refreshAllFilters('outros');
            updateCounters();
            atualizarResumoRedeUI();
            updateHistoryUI();
        } catch (e) { console.error('Mapa: falha ao aplicar a obra da outra janela.', e); }
        return true;
    }

    if (MODO_JANELA_MAPA) {
        let assinaturaMapa = '', envioMapa = 0;
        const aplicarDaPrincipal = dados => {
            window.VinculacaoHost.aplicarInstantaneo(dados);
            assinaturaMapa = JSON.stringify(instantaneoDoMapa());              // não devolve à principal o que veio dela
        };
        window.__agendarEnvioMapa = () => {
            if (envioMapa) return;
            envioMapa = requestAnimationFrame(() => {
                envioMapa = 0;
                const dados = instantaneoDoMapa(), assinatura = JSON.stringify(dados);
                if (assinatura === assinaturaMapa) return;
                assinaturaMapa = assinatura;
                try { window.top.__enviarDadosMapa?.(dados); } catch (e) { /* moldura de outra origem: sem sincronização */ }
            });
        };
        window.JanelaConteudo = {
            capturar: () => ({ dados: instantaneoDoMapa() }),
            async restaurar(estado) {
                if (estado?.fundo) definirFundoCad(estado.fundo);
                if (estado?.dados) aplicarDaPrincipal(estado.dados);
                await window.VinculacaoUI?.abrir(window.VinculacaoHost, { aba: 'mapa' });
            },
            dados(dados) { aplicarDaPrincipal(dados); window.VinculacaoUI?.atualizarObra(false, true); },
            suspender() {},
            recalcular() {}
        };
    }

    document.getElementById('select-projeto')?.addEventListener('change', () => window.VinculacaoUI?.atualizarObra());

    /* ═══════════════════════════════════════
       TRANSPORTE DE MATERIAL (TASK-245) — o cálculo e o modal estão em static/transporte.js.
       O item "Transporte de material…" só aparece quando o projeto ativo tem a parte `transporte` do pacote.
    ═══════════════════════════════════════ */
    const _transporteCache = new Map();          // projeto → Promise(dados | {})
    function _projetoAtivoNome() {
        const sel = document.getElementById('select-projeto');
        return (sel && sel.value) || localStorage.getItem('projeto_selecionado') || '';
    }
    function _dadosTransporte(projeto) {
        if (!projeto) return Promise.resolve({});
        if (!_transporteCache.has(projeto)) {
            _transporteCache.set(projeto, fetch('/api/orcamento/transporte?projeto=' + encodeURIComponent(projeto))
                .then(r => (r.ok ? r.json() : {})).catch(() => { _transporteCache.delete(projeto); return {}; }));
        }
        return _transporteCache.get(projeto);
    }
    async function atualizarItemTransporte() {
        const btn = document.getElementById('btn-transporte');
        if (!btn) return;
        const projeto = _projetoAtivoNome(), dados = await _dadosTransporte(projeto);
        if (projeto === _projetoAtivoNome()) btn.hidden = !(dados && dados.formato);
    }
    /** "Lançar em Outros": troca as linhas TRANSP/TRANSP_KM de um cálculo anterior pelas novas (um passo de Ctrl+Z). */
    function lancarTransporteEmOutros(linhas, dados) {
        const antiga = tableStates.outros.data;
        const velhas = antiga.filter(r => r && !r.comentario && Transporte.ehLinhaTransporte(r.ativo, dados));
        if (!linhas.length && !velhas.length) return 'Nada a lançar: a obra não tem material pesado.';
        pushHistory();
        const nova = antiga.filter(r => !velhas.includes(r));
        let fim = nova.length;
        while (fim > 0 && nova[fim - 1] && !String(nova[fim - 1].ativo || '').trim() && !nova[fim - 1].grupoDecl && !nova[fim - 1].grupoDe) fim--;   // antes das linhas vazias do fim
        nova.splice(fim, 0, ...linhas.map(texto => Object.assign(window.GradeTabela.novaLinha('outros', 'I'), { ativo: texto, entidade: 'OUTROS', origemTransporte: true })));
        const pos = new Map(); nova.forEach((r, k) => pos.set(r, k));
        _remapearVinculosOutros(antiga.map(r => pos.get(r)));
        tableStates.outros.data = nova;
        renderTable('outros');
        try { atualizarResumoRedeUI(); } catch (e) { /* tela carregando */ }
        if (window.GridRProjeto && GridRProjeto.marcarSujo) GridRProjeto.marcarSujo(false);
        const msg = linhas.length
            ? `${linhas.length} linha(s) de transporte lançada(s) em Outros` + (velhas.length ? ` (${velhas.length} antiga(s) substituída(s))` : '') + '. Ctrl+Z desfaz.'
            : `Sem serviço de transporte: ${velhas.length} linha(s) antiga(s) removida(s). Ctrl+Z desfaz.`;
        showToast(msg);
        return msg;
    }
    async function abrirTransporte() {
        const projeto = _projetoAtivoNome(), dados = await _dadosTransporte(projeto);
        if (!dados || !dados.formato) { showToast('O projeto ativo não tem a parte "transporte" no pacote.'); return; }
        Transporte.abrirModal({
            dados,
            obra: window.ObraParams ? ObraParams.obraDoProjeto() : {},
            async obterResultado() {
                // o mesmo pedido do Montar Orçamento, sem as linhas TRANSP/TRANSP_KM que já estão em Outros (a Totalizadora já deixa de fora congeladas e comentários)
                const p = await obterPayloadCalculo();
                const outros = p.outros.filter(o => !Transporte.ehLinhaTransporte(o.ativo, dados));
                const r = await fetch('/api/orcamento/calcular', { method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ cabos: p.cabos, outros, projeto, obra: window.ObraParams ? ObraParams.obraDoProjeto() : {} }) });
                const d = await r.json().catch(() => ({}));
                if (!r.ok) return { erro: 'O orçamento do Resumo não pôde ser calculado' + (d && d.detail ? ` (${typeof d.detail === 'string' ? d.detail : 'erro'})` : '') + '.' };
                return d;
            },
            lancar: linhas => lancarTransporteEmOutros(linhas, dados)
        });
    }
    document.getElementById('btn-transporte')?.addEventListener('click', abrirTransporte);
    document.getElementById('select-projeto')?.addEventListener('change', atualizarItemTransporte);
    document.getElementById('btn-menu-obra')?.addEventListener('pointerdown', atualizarItemTransporte);
    atualizarItemTransporte();
    setTimeout(atualizarItemTransporte, 1500);                // o projeto é escolhido por código ao carregar a lista (sem evento)

    window.abrirModalVinculacao = function () {
        if (window.VinculacaoUI) window.VinculacaoUI.abrir(window.VinculacaoHost);
    };

    /** Regras de conectores (cabo + estrutura vinculados → ativo): continuam no modal antigo, agora só com esta seção. */
    window.abrirModalConectores = function () {
        const modal = document.getElementById('modal-vinculacao');
        if (!modal) return;
        modal.classList.remove('hidden');
        modal.style.zIndex = 9500;
        const content = document.getElementById('vinc-conectores-content');
        if (content && content.classList.contains('hidden') && window.toggleVincConectores) window.toggleVincConectores();
    };

    /**
     * TASK-033: acha estruturas/postes sem o vínculo certo pro seu tipo (regras de vinculação,
     * TASK-032). Achados no mesmo formato de /api/validacao/planilhas, para aparecerem juntos no
     * modal de Validação — sem round-trip ao backend, já que o vínculo só existe no frontend.
     */
    function avaliarVinculacaoLocal() {
        // TASK-084: usa a Tabela de Cabos + Tabela de Estruturas do projeto (a mesma lógica do modal e do modo autônomo)
        const cfg = window.__vinculacaoConfig;
        if (!cfg || typeof VinculacaoRegras === 'undefined') return [];
        return VinculacaoRegras.achados(tableStates.cabos.data, tableStates.outros.data, cfg, _prefixoCaboVinculo, _codigoEstruturaVinculo, _codigosEstruturaVinculo);
    }

    /* ═══════════════════════════════════════
       TOAST
    ═══════════════════════════════════════ */
    let toastTimer = null;
    function showToast(msg) {
        const t = document.getElementById('resumo-toast');
        t.textContent = msg;
        t.classList.add('show');
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => t.classList.remove('show'), 1600);
    }
    window.showToast = showToast;  // usado pelo painel de regras (painel_regras.js)

    /* ═══════════════════════════════════════
       UTILS
    ═══════════════════════════════════════ */
    function deepClone(obj) {
        const c = { entidade: obj.entidade || '0', operacao: obj.operacao || 'M', ativo: obj.ativo || '' };
        if (obj.qtdAtivos !== undefined) c.qtdAtivos = obj.qtdAtivos;
        // TASK-092: texto original do cabo ("3xA35P(7,94)-13,8kV 45 m") e Fases/Tensão/Nível editados — as regras de vínculo usam
        const txtOrig = obj.textoOriginal || (obj._raw && obj.entidade === 'CABO' ? obj._raw.texto : '');
        if (txtOrig) c.textoOriginal = txtOrig;
        ['fases', 'tensao', 'nivel'].forEach(k => { if (obj[k] !== undefined && obj[k] !== '') c[k] = obj[k]; });
        if (obj.entidadeManual) c.entidadeManual = true;   // TASK-065: classificação definida à mão não é recalculada
        if (obj.congelada) c.congelada = true; if (obj.bloqueada) c.bloqueada = true; if (obj.comentario) c.comentario = true;
        if (obj._congeladaPorF) c._congeladaPorF = true;                                              // TASK-235: a operação F congelou a linha
        if (obj.grupoDecl) { c.grupoDecl = true; c.gid = obj.gid; if (obj.grupoRecolhido) c.grupoRecolhido = true; } else if (obj.grupoDe) c.grupoDe = obj.grupoDe;      // TASK-195: grupos de linhas (@)
        // TASK-182: marcas da linha — sem isto o undo/redo e o SALVAR do projeto as perdiam (tudo passa por esta função)
        // TASK-032: coordenada de origem (presente em obj._raw na primeira montagem da linha, e
        // em obj diretamente depois) e vínculo cabo<->estrutura — preservados através de
        // undo/redo e salvar/carregar obra, já que tudo passa por esta função.
        const x = obj._x !== undefined ? obj._x : (obj._raw && obj._raw._x);
        const y = obj._y !== undefined ? obj._y : (obj._raw && obj._raw._y);
        if (x !== undefined) c._x = x;
        if (y !== undefined) c._y = y;
        // TASK-091: linha da rede (cabo) e poste (estrutura) para o mapa do vínculo
        const geo = obj._geo || (obj._raw && obj._raw._geo), poste = obj._poste || (obj._raw && obj._raw._poste);
        if (geo) { c._geo = geo; c._geoTipo = obj._geoTipo || (obj._raw && obj._raw._geoTipo); }
        if (poste) c._poste = poste;
        const aprox = obj._posteAprox || (obj._raw && obj._raw._posteAprox);                             // TASK-233: poste sem bloco no CAD (posição aproximada)
        if (Array.isArray(aprox) && aprox.length === 2 && aprox.every(Number.isFinite)) c._posteAprox = aprox.slice();
        for (const [campo, min] of [['_mapaPoste', 0], ['_mapaCabo', -1]]) {
            const par = obj[campo];
            if (Array.isArray(par) && par.length === 2 && par.every(n => Number.isFinite(n) && n >= min && n <= 1)) c[campo] = par.slice();
        }
        for (const campo of ['_mapaPoste500', '_mapaCabo500']) {                                                  // TASK-224: ajuste manual do mapa em 1:500 (deslocamento em metros)
            const par = obj[campo];
            if (Array.isArray(par) && par.length === 2 && par.every(n => Number.isFinite(n) && Math.abs(n) <= 100000)) c[campo] = par.slice();
        }
        for (const campo of ['_mapaId', '_mapaPosteRef']) if (idPontoValido(obj[campo])) c[campo] = obj[campo];
        if (Array.isArray(obj.vinculoEstruturas)) c.vinculoEstruturas = obj.vinculoEstruturas.slice();
        if (obj._mapaTopologia === true) c._mapaTopologia = true;
        if (obj._mapaPosteFicticio === true) {
            c._mapaPosteFicticio = c.congelada = c.bloqueada = true;
            if (Array.isArray(obj._mapaEstruturasFicticias)) c._mapaEstruturasFicticias = obj._mapaEstruturasFicticias.slice();
        }
        if (Array.isArray(obj._mapaAncoras)) c._mapaAncoras = obj._mapaAncoras
            .filter(v => Array.isArray(v) && idPontoValido(v[0]) && typeof v[1] === 'string').map(v => [v[0], v[1]]);
        if (Array.isArray(obj.vinculoPostes) && (obj.vinculoPostes.length || c._mapaTopologia)) c.vinculoPostes = obj.vinculoPostes.slice();                   // Cabos ⇄ Postes
        if (obj.vinculoOrigem) c.vinculoOrigem = obj.vinculoOrigem;                                       // TASK-084: manual | script
        if (Array.isArray(obj.estruturasFicticias) && obj.estruturasFicticias.length) c.estruturasFicticias = obj.estruturasFicticias.slice();   // TASK-095: estruturas fictícias do cabo
        if (obj.origemAvulso) { c.origemAvulso = obj.origemAvulso; if (obj.texto) c.texto = obj.texto; }                 // TASK-171: linha de Outros vinda da sugestão de avulsos da estrutura fictícia
        if (obj.cabosFicticiosSlots && Object.keys(obj.cabosFicticiosSlots).length) c.cabosFicticiosSlots = JSON.parse(JSON.stringify(obj.cabosFicticiosSlots));   // TASK-094: fictícios das demais estruturas da linha
        if (Array.isArray(obj.cabosFicticios) && obj.cabosFicticios.length) c.cabosFicticios = obj.cabosFicticios.slice();   // TASK-084: cabos fictícios da estrutura
        return c;
    }

    /* ═══════════════════════════════════════
       CLASSIFICAÇÃO AUTOMÁTICA DE ENTIDADE
    ═══════════════════════════════════════ */
    /* ═══════════════════════════════════════
       TASK-167 — PARÂMETROS PREENCHIDOS × QUE FALTAM (balão ao passar o mouse + cores opcionais)
       O balão vale sempre; a cor (verde = tudo preenchido, laranja = falta algum) só com a opção ligada em Configurações › Aparência › Avançado.
    ═══════════════════════════════════════ */
    const COR_PARAMS_CHAVE = 'colorir_ativos_parametros';
    function corParametrosLigada() { try { return localStorage.getItem(COR_PARAMS_CHAVE) === '1'; } catch (e) { return false; } }
    window.definirCorParametros = function (ligada) {
        try { localStorage.setItem(COR_PARAMS_CHAVE, ligada ? '1' : '0'); } catch (e) { /* vale só nesta sessão */ }
        document.body.classList.toggle('colorir-parametros', !!ligada);
        aplicarCoresParametros();
    };
    window.corParametrosLigada = corParametrosLigada;

    /** Status (por estrutura) de uma linha de Outros: o que o pedido liga, o que o vínculo injeta, a obra e os padrões da base. */
    // TASK-235: o poste de cada ponto (grupo «@») — vale para o parâmetro POSTE das estruturas do ponto que estão em outra linha (inclusive a linha F)
    const _posteDaLinhaCache = new WeakMap();
    function _posteDaLinhaDoPonto(row) {
        const c = _posteDaLinhaCache.get(row);
        if (c && c.ativo === row.ativo) return c.poste;
        let poste = null;
        try { poste = postesDaLinhaVinculo(row.ativo)[0] || null; } catch (e) { poste = null; }
        _posteDaLinhaCache.set(row, { ativo: row.ativo, poste });
        return poste;
    }
    function _postesDosPontos() {
        const m = new Map();
        for (const r of tableStates.outros.data) {
            if (!r || !r.grupoDe || r.comentario || m.has(r.grupoDe) || !String(r.ativo || '').trim()) continue;
            const p = _posteDaLinhaDoPonto(r);
            if (p) m.set(r.grupoDe, p);
        }
        return m;
    }
    function _ctxComPosteDoPonto(ctx, row, postes) {
        const p = row && row.grupoDe ? (postes || _postesDosPontos()).get(row.grupoDe) : null;
        return p ? Object.assign({}, ctx, { poste_do_ponto: p }) : ctx;
    }
    function _statusParametrosDaLinha(item, pvLinha, assin, ctxBase, obra, defLinha, alinhado) {
        ctxBase = _ctxComPosteDoPonto(ctxBase, item);          // alinhado: um resultado POR ITEM do texto (null quando o ativo não declara parâmetros)          // defLinha (TASK-170): {CODIGO: {define, faltou}} desta linha
        if (!item || !String(item.ativo || '').trim() || typeof AtivoStatus === 'undefined') return [];
        const op = item.operacao || 'I';
        const sem = AnotacaoOutros.interpretarLinha(item.ativo, op, _classicoTotalizadora, ctxBase);
        if (sem.erros.length) return [];
        const com = pvLinha ? AnotacaoOutros.interpretarLinha(item.ativo, op, _classicoTotalizadora, Object.assign({}, ctxBase, { parametros_estrutura: pvLinha })) : sem;
        const lista = sem.itens.map((it, i) => AtivoStatus.statusDoAtivo(((com.itens[i] || it).ativo), it.ativo, assin, obra, defLinha ? defLinha[AtivoParametros.nomeBase(it.ativo)] : undefined) || null);
        return alinhado ? lista : lista.filter(Boolean);
    }
    // Impressão digital do que decide o vínculo (cabos, linhas de Outros, base, configuração): enquanto não muda, o resultado de `parametrosPorEstrutura`/`defineDeEstruturas`
    // (que varrem as duas tabelas inteiras) é reaproveitado — antes era refeito a cada desenho de tabela, a cada tecla com a caixa de parâmetros aberta e a cada passada das cores.
    const _idObj = o => { if (!o || typeof o !== 'object') return 0; let v = _idsObj.get(o); if (!v) { v = ++_seqId; _idsObj.set(o, v); } return v; };
    function _impressaoVinculo() {
        const cfg = window.__vinculacaoConfig;
        let base = null; try { base = orcamentoBaseData; } catch (e) { /* base ainda não declarada */ }
        const lista = (l, campos) => l.map(r => r ? campos.map(k => { const v = r[k]; return Array.isArray(v) ? v.join('.') : (v && typeof v === 'object' ? JSON.stringify(v) : (v == null ? '' : v)); }).join('|') : '').join('\n');
        return [_idObj(cfg), cfg ? _idObj(cfg.postes) : 0, _idObj(base), base ? base.length : 0, localStorage.getItem('projeto_selecionado') || '',
            lista(tableStates.cabos.data, ['ativo', 'operacao', 'vinculoEstruturas', 'vinculoPostes', 'vinculoOrigem', 'congelada', 'fases', 'tensao', 'nivel', 'textoOriginal', 'estruturasFicticias', 'cabosFicticios', 'cabosFicticiosSlots']),
            lista(tableStates.outros.data, ['ativo', 'operacao', 'congelada', 'cabosFicticiosSlots', 'cabosFicticios', 'origemAvulso'])].join('##');
    }
    let _cacheVinculo = null;
    function _vinculoCache() {
        const fp = _impressaoVinculo();
        if (!_cacheVinculo || _cacheVinculo.fp !== fp) _cacheVinculo = { fp, dv: undefined, pv: undefined };
        return _cacheVinculo;
    }
    function _defineDoVinculo() {                                                         // TASK-170: de onde veio cada parâmetro do `define`
        const cfg = window.__vinculacaoConfig;
        if (!cfg || typeof VinculacaoRegras === 'undefined' || !VinculacaoRegras.defineDeEstruturas) return {};
        const c = _vinculoCache();
        if (c.dv === undefined) c.dv = VinculacaoRegras.defineDeEstruturas(tableStates.cabos.data, tableStates.outros.data, cfg, _prefixoCaboVinculo, _codigoEstruturaVinculo, _codigosEstruturaVinculo);
        return c.dv;
    }
    function _contextoParametros() {
        const proj = localStorage.getItem('projeto_selecionado') || '', base = orcamentoBaseData;
        if (!_cacheAssin || _cacheAssin.base !== base || _cacheAssin.n !== base.length || _cacheAssin.proj !== proj) _cacheAssin = { base, n: base.length, proj, assin: AtivoParametros.assinaturasDoProjeto(base, proj) };
        return { assin: _cacheAssin.assin, ctx: _ctxOutros(), obra: window.ObraParams ? ObraParams.valores() : null };
    }

    /** Nomes que o cálculo resolve na base (ativo/apelido → componente → código) para o projeto e a origem; recalcula só quando a base ou o projeto mudam. */
    let _cacheNomesBase = null;
    function _nomesDaBaseDaTela(origem) {
        const proj = localStorage.getItem('projeto_selecionado') || '';
        if (!_cacheNomesBase || _cacheNomesBase.base !== orcamentoBaseData || _cacheNomesBase.proj !== proj || _cacheNomesBase.n !== orcamentoBaseData.length) _cacheNomesBase = { base: orcamentoBaseData, proj, n: orcamentoBaseData.length, por: {} };
        return _cacheNomesBase.por[origem] || (_cacheNomesBase.por[origem] = AtivoNaBase.nomesDaBase(orcamentoBaseData, proj, origem));
    }
    /** Os ativos (nomes) que a linha de Cabos/Outros pede, como o cálculo os lê: notação compacta, apelidos e notação do LABEL já traduzidos. */
    function _ativosPedidosDaLinha(tipo, row, ctxBase) {
        if (!row || !String(row.ativo || '').trim()) return [];
        if (tipo === 'cabos') {
            if (AnotacaoCabos.ehNotacao(row.ativo)) { const r = AnotacaoCabos.interpretar(row.ativo); return r.erros.length ? [] : r.itens.map(i => i.ativo); }
            const p = _prefixoCaboVinculo(row.ativo);
            return p ? [p] : [];
        }
        const sem = AnotacaoOutros.interpretarLinha(row.ativo, row.operacao || 'I', _classicoTotalizadora, ctxBase);
        return sem.erros.length ? [] : sem.itens.map(i => i.ativo);
    }
    /** Pinta o ativo das linhas de Outros e Cabos (só com a opção ligada; desligada, tira as cores): Outros → verde/laranja pelos parâmetros; vermelho = o ativo não existe na base (Cabos e Outros).
     *  Linha congelada ou removendo (R, *R) nunca é pintada (TASK-247). */
    function aplicarCoresParametros() {
        const inputs = document.querySelectorAll('#table-outros .inp-ativo, #table-cabos .inp-ativo');
        const ligada = corParametrosLigada();
        document.body.classList.toggle('colorir-parametros', ligada);
        const limpar = i => { i.classList.remove('par-ok', 'par-falta', 'par-nao-base'); if (i.dataset.titNaoBase) { i.title = i.dataset.titNaoBase === '-' ? '' : i.dataset.titNaoBase; delete i.dataset.titNaoBase; } };
        if (!ligada) { inputs.forEach(limpar); return; }
        let c;
        try { c = _contextoParametros(); } catch (e) { return; }                       // base ainda carregando: pinta quando chegar
        const pv = _parametrosVinculoPorLinha(), dv = _defineDoVinculo();
        const baseCarregada = orcamentoBaseData.length > 0;
        // resultado por linha guardado (texto + operação + o que vem do vínculo/obra/base): só as linhas que mudaram são calculadas de novo
        const chaveGlobal = [_idObj(orcamentoBaseData), orcamentoBaseData.length, _idObj(window.__vinculacaoConfig), localStorage.getItem('projeto_selecionado') || '', JSON.stringify(c.obra || null)].join('|');
        if (!_cacheCores || _cacheCores.g !== chaveGlobal) _cacheCores = { g: chaveGlobal, m: new Map() };
        inputs.forEach(inp => {
            const tipo = inp.dataset.parTipo === 'cabos' ? 'cabos' : 'outros';
            const idx = Number(inp.dataset.parIdx), row = tableStates[tipo].data[idx];
            limpar(inp);
            if (!row || row.congelada || ['R', '*R'].includes(String(row.operacao || '').trim().toUpperCase())) return;          // TASK-247: removendo não entra nas cores
            const chave = tipo + '|' + row.ativo + '|' + (row.operacao || '') + '|' + (tipo === 'outros' ? JSON.stringify(pv[idx] || 0) + '|' + JSON.stringify(dv[idx] || 0) : '');
            let r = _cacheCores.m.get(chave);
            if (r === undefined) {
                r = '';
                if (baseCarregada) {
                    let fora = [];
                    try { fora = AtivoNaBase.foraDaBase(_nomesDaBaseDaTela(tipo === 'cabos' ? 'CABOS' : 'OUTROS'), _ativosPedidosDaLinha(tipo, row, c.ctx)); } catch (e) { fora = []; }
                    if (fora.length) r = 'par-nao-base';
                }
                if (!r && tipo === 'outros') {
                    const estado = AtivoStatus.estadoDaLinha(_statusParametrosDaLinha(row, pv[idx], c.assin, c.ctx, c.obra, dv[idx]));
                    r = estado === 'ok' ? 'par-ok' : (estado === 'falta' ? 'par-falta' : '');
                }
                _cacheCores.m.set(chave, r);
                if (_cacheCores.m.size > 5000) _cacheCores.m.clear();
            }
            if (r) inp.classList.add(r);                                    // vermelho: a explicação ("ATIVO NÃO EXISTE NA BASE") aparece no balão ao passar o mouse (_statusDoAlvoDoBalao)
        });
    }
    // Depois de desenhar uma tabela, as cores são aplicadas UMA vez (as duas tabelas costumam ser desenhadas em sequência), fora do caminho do desenho.
    function agendarCoresAposDesenhar() {
        if (!corParametrosLigada() && !document.body.classList.contains('colorir-parametros')) return;
        if (_tmCoresRender) return;
        _tmCoresRender = setTimeout(() => { _tmCoresRender = 0; aplicarCoresParametros(); }, 0);
    }
    window.aplicarCoresParametros = aplicarCoresParametros;
    let _tmCoresParametros = null;
    // TASK-247: a cor acompanha a edição — digitando, recalcula logo após a última tecla (antes, 350 ms); gravando (balão, operação), na hora
    function agendarCoresParametros() { if (!corParametrosLigada()) return; clearTimeout(_tmCoresParametros); _tmCoresParametros = setTimeout(aplicarCoresParametros, 60); }
    function aplicarCoresJa() { if (!corParametrosLigada()) return; clearTimeout(_tmCoresParametros); aplicarCoresParametros(); }

    function _htmlBalaoParametros(status) {
        const esc = t => String(t == null ? '' : t).replace(/[&<>"]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[m]));
        return status.map(s => {
            if (s.naoExiste) return `<div class="est"><b class="h">${esc(s.ativo)}</b> <span class="f">· ATIVO NÃO EXISTE NA BASE</span></div>`;
            const cab = s.completo ? `<span class="v">· completo (${s.total} de ${s.total})</span>` : `<span class="f">· faltam ${s.faltam + s.invalidos} de ${s.total}</span>`;
            const linhas = s.params.map(p => {
                if (p.falta) return `<div><span class="f">⚠</span> ${esc(p.nome)} <span class="f">· falta</span> <span class="m">— ${esc(p.motivo)}</span></div>`;
                if (p.invalido) return `<div><span class="f">⚠</span> ${esc(p.nome)} = ${esc(p.valor)} <span class="f">· inválido</span> <span class="m">— ${esc(p.motivo)}</span></div>`;
                return `<div><span class="v">✓</span> ${esc(p.nome)} = ${esc(p.valor)} <span class="m">· ${esc(AtivoStatus.rotuloOrigem(p.origem, p.detalhe))}</span></div>`;
            }).join('');
            return `<div class="est"><b class="h">${esc(s.ativo)}</b> ${cab}${linhas}</div>`;
        }).join('');
    }
    /** Ativos pedidos pela linha que NÃO existem na base (nem como ativo/apelido, nem como componente ou código) → itens do balão "ATIVO NÃO EXISTE NA BASE". Só com a verificação ligada. */
    function _foraDaBaseDoAlvo(tipo, idx, c) {
        if (!corParametrosLigada() || !orcamentoBaseData.length) return [];
        const row = tableStates[tipo].data[idx];
        if (!row || row.congelada) return [];
        const fora = AtivoNaBase.foraDaBase(_nomesDaBaseDaTela(tipo === 'cabos' ? 'CABOS' : 'OUTROS'), _ativosPedidosDaLinha(tipo, row, c.ctx));
        return Array.from(new Set(fora)).map(nome => ({ ativo: nome, naoExiste: true }));
    }
    function _statusDoAlvoDoBalao(el) {
        const idx = Number(el.dataset.parIdx);
        const c = _contextoParametros();
        if (el.dataset.parTipo === 'cabos') return _foraDaBaseDoAlvo('cabos', idx, c);          // Cabos: só o aviso de "ativo fora da base"; parâmetros são coisa de Outros
        if (el.dataset.parTipo === 'outros') return _foraDaBaseDoAlvo('outros', idx, c).concat(_statusParametrosDaLinha(tableStates.outros.data[idx], _parametrosVinculoPorLinha()[idx], c.assin, c.ctx, c.obra, _defineDoVinculo()[idx]));
        const lin = tableStates.totalizadora.data[idx];                                  // Totalizadora: o ativo já está como o cálculo o enxerga
        const st = lin && AtivoStatus.statusDoAtivo(lin.ativo, lin.ativo, c.assin, c.obra);
        return st ? [st] : [];
    }


    /* ═══════════════════════════════════════
       TASK-171 — AVULSOS DA ESTRUTURA FICTÍCIA (alças, conectores) → tabela Outros, com um clique
    ═══════════════════════════════════════ */
    let _cacheAvulsos = null;
    function _gruposEAssinaturas() {                                                   // recalcula só quando a base muda
        const proj = localStorage.getItem('projeto_selecionado') || '';
        if (!_cacheAvulsos || _cacheAvulsos.base !== orcamentoBaseData || _cacheAvulsos.proj !== proj) {
            _cacheAvulsos = { base: orcamentoBaseData, proj, assin: AtivoParametros.assinaturasDoProjeto(orcamentoBaseData, proj), grupos: AvulsosFicticia.gruposDeChamada(orcamentoBaseData, proj) };
        }
        return _cacheAvulsos;
    }
    function _avulsosDoCabo(ic) {
        const cfg = window.__vinculacaoConfig, cabo = tableStates.cabos.data[ic];
        if (!cfg || !cabo || typeof AvulsosFicticia === 'undefined' || !VinculacaoRegras.estruturasFicticiasDe(cabo).length) return [];
        const c = _gruposEAssinaturas();
        return AvulsosFicticia.avulsosDoCabo(cabo, cfg, _prefixoCaboVinculo, c.assin, c.grupos).map(x => Object.assign({}, x, AvulsosFicticia.estadoDosAvulsos(ic, x.ficticia, x.itens, tableStates.outros.data)));
    }
    function _adicionarAvulsos(ic, ficticia) {
        const lista = _avulsosDoCabo(ic).find(x => x.ficticia === String(ficticia).toUpperCase());
        const itens = lista ? lista.itens.filter(i => !i.completar) : [];
        if (!itens.length) { showToast('Nada a adicionar: complete o cabo existente nos itens marcados.'); return 0; }
        pushHistory();
        const chave = AvulsosFicticia.chaveOrigem(ic, ficticia);
        tableStates.outros.data.forEach(r => { if (r && r.origemAvulso === chave) r.origemAvulso = chave + '!velho'; });          // a sugestão anterior fica como "avulso desatualizado"
        itens.forEach(i => tableStates.outros.data.push({ entidade: 'OUTROS', operacao: i.operacao, ativo: AvulsosFicticia.textoDaLinha(i), qtdAtivos: i.qtd, texto: `avulso da ${ficticia} fictícia — cabo #${ic + 1}`, origemAvulso: chave }));
        renderTable('outros'); atualizarResumoRedeUI();
        if (window.GridRProjeto && GridRProjeto.marcarSujo) GridRProjeto.marcarSujo(false);
        showToast(`${itens.length} avulso(s) da ${ficticia} adicionado(s) em Outros.`);
        return itens.length;
    }
    /** Índices das linhas de Outros de uma sugestão desatualizada (o cabo ou a fictícia mudou). */
    function _indicesAvulsoVelho() {
        const out = new Set(), dados = tableStates.outros.data;
        if (!dados.some(r => r && r.origemAvulso)) return out;
        dados.forEach((r, i) => { if (r && r.origemAvulso && r.origemAvulso.endsWith('!velho')) out.add(i); });
        try { tableStates.cabos.data.forEach((c, ic) => { if (c) _avulsosDoCabo(ic).forEach(x => { if (x.estado === 'desatualizado') x.linhas.forEach(i => out.add(i)); }); }); } catch (e) { /* base ainda carregando */ }
        return out;
    }

    /** TASK-081: contexto da base (apelidos, postes, assinaturas) para traduzir a notação compacta de Outros. */
    // O contexto da base (apelidos, postes, assinaturas) varre a base INTEIRA (milhares de linhas) e era refeito a cada chamada — várias vezes por tecla ao preencher parâmetros.
    // Agora é guardado e só refeito quando a base, o projeto ou a convenção de postes mudam.
    var _cacheCtxOutros = null;                       // var: a função é chamada já na montagem inicial das tabelas (antes desta linha executar)
    function _ctxOutros() {
        let base = [];
        try { base = orcamentoBaseData || []; } catch (e) { /* a base ainda não foi declarada (a tela está carregando) */ }
        const proj = (typeof localStorage !== 'undefined' ? localStorage.getItem('projeto_selecionado') : '') || '';
        const postes = window.__vinculacaoConfig ? window.__vinculacaoConfig.postes : undefined;                         // TASK-174: convenção de postes do projeto
        const c = _cacheCtxOutros;
        if (c && c.base === base && c.n === base.length && c.proj === proj && c.postes === postes) return c.ctx;
        _cacheCtxOutros = { base, n: base.length, proj, postes, ctx: AnotacaoOutros.contextoDaBase(base, proj, postes) };
        return _cacheCtxOutros.ctx;
    }

    /** Leitor do formato ANTIGO usado pela Totalizadora: token a token, "qtd-ativo" (o "*" na frente da quantidade é negativo). */
    function _classicoTotalizadora(texto) {
        const saida = [];
        AtivoParametros.dividirTokens(texto).forEach(p => {
            if (!p.trim()) return;
            let q = 1, aName = p.trim();
            const m = p.match(/^([\*\-]?\d+(?:[.,]\d+)?)[Xx\-](.+)$/i);
            if (m) {
                let qStr = m[1], neg = false;
                if (qStr.startsWith('*') || qStr.startsWith('-')) { neg = true; qStr = qStr.substring(1); }
                q = parseFloat(qStr.replace(',', '.'));
                if (neg) q = -q;
                aName = m[2].trim();
            }
            saida.push({ ativo: aName.toUpperCase(), qtd: q });
        });
        return saida;
    }

    /**
     * Extrai o prefixo do ativo e o comprimento (metros) de uma linha da tabela Cabos, no formato
     * "ATIVO FASE COMPRIMENTO". Mesma normalização de calcularQtdAtivos() (linha ~903): funde
     * prefixos partidos por espaço ("CAA 2" → "CAA2") antes de separar os tokens, senão o
     * comprimento e a fase seriam lidos errado nesse formato. O comprimento retornado é o valor
     * bruto (último token) — nunca multiplicado por qtdAtivos/fases (RULES.md Regra 8, confirmado
     * pelo usuário em TASK-008-29-09-2026.md).
     */
    function extrairPrefixoEComprimentoCabo(ativoRaw) {
        if (!ativoRaw || !ativoRaw.trim()) return null;
        if (AnotacaoCabos.ehNotacao(ativoRaw)) {                                                  // TASK-083: código do cabo principal + comprimento
            const pc = AnotacaoCabos.prefixoEComprimento(ativoRaw);
            return pc ? { prefixo: pc[0], comprimento: pc[1] } : null;
        }
        const txt = ativoRaw.trim().toUpperCase();
        const txtNorm = txt
            .replace(/\//g, '')
            .replace(/^CAA\s+(\d)/i, 'CAA$1')
            .replace(/^CA\s+(\d)/i, 'CA$1')
            .replace(/^CU\s+(\d)/i, 'CU$1')
            .replace(/^CAZ\s+(\d)/i, 'CAZ$1')
            .replace(/^P\s+(\d)/i, 'P$1');
        const cleaned = txtNorm.replace(/\s+M\s*$/i, '').trim();
        const tokens = cleaned.split(/\s+/).filter(Boolean);
        if (tokens.length < 2) return null;  // sem comprimento identificável
        const prefixo = tokens[0];
        const comprimento = parseFloat(tokens[tokens.length - 1].replace(',', '.'));
        if (isNaN(comprimento)) return null;
        return { prefixo, comprimento };
    }

    /** TASK-116: formato antigo com o cabo reconhecido mas sem número no fim ("CAA2 ABC"): devolve só o prefixo; com comprimento ou vazio → null. */
    function extrairPrefixoCaboSemComprimento(ativoRaw) {
        if (!ativoRaw || !ativoRaw.trim() || extrairPrefixoEComprimentoCabo(ativoRaw)) return null;
        const tokens = ativoRaw.trim().toUpperCase().replace(/\//g, '').replace(/^(CAA|CA|CU|CAZ|P)\s+(\d)/, '$1$2').replace(/\s+M\s*$/, '').trim().split(/\s+/).filter(Boolean);
        if (!tokens.length || !isNaN(parseFloat(tokens[tokens.length - 1].replace(',', '.')))) return null;
        return tokens[0];
    }

    /** TASK-094: códigos de TODAS as estruturas da linha, sem o poste (`9400D, N1, N4` → [N1, N4]); só poste → []. */
    function _codigosEstruturaVinculo(ativoEstrutura) {
        return AnotacaoOutros.codigosEstruturas((ativoEstrutura || '').trim(), _classicoTotalizadora, _ctxOutros());
    }

    function _codigoEstruturaVinculo(ativoEstrutura) {
        // TASK-081/084: o código da estrutura é o do primeiro item da linha que NÃO é poste (9200D, N1 → N1; 1-U4 → U4); linha só com poste → o próprio poste.
        const texto = (ativoEstrutura || '').trim();
        const viaModulo = AnotacaoOutros.codigoEstrutura(texto, _classicoTotalizadora, _ctxOutros());
        if (viaModulo) return viaModulo;
        // TASK-036: o ativo de uma linha Outros é "<qtd>-<código>" (ex.: "1-U4"), nunca o código puro.
        const semQtd = texto.match(/^[*-]?[0-9]+(?:[.,][0-9]+)?[Xx-](.+)$/);
        const codigo = semQtd ? semQtd[1] : texto;
        return codigo.indexOf('(') >= 0 ? codigo.slice(0, codigo.indexOf('(')).trim() : codigo;   // TASK-079: N1(XLPE,DT) → N1
    }

    function _prefixoCaboVinculo(ativo) {
        const info = extrairPrefixoEComprimentoCabo(ativo);
        return info ? info.prefixo : null;
    }

    /** TASK-084/096: {índice em Outros: {CODIGO_DA_ESTRUTURA: {CABO: "C" | ["A35P", "C16"]}}} — os parâmetros que cada estrutura da linha recebe do(s) cabo(s) vinculado(s) (tabelas de cabos/estruturas do projeto). */
    function _parametrosVinculoPorLinha() {
        const cfg = window.__vinculacaoConfig;
        if (!cfg || typeof VinculacaoRegras === 'undefined') return {};
        const c = _vinculoCache();
        if (c.pv === undefined) c.pv = VinculacaoRegras.parametrosPorEstrutura(tableStates.cabos.data, tableStates.outros.data, cfg, _prefixoCaboVinculo, _codigoEstruturaVinculo, _codigosEstruturaVinculo);
        return c.pv;
    }

    // autoClassifyEntidade foi religada ao motor de regras do leitor (TASK-006), reusando a
    // mesma Tabela de Classificacao da aba Leitor. Operacao fixa em 'I' (neutro) porque esta
    // funcao so recebe o texto do ativo, sem operacao real conhecida -- nenhuma regra de
    // classificacao hoje exige um valor de operacao que 'I' nao satisfaca. O texto do ativo e
    // passado tambem como "texto bruto" porque e exatamente isso que o codigo original
    // verificava aqui (ver .ai/tasks/TASK-006-25-09-2026.md).
    function autoClassifyEntidade(ativoTexto) {
        if (!ativoTexto) return '0';
        const classificar = t => RegrasLeitorEngine.classificar('I', t, '', '', t, window.__regrasLeitorClassificacao || []).entidade;
        const original = String(ativoTexto).replace(/\([^()]*\)/g, '');             // TASK-079: N1(XLPE,DT) classifica como N1
        const e = classificar(original);
        // TASK-081/082: linha compacta ("9200D, N1") é classificada pelo texto como foi escrito (regras novas, ex.: Pernambuco); se nenhuma regra a reconhece,
        // tenta o formato antigo equivalente ("1-DT9/200 1-N1"), como as regras dos outros projetos sempre leram.
        if (e === '0' && AnotacaoOutros.ehCompacta(ativoTexto)) return classificar(AnotacaoOutros.textoEquivalente(String(ativoTexto), _classicoTotalizadora, _ctxOutros()).replace(/\([^()]*\)/g, ''));
        return e;
    }

    /** TASK-065: depois de editar/colar o ativo de uma linha de Outros, a entidade acompanha (mesmo motor do Leitor).
     *  Linha travada à mão (`entidadeManual`) não muda; se a regra não reconhece o ativo, mantém a entidade que já tinha. */
    function reclassificarEntidade(row) {
        if (!row || row.entidadeManual) return false;
        const nova = autoClassifyEntidade(row.ativo);
        if (nova && nova !== '0' && nova !== row.entidade) { row.entidade = nova; return true; }
        return false;
    }


    /* ═══════════════════════════════════════
       CÁLCULO DE QTD ATIVOS (CABOS)
    ═══════════════════════════════════════ */
    /**
     * Calcula a quantidade de ativos de uma linha de cabo.
     * @param {string} ativoTexto - Texto do campo Ativo (ex: "CAA 2 ABC 35 m")
     * @param {string|null} faseFromNext - Fase herdada da próxima linha (para linhas standalone)
     * @returns {number} Quantidade de ativos calculada
     */
    function calcularQtdAtivos(ativoTexto, faseFromNext) {
        if (!ativoTexto || !ativoTexto.trim()) return 0;
        if (AnotacaoCabos.ehNotacao(ativoTexto)) return AnotacaoCabos.fasesDaLinha(ativoTexto);   // TASK-083: 3xA35P... = 3 ativos

        let txt = ativoTexto.trim().toUpperCase();

        // Normaliza prefixos: "CAA 2" → "CAA2", "CA 4" → "CA4", "P 50" → "P50"
        const txtNorm = txt
            .replace(/\//g, '')
            .replace(/^CAA\s+(\d)/i, 'CAA$1')
            .replace(/^CA\s+(\d)/i, 'CA$1')
            .replace(/^CU\s+(\d)/i, 'CU$1')
            .replace(/^CAZ\s+(\d)/i, 'CAZ$1')
            .replace(/^P\s+(\d)/i, 'P$1');

        // Remove 'm' final (marcador de metros)
        const cleaned = txtNorm.replace(/\s+M\s*$/i, '').trim();
        const tokens = cleaned.split(/\s+/);

        if (tokens.length === 0) return 0;

        const prefixo = tokens[0];

        // ── Regra M ou CAZ: sempre 1 ──
        if (/^M\d/i.test(prefixo) || /^M$/i.test(prefixo) || /^CAZ/i.test(prefixo)) {
            return 1;
        }

        // Determinar se é uma linha "standalone" (apenas ativo, sem fase/comprimento)
        // Ex: "P50" → 1 token; "CAA2" → 1 token
        // Uma linha completa tem pelo menos: ATIVO FASE COMPRIMENTO (3 tokens)
        let fase = null;

        if (tokens.length >= 3) {
            // Formato padrão: ATIVO FASE COMPRIMENTO
            // O último token (após remover 'm') é o comprimento
            // O(s) token(s) do meio formam a fase
            fase = tokens.slice(1, tokens.length - 1).join('');
        } else if (tokens.length === 1) {
            // Linha standalone - herda da próxima linha
            if (faseFromNext) {
                fase = faseFromNext;
            } else {
                return 1; // Sem info, retorna 1 como fallback
            }
        } else if (tokens.length === 2) {
            // Poderia ser ATIVO+COMPRIMENTO (sem fase) ou ATIVO+FASE (sem comprimento)
            // Tenta: se o segundo token for numérico, é comprimento sem fase
            const second = tokens[1];
            if (/^[\d.,]+$/.test(second)) {
                // Apenas comprimento, sem fase → standalone behavior
                if (faseFromNext) {
                    fase = faseFromNext;
                } else {
                    return 1;
                }
            } else {
                // Segundo token é a fase (sem comprimento)
                fase = second;
            }
        }

        if (!fase) return 1;

        const faseLen = fase.length;

        // ── Regra CAA2 / CAA 2 com fase de 1 char → +1 ──
        if (/^CAA2/i.test(prefixo) && faseLen === 1) {
            return faseLen + 1; // = 2
        }

        // ── Regra CA4 / CA 4 → sempre +1 ──
        if (/^CA4/i.test(prefixo)) {
            return faseLen + 1;
        }

        // ── Caso padrão: len(fase) ──
        return faseLen || 1;
    }

    /**
     * Extrai a fase de um texto de ativo completo (para herança de linhas standalone).
     * Retorna null se não conseguir extrair.
     */
    function extrairFase(ativoTexto) {
        if (!ativoTexto || !ativoTexto.trim()) return null;
        if (AnotacaoCabos.ehNotacao(ativoTexto)) return 'A'.repeat(AnotacaoCabos.fasesDaLinha(ativoTexto));   // TASK-083

        let txt = ativoTexto.trim().toUpperCase();
        const txtNorm = txt
            .replace(/\//g, '')
            .replace(/^CAA\s+(\d)/i, 'CAA$1')
            .replace(/^CA\s+(\d)/i, 'CA$1')
            .replace(/^CU\s+(\d)/i, 'CU$1')
            .replace(/^CAZ\s+(\d)/i, 'CAZ$1')
            .replace(/^P\s+(\d)/i, 'P$1');

        const cleaned = txtNorm.replace(/\s+M\s*$/i, '').trim();
        const tokens = cleaned.split(/\s+/);

        if (tokens.length >= 3) {
            return tokens.slice(1, tokens.length - 1).join('');
        }
        return null;
    }

    /**
     * Verifica se uma linha de cabo é "standalone" (apenas ativo, sem fase/comprimento).
     */
    function isStandaloneLine(ativoTexto) {
        if (!ativoTexto || !ativoTexto.trim()) return false;
        if (AnotacaoCabos.ehNotacao(ativoTexto)) return false;                                    // TASK-083: a notação traz fase e comprimento
        let txt = ativoTexto.trim().toUpperCase();
        const txtNorm = txt
            .replace(/\//g, '')
            .replace(/^CAA\s+(\d)/i, 'CAA$1')
            .replace(/^CA\s+(\d)/i, 'CA$1')
            .replace(/^CU\s+(\d)/i, 'CU$1')
            .replace(/^CAZ\s+(\d)/i, 'CAZ$1')
            .replace(/^P\s+(\d)/i, 'P$1');
        const cleaned = txtNorm.replace(/\s+M\s*$/i, '').trim();
        const tokens = cleaned.split(/\s+/);

        if (tokens.length === 1) return true;
        if (tokens.length === 2 && /^[\d.,]+$/.test(tokens[1])) return true;
        return false;
    }

    /**
     * Recalcula qtdAtivos para todas as linhas da tabela de cabos.
     * Processa de baixo para cima para resolver dependências de linhas standalone.
     */
    function recalcAllQtdAtivos(data = tableStates.cabos.data) {
        for (let i = data.length - 1; i >= 0; i--) {
            const row = data[i];
            if (!row || row.congelada) continue;                                  // linha congelada: nenhum algoritmo mexe nela
            let faseFromNext = null;
            if (isStandaloneLine(row.ativo) && i + 1 < data.length && !data[i + 1].congelada) {
                faseFromNext = extrairFase(data[i + 1].ativo);
            }
            
            const isNegative = row.qtdAtivos !== undefined && row.qtdAtivos !== null && row.qtdAtivos.toString().trim().startsWith('-');
            let newVal = calcularQtdAtivos(row.ativo, faseFromNext);
            
            if (isNegative && newVal > 0) {
                row.qtdAtivos = "-" + newVal;
            } else {
                row.qtdAtivos = newVal;
            }
        }
    }

    /**
     * Recalcula qtdAtivos apenas para a linha alterada e a anterior (se for standalone).
     */
    function recalcRowAndAbove(index, data = tableStates.cabos.data) {
        if (index < 0 || index >= data.length) return;

        function calc(i) {
            if (i < 0 || i >= data.length) return;
            const r = data[i];
            if (!r || r.congelada) return;
            let faseNext = null;
            if (isStandaloneLine(r.ativo) && i + 1 < data.length && !data[i + 1].congelada) {
                faseNext = extrairFase(data[i + 1].ativo);
            }
            const isNeg = r.qtdAtivos !== undefined && r.qtdAtivos !== null && r.qtdAtivos.toString().trim().startsWith('-');
            let nv = calcularQtdAtivos(r.ativo, faseNext);
            r.qtdAtivos = (isNeg && nv > 0) ? "-" + nv : nv;
        }

        calc(index);
        if (index - 1 >= 0 && data[index - 1] && isStandaloneLine(data[index - 1].ativo)) {
            calc(index - 1);
        }
    }

    /* ═══════════════════════════════════════
       INTEGRAÇÃO IA (GEMINI) E MEMÓRIA
    ═══════════════════════════════════════ */
    /* ═══════════════════════════════════════
       OBRAS: operações reutilizadas pelo modal "Carregar Obra" e pelos comandos da IA (TASK-028)
    ═══════════════════════════════════════ */
    /** Linhas de uma tabela de um snapshot salvo (aceita {cabos:{data:[..]}} e {cabos:[..]}). */
    function linhasDoSnap(snap, tipo) {
        const t = snap && snap[tipo];
        const lista = Array.isArray(t) ? t : (t && Array.isArray(t.data) ? t.data : []);
        return lista.filter(Boolean);
    }

    function compararObraNoResumo(obra, nome) {
        if (!obra || !['cabos', 'outros'].some(tipo => Array.isArray(obra[tipo]) || Array.isArray(obra[tipo]?.data))) {
            showToast('Dados da obra inválidos.');
            return false;
        }
        const atual = { cabos: tableStates.cabos.data.slice(), outros: tableStates.outros.data.slice() };
        const resultado = window.CompararObras.comparar(atual, obra, {
            cabo: extrairPrefixoEComprimentoCabo,
            codigoOutro: linha => _codigoEstruturaVinculo(linha.ativo),
            estruturas: linha => VinculacaoRegras.codigosDaLinha(linha, _codigoEstruturaVinculo, _codigosEstruturaVinculo)
        });
        window.CompararObras.abrir(resultado, {
            nome, referencia: window.GridRProjeto?.nome(),
            localizar(tipo, i) {
                const linha = atual[tipo][i];
                const indice = tableStates[tipo].data.indexOf(linha);
                const tr = indice < 0 ? null : bodyDe(tipo).querySelector(`tr[data-index="${indice}"]`);
                if (!tr) {
                    showToast('A linha da referência mudou. Faça uma nova comparação para localizá-la.');
                    return false;
                }
                mostrarTabelasDoResumo(tipo);
                // Revela só a linha localizada. Os filtros e grupos continuam valendo para o restante.
                tr.style.display = '';
                tr.scrollIntoView({ block: 'center', inline: 'nearest' });
                focarCelula(tipo, indice, 1);
                return true;
            }
        });
        return true;
    }

    function reidentificarPontosCopiados(linhas, cabos = []) {
        const ids = new Map();
        for (const linha of linhas) {
            if (idPontoValido(linha._mapaId)) ids.set(linha._mapaId, ids.has(linha._mapaId) ? null : idPontoMapa());
        }
        for (const linha of linhas) {
            const id = ids.get(linha._mapaId), ref = ids.get(linha._mapaPosteRef);
            if (id) linha._mapaId = id; else delete linha._mapaId;
            if (ref) linha._mapaPosteRef = ref; else delete linha._mapaPosteRef;
        }
        for (const cabo of cabos) {
            if (Array.isArray(cabo._mapaAncoras)) cabo._mapaAncoras = cabo._mapaAncoras
                .filter(([id]) => ids.get(id)).map(([id, codigo]) => [ids.get(id), codigo]);
        }
    }

    /** Acrescenta as linhas da obra às tabelas atuais (um passo de histórico). */
    function adicionarObraAoProjeto(snap, anotacoesNanoCAD = false) {
        if (!snap) { showToast('Dados da obra inválidos.'); return false; }
        baselineManual = null;
        pushHistory();  // snapshot ANTES de adicionar
        const deslocamento = tableStates.outros.data.length;
        const qtdOutros = linhasDoSnap(snap, 'outros').length;
        let cabosCopiados = [];
        ['cabos', 'outros'].forEach(type => {
            const linhas = linhasDoSnap(snap, type);
            if (!linhas.length) return;
            const novas = linhas.map(row => JSON.parse(JSON.stringify(row)));
            if (type === 'outros' && novas.some(row => row._mapaId || row._mapaPosteRef)) reidentificarPontosCopiados(novas, cabosCopiados);
            if (window.GruposLinhas) window.GruposLinhas.reidentificar(novas);                   // TASK-195: grupos da obra ganham ids novos (a mesma obra pode entrar duas vezes)
            if (type === 'cabos') {
                cabosCopiados = novas;
                for (const row of novas) {
                    for (const campo of ['vinculoEstruturas', 'vinculoPostes']) {
                        if (Array.isArray(row[campo])) row[campo] = row[campo]
                            .filter(id => Number.isFinite(id) && id >= 0 && Math.floor(id) < qtdOutros)
                            .map(id => Number((id + deslocamento).toFixed(2)));
                    }
                }
                if (anotacoesNanoCAD) recalcAllQtdAtivos(novas);
            }
            novas.forEach(row => tableStates[type].data.push(row));
            if (type === 'cabos' && !anotacoesNanoCAD) recalcAllQtdAtivos();
            renderTable(type);
            refreshAllFilters(type);
        });
        buildAtivoSets();
        buildDataLists();
        return true;
    }

    /** Acrescenta as linhas da obra com sinal invertido (retirada) — um passo de histórico. */
    function subtrairObraDoProjeto(snap) {
        if (!snap) { showToast('Dados da obra inválidos.'); return false; }
        baselineManual = null;
        pushHistory();  // snapshot ANTES de subtrair
        ['cabos', 'outros'].forEach(type => {
            const linhas = linhasDoSnap(snap, type);
            if (!linhas.length) return;
            const clonadas = linhas.map(row => JSON.parse(JSON.stringify(row)));
            if (type === 'outros') reidentificarPontosCopiados(clonadas);
            if (window.GruposLinhas) window.GruposLinhas.reidentificar(clonadas);
            clonadas.forEach(clonada => {
                if (clonada.grupoDecl) { tableStates[type].data.push(clonada); return; }       // TASK-195: a declaração do grupo («@P1») continua sendo só um título
                if (type === 'cabos') {
                    delete clonada._mapaAncoras;
                    clonada.qtdAtivos = "-" + Math.abs(clonada.qtdAtivos || 0);
                }
                else clonada.ativo = "*" + clonada.ativo;
                tableStates[type].data.push(clonada);
            });
            renderTable(type);
            refreshAllFilters(type);
        });
        buildAtivoSets();
        buildDataLists();
        return true;
    }

    /* ═══ MODELOS COM VARIÁVEL V (TASK-058) ═══ */
    const projetoAtualObras = () => localStorage.getItem('projeto_selecionado_codigo') || '229';

    async function mensagensDeErro(r) {
        let j = null;
        try { j = await r.json(); } catch (e) { /* corpo vazio */ }
        const d = j && j.detail;
        if (d && Array.isArray(d.erros)) return d.erros;
        return [typeof d === 'string' ? d : `Erro HTTP ${r.status}`];
    }

    /** Janela "Parâmetros do modelo": pede os valores das variáveis V, o servidor gera a obra padrão (sem V) e devolve
     *  {cabos:[...], outros:[...]}. Resolve null se o usuário cancelar. Nada é gravado; o modelo não muda. */
    async function pedirParametrosModelo(obra) {
        let info;
        try {
            const r = await fetch(`/api/obras/${encodeURIComponent(obra.id)}/modelo?projeto=${encodeURIComponent(projetoAtualObras())}`);
            if (!r.ok) throw new Error((await mensagensDeErro(r)).join(' '));
            info = await r.json();
        } catch (e) {
            showToast('Não foi possível abrir o modelo: ' + e.message);
            return null;
        }
        return new Promise(resolve => {
            const modal = document.getElementById('modal-parametros-modelo');
            const campos = document.getElementById('param-modelo-campos');
            const erro = document.getElementById('param-modelo-erro');
            document.getElementById('param-modelo-nome').textContent = `— ${info.nome}`;
            erro.textContent = '';
            campos.replaceChildren();
            const inputs = [];
            if ((info.erros || []).length) erro.textContent = info.erros.join(' ');
            info.parametros.forEach(p => {
                const bloco = document.createElement('div');
                bloco.style.cssText = 'margin-bottom:10px;';
                const rot = document.createElement('label');
                rot.style.cssText = 'display:block; font-size:0.85rem; font-weight:600;';
                rot.textContent = p.rotulo + (p.tipo === 'ativo' ? ' (escolha o ativo)' : '');
                const onde = document.createElement('div');
                onde.style.cssText = 'font-size:0.7rem; color:#8b949e;';
                onde.textContent = (p.onde || []).join(' · ');
                let inp;
                if (p.tipo === 'ativo') {       // TASK-060: ativo variável = escolha entre as opções do modelo
                    inp = document.createElement('select');
                    inp.className = 'modal-input';
                    const vazio = new Option('— escolha —', '');
                    inp.append(vazio);
                    (p.opcoes || []).forEach(o => inp.append(new Option(o, o)));
                    inp.value = p.padrao || '';
                    inp.dataset.obrigatorio = p.padrao ? '' : '1';
                    inp.addEventListener('change', () => atualizarBotaoGerar());
                } else {
                    inp = document.createElement('input');
                    inp.type = 'text'; inp.inputMode = 'decimal'; inp.className = 'modal-input';
                    inp.placeholder = 'número maior que zero';
                    if (p.padrao !== null && p.padrao !== undefined) inp.value = String(p.padrao).replace('.', ',');
                }
                inp.dataset.chave = p.chave;
                bloco.append(rot, onde, inp);
                campos.appendChild(bloco);
                inputs.push(inp);
            });
            const fechar = v => { modal.classList.add('hidden'); resolve(v); };
            document.getElementById('btn-cancel-param-modelo').onclick = () => fechar(null);
            const btnGerar = document.getElementById('btn-gerar-param-modelo');
            // Gerar só habilita com todas as escolhas de ativo feitas (e sem erro do modelo)
            const atualizarBotaoGerar = () => {
                btnGerar.disabled = (info.erros || []).length > 0 || inputs.some(i => i.tagName === 'SELECT' && i.dataset.obrigatorio === '1' && !i.value);
            };
            atualizarBotaoGerar();
            const gerar = async () => {
                if (btnGerar.disabled) return;
                const valores = {};
                inputs.forEach(i => { valores[i.dataset.chave] = i.value.trim(); });
                btnGerar.disabled = true;
                try {
                    const r = await fetch(`/api/obras/${encodeURIComponent(obra.id)}/gerar?projeto=${encodeURIComponent(projetoAtualObras())}`, {
                        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ valores })
                    });
                    if (!r.ok) { erro.textContent = (await mensagensDeErro(r)).join(' '); return; }
                    const g = await r.json();
                    fechar({ cabos: g.cabos || [], outros: g.outros || [] });
                } catch (e) {
                    erro.textContent = 'Falha ao gerar a obra: ' + e.message;
                } finally {
                    atualizarBotaoGerar();
                }
            };
            btnGerar.onclick = gerar;
            inputs.forEach(i => { i.onkeydown = ev => { if (ev.key === 'Enter') gerar(); }; });
            modal.classList.remove('hidden');
            if (inputs[0]) inputs[0].focus();
        });
    }

    /** Há `V` ou `X(...)` de modelo sobrando nas tabelas? (espelha services/modelos_obra.py — Regra 4: mudou lá, mude aqui). */
    const RE_V_CABO = /^.*\S\s+V(\([^()\s]+\))?(\s+[mM])?\s*$/;
    const RE_V_OUTROS = /(^|\s)\*?(?:V(?:\([^()\s]+\))?-\S|(?:V(?:\([^()\s]+\))?|\d+(?:[.,]\d+)?)-X\((?:[^():\s,]+:)?[^()\s]*\))/;
    function contarVariaveisSobrando() {
        let n = 0;
        (tableStates.cabos.data || []).forEach(r => { if (r && RE_V_CABO.test(String(r.ativo || '').trim())) n++; });
        (tableStates.outros.data || []).forEach(r => { if (r && RE_V_OUTROS.test(String(r.ativo || '').trim())) n++; });
        return n;
    }

    let execucaoOrigemAutonomo = null;

    let revisaoObraAberta = 0;

    function restoreObraSnapshot(snap) {
        if (!snap) { showToast("Dados da obra inválidos."); return; }
        // Normaliza: o snap pode ser tableStates completo {cabos:{data,bodyId}, outros:{data,bodyId}}
        // ou um snap direto {cabos:[...], outros:[...]}
        const cabosData  = Array.isArray(snap.cabos)  ? snap.cabos  : (snap.cabos  && snap.cabos.data  ? snap.cabos.data  : null);
        const outrosData = Array.isArray(snap.outros) ? snap.outros : (snap.outros && snap.outros.data ? snap.outros.data : null);
        const totData    = Array.isArray(snap.totalizadora) ? snap.totalizadora : (snap.totalizadora && snap.totalizadora.data ? snap.totalizadora.data : null);

        if (!cabosData && !outrosData && !totData) {
            showToast("Dados da obra inválidos.");
            return;
        }

        revisaoObraAberta++;
        window.dispatchEvent(new Event('obra-resumo-substituida'));
        window.VinculacaoUI?.atualizarObra(true);
        _historicoSemSujo = true;
        try { pushHistory(); } finally { _historicoSemSujo = false; }          // abrir/carregar um projeto não o deixa "alterado"
        // TASK-059: obra gerada pelo modo autônomo -> ao salvar, o servidor compara com o que o usuário deixou (aprendizado)
        baselineManual = null;
        execucaoOrigemAutonomo = (snap.autonomo && typeof snap.autonomo.execucao_id === 'string' && !snap.autonomo.revertida) ? snap.autonomo.execucao_id : null;
        tableStates.cabos.data  = (cabosData  || []).map(deepClone);
        tableStates.outros.data = (outrosData || []).map(deepClone);
        tableStates.cabos.data.forEach(r => sincronizarHash(r, ''));            // TASK-182: ativo que chega com "#" no começo abre como comentário
        tableStates.outros.data.forEach(r => sincronizarHash(r, ''));
        
        if (totData) {
            tableStates.totalizadora.data = totData.map(r => JSON.parse(JSON.stringify(r)));
        } else {
            tableStates.totalizadora.data = [];
        }

        recalcAllQtdAtivos();
        renderTable('cabos');
        renderTable('outros');
        if (window.renderTotalizadora) window.renderTotalizadora();
        
        buildAtivoSets();
        buildDataLists();
        refreshAllFilters('cabos');
        refreshAllFilters('outros');
        updateHistoryUI();
    }


    /* ═══════════════════════════════════════
       LER PROJETO (TASK-062)
       Lê o PDF/DXF/DWG, extrai os textos e processa direto aqui, sem abrir o Leitor nem trocar de aba.
       Mesmo motor do Leitor (static/ler_projeto.js); Ctrl+Z desfaz.
    ═══════════════════════════════════════ */
    function perguntarLerProjeto(nomeArquivo, nCabos, nOutros) {
        return new Promise(resolve => {
            const fundo = document.createElement('div');
            fundo.id = 'modal-ler-projeto';
            fundo.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:10000;display:flex;align-items:center;justify-content:center;';
            const caixa = document.createElement('div');
            caixa.style.cssText = 'background:var(--bg-canvas,#0d1117);color:var(--text-primary,#e6edf3);border:1px solid var(--border-default,#30363d);border-radius:10px;padding:20px;max-width:460px;width:92%;font-family:Inter,sans-serif;';
            const titulo = document.createElement('h3');
            titulo.textContent = 'Ler projeto';
            titulo.style.margin = '0 0 10px';
            const msg = document.createElement('p');
            msg.style.cssText = 'margin:0 0 16px;font-size:.9rem;line-height:1.4;';
            msg.textContent = `As tabelas já têm conteúdo (${nCabos} cabos, ${nOutros} outros). O que fazer com "${nomeArquivo}"?`;
            const linha = document.createElement('div');
            linha.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;';
            const fim = v => { fundo.remove(); document.removeEventListener('keydown', tecla, true); resolve(v); };
            const mk = (rotulo, valor, primario) => {
                const b = document.createElement('button');
                b.className = primario ? 'btn-primary' : 'btn-secondary';
                b.textContent = rotulo;
                b.addEventListener('click', () => fim(valor));
                linha.appendChild(b);
                return b;
            };
            const tecla = e => { if (e.key === 'Escape') { e.preventDefault(); fim(null); } };
            document.addEventListener('keydown', tecla, true);
            mk('Cancelar', null, false);
            mk('Adicionar', 'adicionar', false);
            mk('Subtrair', 'subtrair', false);
            mk('Comparar', 'comparar', false);
            const bs = mk('Substituir', 'substituir', true);
            caixa.append(titulo, msg, linha);
            fundo.appendChild(caixa);
            document.body.appendChild(fundo);
            bs.focus();
        });
    }

    let lendoProjeto = false;
    async function lerProjeto(arquivo) {
        if (!arquivo || lendoProjeto) return;
        const btn = document.getElementById('btn-ler-projeto');
        const rotuloOriginal = btn ? btn.textContent : '';
        if (!window.LerProjeto.extensaoValida(arquivo.name)) { showToast('Formato não suportado. Use PDF, DXF ou DWG.'); return; }
        const projeto = localStorage.getItem('projeto_selecionado_codigo');
        if (!projeto) { showToast('Selecione um projeto antes de ler o arquivo.'); return; }

        let modo = 'substituir';
        const temConteudo = tableStates.cabos.data.length + tableStates.outros.data.length > 0;
        if (temConteudo) {
            modo = await perguntarLerProjeto(arquivo.name, tableStates.cabos.data.length, tableStates.outros.data.length);
            if (!modo) { showToast('Leitura cancelada. Nada foi alterado.'); return; }
        }

        lendoProjeto = true;
        if (btn) { btn.disabled = true; btn.textContent = 'Lendo…'; }
        try {
            await window.carregarRegrasLeitorResumo(projeto);
            const proc = window.__regrasLeitorProcessamento || [], cls = window.__regrasLeitorClassificacao || [];
            if (!proc.length && !cls.length) { showToast('Este projeto não tem regras do Leitor; nada foi alterado.'); return; }
            const itens = await window.LerProjeto.extrair(arquivo, localStorage.getItem('cloudconvert_api_key') || '',
                modo !== 'comparar' && modo !== 'subtrair');
            const r = await window.LerProjeto.processar(itens, proc, cls, RegrasLeitorEngine);
            if (modo === 'comparar' || modo === 'subtrair') {
                const obra = { cabos: r.cabos.map(deepClone), outros: r.outros.map(deepClone) };
                recalcAllQtdAtivos(obra.cabos);
                if (modo === 'comparar') compararObraNoResumo(obra, arquivo.name);
                else if (subtrairObraDoProjeto(obra)) showToast('Projeto subtraído. Ctrl+Z desfaz.');
                return;
            }

            pushHistory();  // snapshot ANTES (Ctrl+Z desfaz)
            if (modo === 'adicionar') {
                extractedDataCache = extractedDataCache.concat(r.exportados);
                tableStates.cabos.data = tableStates.cabos.data.concat(r.cabos.map(deepClone));
                tableStates.outros.data = tableStates.outros.data.concat(r.outros.map(deepClone));
                window._ramaisData = (window._ramaisData || []).concat(r.ramais);
                baselineManual = null;  // mistura de conteúdo: deixa de ser "o que o programa entregou"
            } else {
                extractedDataCache = r.exportados;
                tableStates.cabos.data = r.cabos.map(deepClone);
                tableStates.outros.data = r.outros.map(deepClone);
                tableStates.totalizadora.data = [];
                window._ramaisData = r.ramais;
                execucaoOrigemAutonomo = null;
            }
            recalcAllQtdAtivos();
            if (modo !== 'adicionar') {
                sessaoManual = (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(36).slice(2);
                baselineManual = {
                    cabos: tableStates.cabos.data.map(x => ({ entidade: x.entidade, operacao: x.operacao, ativo: x.ativo })),
                    outros: tableStates.outros.data.map(x => ({ entidade: x.entidade, operacao: x.operacao, ativo: x.ativo }))
                };
            }
            if (!MODO_JANELA_MAPA) { try { localStorage.setItem('processar_dados', JSON.stringify(extractedDataCache)); } catch (e) { /* sem espaço: a tela segue valendo, só não sobrevive ao F5 */ } }
            renderTable('cabos');
            renderTable('outros');
            if (window.renderTotalizadora) window.renderTotalizadora();
            buildAtivoSets();
            buildDataLists();
            refreshAllFilters('cabos');
            refreshAllFilters('outros');
            updateCounters();
            updateHistoryUI();
            showToast(`Projeto lido: ${r.cabos.length} cabos, ${r.outros.length} outros${r.ramais.length ? `, ${r.ramais.length} ramais` : ''}. Ctrl+Z desfaz.`);
        } catch (e) {
            showToast(`${e.message || e}`.slice(0, 160));
        } finally {
            lendoProjeto = false;
            if (btn) { btn.disabled = false; btn.textContent = rotuloOriginal; }
        }
    }
    window.lerProjeto = lerProjeto;

    (function ligarLerProjeto() {
        const btn = document.getElementById('btn-ler-projeto'), inp = document.getElementById('input-ler-projeto');
        if (!btn || !inp) return;
        btn.addEventListener('click', () => { inp.value = ''; inp.click(); });
        inp.addEventListener('change', () => { const f = inp.files && inp.files[0]; inp.value = ''; if (f) lerProjeto(f); });
        ['dragenter', 'dragover'].forEach(ev => btn.addEventListener(ev, e => { e.preventDefault(); btn.classList.add('arrastando'); }));
        ['dragleave', 'drop'].forEach(ev => btn.addEventListener(ev, e => { e.preventDefault(); btn.classList.remove('arrastando'); }));
        btn.addEventListener('drop', e => { const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; if (f) lerProjeto(f); });
    })();

    const aiChatMessages = document.getElementById('chat-messages');
    const aiChatInput = document.getElementById('chat-input');
    const btnSendChat = document.getElementById('btn-send-chat');
    
    // API Key config — três provedores (TASK-055): Gemini, OpenAI e Claude. Campos, headers e
    // chaves de localStorage de cada um, usados pelas funções abaixo em vez de ids fixos.
    const modalApikey = document.getElementById('modal-apikey');
    const inputApikey = document.getElementById('input-apikey');       // listeners de blur/input do Gemini
    const PROVIDER_CAMPOS = {
        gemini: {
            apikeyId: 'input-apikey', modelId: 'input-model', statusId: 'apikey-status',
            keyStorage: 'gemini_api_key', modelStorage: 'gemini_model',
            keyHeader: 'X-Gemini-Key', modelHeader: 'X-Gemini-Model',
        },
        openai: {
            apikeyId: 'input-apikey-openai', modelId: 'input-model-openai', statusId: 'apikey-status-openai',
            keyStorage: 'openai_api_key', modelStorage: 'openai_model',
            keyHeader: 'X-OpenAI-Key', modelHeader: 'X-OpenAI-Model',
        },
        claude: {
            apikeyId: 'input-apikey-claude', modelId: 'input-model-claude', statusId: 'apikey-status-claude',
            keyStorage: 'anthropic_api_key', modelStorage: 'anthropic_model',
            keyHeader: 'X-Anthropic-Key', modelHeader: 'X-Anthropic-Model',
        },
    };

    function providerAtivo() {
        const btn = document.querySelector('.btn-provider.active');
        return (btn && btn.dataset.provider) || localStorage.getItem('ai_provider') || 'gemini';
    }

    async function fetchModelsPara(provider) {
        const campos = PROVIDER_CAMPOS[provider];
        const inputKeyEl = document.getElementById(campos.apikeyId);
        const inputModelEl = document.getElementById(campos.modelId);
        if (!inputKeyEl || !inputModelEl || provider === 'claude') return; // Claude: lista fixa já no HTML
        const keyToUse = inputKeyEl.value.trim() || "SAVED_IN_BACKEND";
        try {
            const resp = await fetch('/api/gemini/models', {
                headers: { [campos.keyHeader]: keyToUse, 'X-Provider': provider }
            });
            if (resp.ok) {
                const data = await resp.json();
                inputModelEl.innerHTML = '<option value="">Automático</option>';
                (data.models || []).forEach(m => {
                    const opt = document.createElement('option');
                    opt.value = m.id || m;
                    opt.textContent = m.label || m.id || m;
                    inputModelEl.appendChild(opt);
                });
                inputModelEl.value = localStorage.getItem(campos.modelStorage) || '';
            }
        } catch (e) {
            console.error("Erro ao carregar modelos", e);
        }
    }
    const fetchModels = () => fetchModelsPara('gemini');   // compatibilidade com o listener do Gemini

    inputApikey.addEventListener('blur', fetchModels);
    const inputApikeyOpenai = document.getElementById('input-apikey-openai');
    if (inputApikeyOpenai) inputApikeyOpenai.addEventListener('blur', () => fetchModelsPara('openai'));

    // Informa de onde virá a chave quando o usuário não digita nenhuma (nunca mostra o valor).
    // O diagnóstico de origem da chave (/api/health) hoje só existe para o Gemini; para os demais
    // provedores mostramos só se uma chave foi digitada nesta tela (sem distinguir salva/padrão).
    async function atualizarStatusChavePara(provider) {
        const campos = PROVIDER_CAMPOS[provider];
        const el = document.getElementById(campos.statusId);
        const inputKeyEl = document.getElementById(campos.apikeyId);
        if (!el || !inputKeyEl) return;
        if (inputKeyEl.value.trim()) { el.textContent = 'Sua chave será usada e salva.'; return; }
        if (provider !== 'gemini') {
            el.textContent = localStorage.getItem(campos.keyStorage)
                ? 'Sem chave digitada: será usada a chave salva neste navegador.'
                : 'Nenhuma chave disponível: informe a sua para usar a IA.';
            return;
        }
        const textos = {
            salva: 'Sem chave digitada: será usada a chave salva neste servidor.',
            padrao: 'Sem chave digitada: será usada a chave padrão do sistema (com limite de mensagens por minuto).',
            nenhuma: 'Nenhuma chave disponível: informe a sua para usar a IA.'
        };
        try {
            const resp = await fetch('/api/health');
            const info = resp.ok ? await resp.json() : {};
            el.textContent = textos[info.ai_key_source] || '';
        } catch (e) {
            el.textContent = '';
        }
    }
    const atualizarStatusChave = () => atualizarStatusChavePara('gemini');
    inputApikey.addEventListener('input', atualizarStatusChave);
    if (inputApikeyOpenai) inputApikeyOpenai.addEventListener('input', () => atualizarStatusChavePara('openai'));
    const inputApikeyClaude = document.getElementById('input-apikey-claude');
    if (inputApikeyClaude) inputApikeyClaude.addEventListener('input', () => atualizarStatusChavePara('claude'));

    // Troca de provedor no modal: mostra a seção certa e marca o botão ativo (TASK-055).
    function selecionarProvider(provider) {
        if (!PROVIDER_CAMPOS[provider]) provider = 'gemini';
        document.querySelectorAll('.btn-provider').forEach(b => b.classList.toggle('active', b.dataset.provider === provider));
        Object.keys(PROVIDER_CAMPOS).forEach(p => {
            const secao = document.getElementById(`section-${p}`);
            if (secao) secao.classList.toggle('hidden', p !== provider);
        });
        if (provider === 'claude') {
            atualizarStatusChavePara('claude');
        } else {
            fetchModelsPara(provider);
            atualizarStatusChavePara(provider);
        }
    }
    document.querySelectorAll('.btn-provider').forEach(btn => {
        btn.addEventListener('click', () => selecionarProvider(btn.dataset.provider));
    });

    const btnConfigApi = document.getElementById('btn-config-api');
    if (btnConfigApi) {
        btnConfigApi.addEventListener('click', () => {
            Object.entries(PROVIDER_CAMPOS).forEach(([provider, campos]) => {
                const keyEl = document.getElementById(campos.apikeyId);
                const modelEl = document.getElementById(campos.modelId);
                if (keyEl) keyEl.value = localStorage.getItem(campos.keyStorage) || '';
                // Claude tem modelo padrão pré-selecionado (Haiku) quando nada foi salvo ainda.
                if (modelEl) modelEl.value = localStorage.getItem(campos.modelStorage) || (provider === 'claude' ? 'claude-haiku-5-5' : '');
            });
            selecionarProvider(providerAtivo());
            modalApikey.classList.remove('hidden');
        });
    }

    const btnSaveApiKey = document.getElementById('btn-save-apikey');
    if (btnSaveApiKey) {
        btnSaveApiKey.addEventListener('click', () => {
            const provider = providerAtivo();
            const campos = PROVIDER_CAMPOS[provider];
            // Bug corrigido (TASK-055): salvava 'ai_provider' sempre como 'gemini', ignorando o
            // provedor escolhido no seletor.
            localStorage.setItem('ai_provider', provider);
            const keyEl = document.getElementById(campos.apikeyId);
            const modelEl = document.getElementById(campos.modelId);
            if (keyEl) localStorage.setItem(campos.keyStorage, keyEl.value.trim());
            if (modelEl) localStorage.setItem(campos.modelStorage, modelEl.value.trim());
            modalApikey.classList.add('hidden');
            showToast('Configurações salvas!');
        });
    }

    // Chat logic
    let chatHistory = [];
    let projetoHistoricoChat;
    
    function addChatMessage(text, sender) {
        const div = document.createElement('div');
        div.className = `chat-message ${sender}`;
        div.textContent = text;
        div.style.whiteSpace = 'pre-wrap';
        aiChatMessages.appendChild(div);
        aiChatMessages.scrollTop = aiChatMessages.scrollHeight;
        return div;
    }

    function extractTableFromMarkdown(markdown) {
        const lines = markdown.split('\n');
        let inTable = false;
        let result = [];
        for (let line of lines) {
            line = line.trim();
            if (line.startsWith('|') && line.includes('ATIVOS')) {
                // Serve tanto para o formato antigo (AÇÃO|ATIVOS) quanto para o novo (COMANDO|ID|AÇÃO|ATIVOS)
                inTable = true;
                continue;
            }
            if (inTable && line.startsWith('|') && line.includes('---')) continue;
            
            if (inTable && line.startsWith('|')) {
                const parts = line.split('|').slice(1, -1).map(s => s.trim());
                if (parts.length >= 4) {
                    result.push({ comando: parts[0], idStr: parts[1], operacao: parts[2], ativosStr: parts[3] });
                } else if (parts.length >= 2) {
                    result.push({ comando: 'ADICIONAR', idStr: '-', operacao: parts[parts.length-2], ativosStr: parts[parts.length-1] });
                }
            } else if (inTable) {
                // Fim da tabela
                break;
            }
        }
        return result;
    }

    /** Pergunta ao usuário antes de aplicar um comando de obra da IA. Resolve true/false. */
    function confirmarAcaoObra(acao, obra, nCabos, nOutros) {
        return new Promise(resolve => {
            const modal = document.getElementById('modal-obra-ia');
            const titulo = { adicionar: 'Adicionar obra às tabelas atuais', subtrair: 'Subtrair obra das tabelas atuais', carregar: 'Carregar obra (substitui as tabelas atuais)' }[acao];
            document.getElementById('obra-ia-titulo').textContent = titulo;
            document.getElementById('obra-ia-nome').textContent = obra.nome;
            document.getElementById('obra-ia-detalhe').textContent = `${obra.data || ''} · ${nCabos} linha(s) em Cabos, ${nOutros} em Outros`;
            const efeito = {
                adicionar: 'As linhas da obra serão acrescentadas ao fim das tabelas atuais.',
                subtrair: 'As linhas da obra serão acrescentadas como retirada (sinal invertido).',
                carregar: 'ATENÇÃO: as tabelas atuais serão SUBSTITUÍDAS pelo conteúdo da obra.'
            }[acao];
            const el = document.getElementById('obra-ia-efeito');
            el.textContent = `${efeito} Nada muda até você confirmar; Ctrl+Z desfaz.`;
            el.style.color = acao === 'carregar' ? '#f85149' : '#8b949e';
            const acoes = document.getElementById('obra-ia-acoes');
            acoes.replaceChildren();
            const fechar = v => { modal.classList.add('hidden'); resolve(v); };
            const cancelar = botaoPainel('Cancelar', 'btn-secondary');
            cancelar.onclick = () => fechar(false);
            const ok = botaoPainel({ adicionar: 'Adicionar', subtrair: 'Subtrair', carregar: 'Carregar e substituir' }[acao], 'btn-primary');
            ok.onclick = () => fechar(true);
            acoes.append(cancelar, ok);
            modal.classList.remove('hidden');
        });
    }

    /** Comando de obra pedido pela IA ({acao_ui:"obra", acao, obra_id}). O id é validado contra as obras reais do usuário
     *  (a IA não inventa obra), a obra precisa ser do projeto selecionado e o usuário sempre confirma. Devolve {msg}. */
    async function executarAcaoObra(action) {
        const rotulo = { adicionar: 'adicionada', subtrair: 'subtraída', carregar: 'carregada' };
        if (!rotulo[action.acao] || typeof action.obra_id !== 'string' || !action.obra_id) return { msg: 'Comando de obra inválido; nada foi alterado.' };
        let obra;
        try {
            const projetoAtual = localStorage.getItem('projeto_selecionado_codigo') || '229';
            const r = await fetch(`/api/obras/${encodeURIComponent(action.obra_id)}?projeto=${encodeURIComponent(projetoAtual)}`);
            if (r.status === 404) return { msg: 'Não encontrei essa obra entre as suas obras ou as públicas do projeto; nada foi alterado.' };
            if (!r.ok) throw new Error(`HTTP ${r.status}`);
            obra = await r.json();
        } catch (e) {
            return { msg: `Não foi possível consultar a obra (${e.message}); nada foi alterado.` };
        }
        if (obra.projeto !== (localStorage.getItem('projeto_selecionado_codigo') || '229')) {
            return { msg: 'Essa obra é de outro projeto. Troque o projeto selecionado para usá-la; nada foi alterado.' };
        }
        let snap;
        try { snap = typeof obra.dados_json === 'string' ? JSON.parse(obra.dados_json) : obra.dados_json; } catch (e) { snap = null; }
        if (!snap || typeof snap !== 'object') return { msg: 'Os dados dessa obra estão ilegíveis; nada foi alterado.' };
        if (obra.tipo === 'modelo') {       // TASK-058: a IA nunca preenche valores; o usuário os digita na janela
            snap = await pedirParametrosModelo(obra);
            if (!snap) return { msg: 'Ação cancelada. Nada foi alterado.' };
        }
        const nCabos = linhasDoSnap(snap, 'cabos').length, nOutros = linhasDoSnap(snap, 'outros').length;
        if (!(await confirmarAcaoObra(action.acao, obra, nCabos, nOutros))) return { msg: 'Ação cancelada. Nada foi alterado.' };
        if (action.acao === 'adicionar') adicionarObraAoProjeto(snap);
        else if (action.acao === 'subtrair') subtrairObraDoProjeto(snap);
        else restoreObraSnapshot(snap);
        return { msg: `Obra "${obra.nome}" ${rotulo[action.acao]}. Ctrl+Z desfaz.` };
    }

    function executeUiAction(action) {
        const alvos = action.alvo === 'ambos' ? ['cabos', 'outros'] : [action.alvo];
        
        alvos.forEach(alvo => {
            if (!tableStates[alvo]) return;

            if (action.acao_ui === 'ordenar') {
                if (!['entidade', 'operacao', 'ativo', 'qtdAtivos'].includes(action.coluna)) { showToast('Coluna não ordenável: ' + action.coluna); return; }
                ordenarTabela(alvo, [{ campo: action.coluna, dir: action.ordem === 'desc' ? 'desc' : 'asc' }]);
            } 
            else if (action.acao_ui === 'filtrar') {
                if (!filterSelections[alvo][action.coluna]) { showToast('Coluna não filtrável: ' + action.coluna); return; }
                filterSelections[alvo][action.coluna].clear();
                filterSelections[alvo][action.coluna].add(action.valor);
                applyFilters(alvo);
                
                // Atualiza visualmente o dropdown
                const containerId = FILTER_CONTAINERS[alvo][action.coluna];
                const container = document.getElementById(containerId);
                if (container) {
                    const trigger = container.querySelector('.r-filter-trigger span');
                    const ind = container.querySelector('.r-filter-indicator');
                    if (trigger) trigger.textContent = '1 sel';
                    if (ind) ind.classList.add('active');
                    
                    const inner = container.querySelector('.r-options-inner');
                    if (inner) {
                        inner.querySelectorAll('input').forEach(cb => {
                            cb.checked = cb.parentElement.dataset.value === action.valor;
                        });
                    }
                    const mainChk = container.querySelector('.r-select-all-opt input');
                    if (mainChk) mainChk.checked = false;
                }
                showToast(`Filtro aplicado na tabela ${alvo}!`);
            }
            else if (action.acao_ui === 'limpar_filtros') {
                camposFiltro(alvo).forEach(field => filterSelections[alvo][field].clear());
                applyFilters(alvo);
                refreshAllFilters(alvo);
                showToast(`Filtros limpos na tabela ${alvo}!`);
            }
        });
    }

    // ── Plano do assistente (TASK-191) ───────────────────────────────────────────────────────────────
    const projetoAssistente = () => localStorage.getItem('projeto_selecionado_codigo') || '229';
    const parametrosAssistente = () => window.ObraParams ? ObraParams.obraDoProjeto() : null;

    const tabelasAssistente = () => JSON.stringify({ cabos: tableStates.cabos.data, outros: tableStates.outros.data });

    function revisaoAssistente() {
        return {
            projeto: projetoAssistente(),
            obra: JSON.stringify(parametrosAssistente()),
            caminho: window.GridRProjeto?.estado()?.caminho,
            tabelas: tabelasAssistente(),
            referencias: { cabos: tableStates.cabos.data.slice(), outros: tableStates.outros.data.slice() }
        };
    }

    function revisaoAindaAtiva(r) {
        return r.projeto === projetoAssistente() && r.obra === JSON.stringify(parametrosAssistente()) &&
            r.caminho === window.GridRProjeto?.estado()?.caminho &&
            r.tabelas === tabelasAssistente() && ['cabos', 'outros'].every(t =>
                r.referencias[t].every((linha, i) => linha === tableStates[t].data[i]));
    }

    let selecaoMapaAssistente = null;
    function selecionarMapaAssistente(selecao) {
        selecaoMapaAssistente = { selecao, revisao: revisaoAssistente() };
        const limpar = document.getElementById('assistente-selecao-mapa');
        limpar.textContent = selecao.rotulo + ' ×';
        limpar.hidden = false;
        window.ResumoLayout.definirPainel('ia', true);
        if (!aiChatInput.value) aiChatInput.value = 'Como está a vinculação neste ' + selecao.tipo + '?';
        aiChatInput.focus();
    }

    function limparSelecaoMapaAssistente() {
        selecaoMapaAssistente = null;
        document.getElementById('assistente-selecao-mapa').hidden = true;
    }
    document.getElementById('assistente-selecao-mapa').addEventListener('click', limparSelecaoMapaAssistente);

    function simularPlanoAssistente(itens, revisao) {
        if (!revisaoAindaAtiva(revisao)) throw new Error('A obra mudou desde o pedido. Gere um novo plano.');
        const tabelas = JSON.parse(revisao.tabelas);
        const estados = { cabos: { data: tabelas.cabos }, outros: { data: tabelas.outros } };
        const outrosAntes = estados.outros.data.slice();
        const cabosAntes = new Map(estados.cabos.data.map((c, i, lista) => [c, {
            ativo: c?.ativo, proximo: lista[i + 1]
        }]));
        const VR = window.VinculacaoRegras;
        const edicoes = new Map(itens.filter(i => i.tabela === 'outros' && i.acao === 'editar').map(i => [i.id, i.ativo]));
        for (const cabo of estados.cabos.data) {
            for (const id of cabo?.vinculoEstruturas || []) {
                const linha = VR.linhaDe(id), slot = VR.slotDe(id);
                if (!edicoes.has(linha)) continue;
                if (_codigosEstruturaVinculo(outrosAntes[linha]?.ativo || '')[slot] !== _codigosEstruturaVinculo(edicoes.get(linha))[slot]) {
                    throw new Error('A edição trocaria uma estrutura já vinculada. Revise os vínculos no mapa.');
                }
            }
        }
        const resultado = Assistente.aplicarPlano(estados, itens, (ativo, tabela, linha) => {
            if (linha) {
                if (tabela === 'outros') reclassificarEntidade(linha);
                return linha.entidade;
            }
            return tabela === 'cabos' ? 'CABO' : autoClassifyEntidade(ativo);
        });
        if (resultado.ignoradas.length) throw new Error(Assistente.resumoAplicacao({ ...resultado, adicionadas: 0, editadas: 0, removidas: 0 }));
        const GL = window.GruposLinhas;
        for (const t of ['cabos', 'outros']) estados[t].data = GL.normalizar(estados[t].data);
        _remapearVinculosOutros(GL.mapa(outrosAntes, estados.outros.data), estados.cabos.data);
        const cabos = estados.cabos.data;
        for (let i = cabos.length - 1; i >= 0; i--) {
            const antes = cabosAntes.get(cabos[i]);
            if (!antes || antes.ativo !== cabos[i].ativo) recalcRowAndAbove(i, cabos);
            else if (antes.proximo !== cabos[i + 1] && isStandaloneLine(cabos[i].ativo)) {
                recalcRowAndAbove(i, cabos);
            }
        }
        return { estados, resultado };
    }

    async function previaPlanoAssistente(itens, revisao, req) {
        if (!document.querySelector('input[name="view_mode"][value="padrao"]')?.checked) {
            throw new Error('Mude para Cabos e Demais ativos para simular as alterações das tabelas.');
        }
        await prepararTotalizadora();
        const { estados } = simularPlanoAssistente(itens, revisao);
        const antes = payloadDaTotalizadora(construirTotalizadora(tableStates.cabos.data, tableStates.outros.data));
        const depois = payloadDaTotalizadora(construirTotalizadora(estados.cabos.data, estados.outros.data));
        const r = await fetch('/api/assistente/previa', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ projeto_codigo: req.projeto_codigo, obra: req.obra, antes, depois })
        });
        const dados = await r.json();
        if (!r.ok) throw new Error(dados.detail || 'Não foi possível simular o orçamento.');
        if (!revisaoAindaAtiva(revisao)) throw new Error('A obra mudou durante a simulação. Gere um novo plano.');
        return dados;
    }

    function mostrarPlanoAssistente(plano, revisao, req) {
        let aplicado;
        let posHistorico;
        const card = Assistente.renderPlano(plano, {
            aplicar(itens) {
                const { estados, resultado } = simularPlanoAssistente(itens, revisao);
                if (!resultado.adicionadas && !resultado.editadas && !resultado.removidas) {
                    return { mensagem: Assistente.resumoAplicacao(resultado), aplicado: false };
                }
                pushHistory();
                limparGrade();
                for (const t of ['cabos', 'outros']) tableStates[t].data = estados[t].data;
                renderTable('cabos');
                renderTable('outros');
                buildAtivoSets();
                buildDataLists();
                refreshAllFilters('cabos');
                refreshAllFilters('outros');
                aplicado = revisaoAssistente();
                posHistorico = historyIdx;
                return { mensagem: Assistente.resumoAplicacao(resultado) + '. Ctrl+Z desfaz.', aplicado: true };
            },
            previa: itens => previaPlanoAssistente(itens, revisao, req),
            desfazer() {
                if (historyIdx !== posHistorico || !revisaoAindaAtiva(aplicado)) return 'Houve outra alteração. Use o histórico do Resumo.';
                undo();
                return 'Desfeito.';
            },
            async obra(item) {
                if (!revisaoAindaAtiva(revisao)) throw new Error('A obra mudou desde o pedido. Gere um novo plano.');
                return (await executarAcaoObra({ acao_ui: 'obra', acao: item.modo, obra_id: item.obra_id })).msg;
            }
        });
        aiChatMessages.appendChild(card);
        aiChatMessages.scrollTop = aiChatMessages.scrollHeight;
    }

    // Provedor reserva automático (dentro da «Configuração de IA»): recarrega o estado sempre que a janela abre
    const caixaReserva = document.getElementById('assistente-reserva');
    if (caixaReserva && modalApikey) {
        const reservaUI = Assistente.reservaUI(caixaReserva, { toast: showToast });
        new MutationObserver(() => { if (!modalApikey.classList.contains('hidden')) reservaUI.recarregar(); }).observe(modalApikey, { attributes: true, attributeFilter: ['class'] });
    }

    const btnAssistenteSkills = document.getElementById('btn-assistente-skills');
    if (btnAssistenteSkills) btnAssistenteSkills.addEventListener('click', () => Assistente.abrirSkills({
        projeto: () => localStorage.getItem('projeto_selecionado_codigo') || '229', toast: showToast
    }));

    let chatAbort = null;                                     // enquanto a IA responde, o botão de enviar vira «Parar» (aborta o pedido)
    const iconeEnviar = btnSendChat.innerHTML;
    const iconeParar = '<svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><rect x="6" y="6" width="12" height="12" rx="2"></rect></svg>';
    function modoParar(ligado) {
        btnSendChat.innerHTML = ligado ? iconeParar : iconeEnviar;
        btnSendChat.title = ligado ? 'Parar (interrompe a resposta)' : 'Enviar Mensagem (Enter)';
        btnSendChat.disabled = false;
    }

    async function sendToGemini() {
        const text = aiChatInput.value.trim();
        if (!text) return;

        // Provedor de IA selecionado (TASK-055): chave/modelo/headers são os do provedor salvo,
        // não mais fixos no Gemini.
        const providerChat = localStorage.getItem('ai_provider') || 'gemini';
        const camposChat = PROVIDER_CAMPOS[providerChat] || PROVIDER_CAMPOS.gemini;
        const apiKey = localStorage.getItem(camposChat.keyStorage) || 'SAVED_IN_BACKEND';
        const customModel = localStorage.getItem(camposChat.modelStorage) || 'SAVED_IN_BACKEND';

        addChatMessage(text, 'user');
        aiChatInput.value = '';
        aiChatInput.disabled = true;
        chatAbort = new AbortController();
        modoParar(true);

        // Uma instrução digitada no chat pode orientar a conversa. Publicação durável passa pelo editor de skills.
        if (/^(lembre-se\b|regra\s*:)/i.test(text)) {
            const projeto = projetoAssistente();
            const aviso = addChatMessage('Orientações publicadas exigem revisão e salvamento por um administrador.', 'system');
            const botao = document.createElement('button');
            botao.className = 'btn-secondary';
            botao.type = 'button';
            botao.textContent = 'Revisar como skill';
            botao.addEventListener('click', () => Assistente.abrirSkills({ projeto: () => projeto, toast: showToast, rascunho: text }));
            aviso.append(' ', botao);
        }

        // Indicador carregando
        const loadingDiv = document.createElement('div');
        loadingDiv.className = 'chat-message ai';
        const statusSpan = document.createElement('span');
        statusSpan.className = 'loading-dots';
        loadingDiv.appendChild(statusSpan);
        aiChatMessages.appendChild(loadingDiv);
        aiChatMessages.scrollTop = aiChatMessages.scrollHeight;
        // Andamento visível: o que está sendo feito + há quantos segundos (a IA pode demorar; o usuário vê que não travou e pode parar)
        let rotuloAtual = 'Pensando';
        const t0 = Date.now();
        const desenharStatus = () => {
            const seg = Math.floor((Date.now() - t0) / 1000);
            statusSpan.textContent = rotuloAtual + (seg >= 2 ? ` · ${seg} s` : '') + (seg >= 15 ? ' · demorando… (■ para parar)' : '');
        };
        desenharStatus();
        const relogio = setInterval(desenharStatus, 1000);

        try {
            if (selecaoMapaAssistente && !revisaoAindaAtiva(selecaoMapaAssistente.revisao)) {
                limparSelecaoMapaAssistente();
                addChatMessage('A seleção do mapa foi limpa porque a obra mudou.', 'system');
            }
            const revisao = revisaoAssistente();
            if (projetoHistoricoChat !== revisao.projeto) {
                chatHistory = [];
                projetoHistoricoChat = revisao.projeto;
            }
            const tabelas = Assistente.tabelaParaServidor(tableStates);      // TASK-191: o servidor lê a tabela sob demanda (ferramenta ler_tabela) e nunca a altera
            const reqBody = {
                prompt: text,
                history: chatHistory,
                provider: providerChat,
                openai_base_url: providerChat === 'openai' ? (localStorage.getItem('openai_base_url') || '') : '',
                projeto_codigo: revisao.projeto,
                selecao_mapa: selecaoMapaAssistente?.selecao,
                tabela_incompleta: tabelas.tabela_incompleta || false,
                cabos: tabelas.cabos,
                outros: tabelas.outros,
                obra: window.ObraParams ? ObraParams.obraDoProjeto() : null
            };
            const response = await fetch('/api/assistente/chat', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    [camposChat.keyHeader]: apiKey,
                    [camposChat.modelHeader]: customModel
                },
                body: JSON.stringify(reqBody),
                signal: chatAbort.signal
            });

            if (!response.ok) {
                if (loadingDiv.parentNode) aiChatMessages.removeChild(loadingDiv);
                if (response.status === 401) {
                    modalApikey.classList.remove('hidden');
                    addChatMessage("Chave da API não encontrada. Por favor, insira e tente novamente.", 'error');
                } else {
                    const err = await response.json().catch(() => ({}));
                    addChatMessage(`Erro: ${err.detail || 'Falha na comunicação'}`, 'error');
                }
                aiChatInput.disabled = false;
                btnSendChat.disabled = false;
                return;
            }

            // Stream NDJSON: ferramenta (progresso) · texto · plano · acao_ui · erro · fim
            const reader = response.body.getReader();
            const decoder = new TextDecoder();
            let buffer = '', replyText = '', erroIA = '', div = null;
            const planos = [], acoesUi = [];
            const rotuloFerramenta = { buscar_ativo: 'Consultando a base', detalhar_ativo: 'Lendo o ativo', ler_tabela: 'Lendo a tabela', validar_linhas: 'Conferindo as linhas',
                propor_alteracoes: 'Montando o plano', totalizar: 'Totalizando', calcular_orcamento: 'Calculando o orçamento', carregar_skill: 'Consultando uma skill',
                listar_skills: 'Consultando as skills', listar_obras: 'Vendo as obras', consultar_obra: 'Lendo a obra', propor_obra: 'Preparando a obra', ajustar_tela: 'Ajustando a tela' };
            const tratar = ev => {
                if (ev.tipo === 'ferramenta') {
                    rotuloAtual = (rotuloFerramenta[ev.nome] || ev.nome) + (ev.detalhe ? ' · ' + ev.detalhe : '');
                    desenharStatus();
                } else if (ev.tipo === 'texto') {
                    replyText += (replyText ? '\n\n' : '') + ev.texto;
                    if (!div) {
                        div = document.createElement('div');
                        div.className = 'chat-message ai';
                        div.style.whiteSpace = 'pre-wrap';
                        aiChatMessages.insertBefore(div, loadingDiv.parentNode ? loadingDiv : null);
                    }
                    div.textContent = replyText;
                } else if (ev.tipo === 'aviso') {                       // o assistente trocou para o provedor reserva
                    const av = document.createElement('div');
                    av.className = 'chat-message system';
                    av.textContent = ev.texto;
                    aiChatMessages.insertBefore(av, loadingDiv.parentNode ? loadingDiv : null);
                } else if (ev.tipo === 'plano') planos.push(ev);
                else if (ev.tipo === 'acao_ui') acoesUi.push(ev);
                else if (ev.tipo === 'erro') erroIA = ev.mensagem || 'Falha na IA.';
            };
            while (true) {
                const { done, value } = await reader.read();
                if (done) break;
                buffer += decoder.decode(value, { stream: true });
                const lido = Assistente.linhasNdjson(buffer);
                buffer = lido.resto;
                lido.eventos.forEach(tratar);
                aiChatMessages.scrollTop = aiChatMessages.scrollHeight;
            }
            Assistente.linhasNdjson(buffer + decoder.decode() + '\n').eventos.forEach(tratar);
            if (loadingDiv.parentNode) aiChatMessages.removeChild(loadingDiv);

            if (erroIA) addChatMessage('Erro: ' + erroIA, 'error');
            else if (!replyText.trim() && !planos.length && !acoesUi.length) addChatMessage('A IA não retornou resposta. Verifique a API Key e o modelo selecionado.', 'error');
            if (erroIA && !replyText.trim() && !planos.length) {
                aiChatInput.disabled = false;
                btnSendChat.disabled = false;
                return;
            }

            if (revisaoAindaAtiva(revisao)) acoesUi.forEach(a => executeUiAction(a));
            chatHistory.push({ role: 'user', parts: [{ text: text }] });
            chatHistory.push({ role: 'model', parts: [{ text: replyText.trim() || (planos.length ? '*(Plano proposto)*' : '*(Ação de interface executada)*') }] });
            planos.forEach(p => mostrarPlanoAssistente(p, revisao, reqBody));
        } catch (error) {
            if (loadingDiv.parentNode) aiChatMessages.removeChild(loadingDiv);
            if (error && error.name === 'AbortError') addChatMessage('Interrompido por você.', 'system');
            else addChatMessage(`Erro de conexão: ${error.message}`, 'error');
        } finally {
            clearInterval(relogio);
            chatAbort = null;
            modoParar(false);
            aiChatInput.disabled = false;
            aiChatInput.focus();
        }
    }

    btnSendChat.addEventListener('click', () => { if (chatAbort) chatAbort.abort(); else sendToGemini(); });
    aiChatInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            sendToGemini();
        }
    });

    // --- MEMÓRIA (Salvar / Carregar) ---
    const modalSaveObra = document.getElementById('modal-save-obra');
    const modalLoadObra = document.getElementById('modal-load-obra');
    
    document.getElementById('btn-save-obra').addEventListener('click', () => {
        document.getElementById('input-obra-nome').value = '';
        document.getElementById('obra-vis-particular').checked = true;
        modalSaveObra.classList.remove('hidden');
    });

    document.getElementById('btn-confirm-save-obra').addEventListener('click', async () => {
        const nome = document.getElementById('input-obra-nome').value.trim();
        if (!nome) return alert('Digite um nome');
        const publica = document.getElementById('obra-vis-publica').checked;
        const projetoSalvar = localStorage.getItem('projeto_selecionado_codigo') || '229';
        if (publica && !confirm(`Todos os usuários do projeto ${projetoSalvar} poderão ver e copiar esta obra. Deseja torná-la pública?`)) return;

        const obra = {
            id: 'obra_' + Date.now(),
            nome: nome,
            data: new Date().toLocaleString(),
            dados_json: JSON.stringify(tableStates),
            projeto: projetoSalvar,
            publica: publica
        };
        if (execucaoOrigemAutonomo) obra.origem_execucao = execucaoOrigemAutonomo;
        else if (baselineManual && localStorage.getItem('is_admin') === 'true') obra.baseline_manual = { sessao: sessaoManual, ...baselineManual };

        try {
            const resp = await fetch('/api/obras', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(obra)
            });
            if (!resp.ok) throw new Error(await lerDetalhe(resp));
            modalSaveObra.classList.add('hidden');
            showToast(publica ? 'Obra salva como pública.' : 'Obra salva com sucesso!');
        } catch (e) {
            showToast('Erro ao salvar: ' + (e.message || 'backend indisponível'));
        }
    });

    // ── Salvar como modelo (TASK-058) ──
    let modeloEditando = null;      // {id, nome, parametros} quando o criador abriu um modelo para editar
    function atualizarAvisoModelo() {
        const el = document.getElementById('modelo-editando-aviso');
        el.style.display = modeloEditando ? '' : 'none';
        el.textContent = modeloEditando ? `Editando o modelo "${modeloEditando.nome}" — "Salvar como modelo" o sobrescreve.` : '';
    }
    async function definirModeloEditando(m) {
        modeloEditando = m;
        atualizarAvisoModelo();
        try {   // rótulos/padrões já configurados, para não se perderem na edição
            const r = await fetch(`/api/obras/${encodeURIComponent(m.id)}/modelo?projeto=${encodeURIComponent(projetoAtualObras())}`);
            if (r.ok && modeloEditando && modeloEditando.id === m.id) modeloEditando.parametros = (await r.json()).parametros || [];
        } catch (e) { /* segue sem os rótulos antigos */ }
    }

    const modalSaveModelo = document.getElementById('modal-save-modelo');
    document.getElementById('btn-cancel-save-modelo').addEventListener('click', () => modalSaveModelo.classList.add('hidden'));
    document.getElementById('btn-save-modelo').addEventListener('click', async () => {
        const dadosModelo = () => ({
            cabos: (tableStates.cabos.data || []).filter(Boolean).map(r => ({ entidade: r.entidade, operacao: r.operacao, ativo: r.ativo, qtdAtivos: r.qtdAtivos, texto: r.texto })),
            outros: (tableStates.outros.data || []).filter(Boolean).map(r => ({ entidade: r.entidade, operacao: r.operacao, ativo: r.ativo, qtdAtivos: r.qtdAtivos, texto: r.texto }))
        });
        const dados = dadosModelo();
        let params, errosModelo = [], avisosModelo = [];
        try {
            const r = await fetch('/api/obras/modelo/detectar', { method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ ...dados, projeto: projetoAtualObras() }) });
            if (!r.ok) throw new Error((await mensagensDeErro(r)).join(' '));
            const det = await r.json();
            params = det.parametros; errosModelo = det.erros || []; avisosModelo = det.avisos || [];
        } catch (e) { showToast('Erro ao procurar variáveis: ' + e.message); return; }
        if (!params.length) {
            alert('Nenhuma variável encontrada. Escreva V no lugar de uma quantidade — Cabos: "CAA 2 ABC V m"; Outros: "V-U4" ou "*V-CFU" — ou X(A,B,C) no lugar do ativo (o usuário escolherá), e tente de novo.');
            return;
        }
        document.getElementById('input-modelo-nome').value = modeloEditando ? modeloEditando.nome : '';
        document.getElementById('modelo-save-erro').textContent = errosModelo.join(' ');
        document.getElementById('modelo-save-avisos').textContent = avisosModelo.join(' ');
        document.getElementById('btn-confirm-save-modelo').disabled = errosModelo.length > 0;
        const antigos = Object.fromEntries(((modeloEditando && modeloEditando.parametros) || []).map(p => [p.chave, p]));
        const area = document.getElementById('modelo-params-config');
        area.replaceChildren();
        const linhas = params.map(p => {
            const bloco = document.createElement('div');
            bloco.style.cssText = 'margin-bottom:10px; border-bottom:1px solid #30363d; padding-bottom:8px;';
            const onde = document.createElement('div');
            onde.style.cssText = 'font-size:0.7rem; color:#8b949e;';
            onde.textContent = (p.onde || []).join(' · ');
            const rot = document.createElement('input');
            rot.type = 'text'; rot.className = 'modal-input'; rot.placeholder = 'Rótulo (ex: Comprimento do vão)';
            rot.value = (antigos[p.chave] && antigos[p.chave].rotulo) || p.rotulo;
            let pad;
            const pa = antigos[p.chave] ? antigos[p.chave].padrao : p.padrao;
            if (p.tipo === 'ativo') {        // TASK-060: as opções vêm do texto X(...) da tabela; aqui só se escolhe o padrão (opcional)
                const ops = document.createElement('div');
                ops.style.cssText = 'font-size:0.75rem; color:#8b949e; margin:4px 0;';
                ops.textContent = `Ativo a escolher entre: ${(p.opcoes || []).join(', ') || '(sem opções válidas)'}`;
                pad = document.createElement('select');
                pad.className = 'modal-input';
                pad.append(new Option('Sem padrão (o usuário precisa escolher)', ''));
                (p.opcoes || []).forEach(o => pad.append(new Option(o, o)));
                pad.value = (pa && (p.opcoes || []).includes(String(pa).toUpperCase())) ? String(pa).toUpperCase() : '';
                bloco.append(onde, rot, ops, pad);
            } else {
                pad = document.createElement('input');
                pad.type = 'text'; pad.className = 'modal-input'; pad.placeholder = 'Valor padrão (opcional)'; pad.inputMode = 'decimal';
                pad.value = pa === null || pa === undefined ? '' : String(pa).replace('.', ',');
                bloco.append(onde, rot, pad);
            }
            area.appendChild(bloco);
            return { chave: p.chave, tipo: p.tipo, rot, pad };
        });
        modalSaveModelo.classList.remove('hidden');
        document.getElementById('btn-confirm-save-modelo').onclick = async () => {
            const nome = document.getElementById('input-modelo-nome').value.trim();
            const erro = document.getElementById('modelo-save-erro');
            if (!nome) { erro.textContent = 'Digite um nome para o modelo.'; return; }
            const projeto = projetoAtualObras();
            if (!confirm(`Modelos são públicos: todos os usuários do projeto ${projeto} poderão usá-lo (só você edita). Salvar?`)) return;
            const corpo = {
                id: modeloEditando ? modeloEditando.id : 'modelo_' + Date.now(), nome, data: new Date().toLocaleString(),
                dados_json: JSON.stringify(dados), projeto, tipo: 'modelo',
                parametros: linhas.map(l => ({ chave: l.chave, rotulo: l.rot.value.trim(),
                    padrao: l.tipo === 'ativo' ? (l.pad.value || null) : (l.pad.value.trim().replace(',', '.') || null) }))
            };
            try {
                const r = await fetch('/api/obras', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(corpo) });
                if (!r.ok) { erro.textContent = (await mensagensDeErro(r)).join(' '); return; }
                const salvo = await r.json().catch(() => ({}));
                modalSaveModelo.classList.add('hidden');
                modeloEditando = null; atualizarAvisoModelo();
                showToast('Modelo salvo.' + ((salvo.avisos || []).length ? ' Atenção: ' + salvo.avisos.join(' ') : ''));
            } catch (e) { erro.textContent = 'Erro ao salvar: ' + e.message; }
        };
    });

    let obrasCache = [];

    /** Desenha a lista de obras do cache aplicando os filtros Visibilidade / Origem (TASK-057). */
    function renderObras() {
        const fVis = document.getElementById('filtro-obra-visibilidade').value;
        const fOri = document.getElementById('filtro-obra-origem').value;
        const fTipo = document.getElementById('filtro-obra-tipo').value;
        const obras = obrasCache.filter(o =>
            (fTipo === 'todos' || (fTipo === 'modelos') === (o.tipo === 'modelo')) &&
            (fVis === 'todas' || (fVis === 'publicas') === !!o.publica) && (fOri === 'todas' || (fOri === 'minhas') === !!o.minha));
        const listDiv = document.getElementById('obras-list');
            listDiv.innerHTML = '';
            
            if (obras.length === 0) {
                listDiv.innerHTML = '<div style="color:#8b949e;padding:10px;">Nenhuma obra encontrada.</div>';
            } else {
                obras.forEach(o => {
                    // Parse do JSON salvo (pode ser tableStates completo ou snap simples)
                    let snap;
                    try {
                        snap = typeof o.dados_json === 'string' ? JSON.parse(o.dados_json) : o.dados_json;
                    } catch (err) {
                        snap = null;
                    }

                    const item = document.createElement('div');
                    item.className = 'obra-item';
                    item.innerHTML = `
                        <div class="obra-info">
                            <span class="obra-nome"></span>
                            <span class="obra-data"></span>
                            <span class="obra-meta" style="font-size:0.72rem; color:#8b949e;"></span>
                        </div>
                        <div class="obra-actions">
                            <button class="btn-primary btn-load-item"  style="background:#238636; border:none; padding: 4px 8px; border-radius: 4px; color: white;">Carregar</button>
                            <button class="btn-secondary btn-add-item"  style="background:#1f6feb; border:none; padding: 4px 8px; border-radius: 4px; color: white;">Adicionar</button>
                            <button class="btn-secondary btn-sub-item"  style="background:#d29922; border:none; padding: 4px 8px; border-radius: 4px; color: white;">Subtrair</button>
                            <button class="btn-secondary btn-cmp-item" type="button">Comparar</button>
                            <button class="btn-secondary btn-exp-item"  style="background:#6e7681; border:none; padding: 4px 8px; border-radius: 4px; color: white;" title="Baixa a obra como arquivo .obra.json">Exportar</button>
                            <button class="btn-secondary btn-vis-item"  style="background:#8957e5; border:none; padding: 4px 8px; border-radius: 4px; color: white; display:none;"></button>
                            <button class="btn-secondary btn-own-item"  style="background:#bf8700; border:none; padding: 4px 8px; border-radius: 4px; color: white; display:none;" title="Torna esta obra antiga (sem dono) uma obra particular sua">Assumir</button>
                            <button class="btn-secondary btn-del-item"  style="background:#da3633; border:none; padding: 4px 8px; border-radius: 4px; color: white;">Excluir</button>
                        </div>
                    `;

                    item.querySelector('.obra-nome').textContent = o.nome;   // textContent: nome de obra nunca vira HTML
                    item.querySelector('.obra-data').textContent = o.data;
                    item.querySelector('.btn-del-item').dataset.id = o.id;
                    const selo = o.sem_dono ? '⚠️ sem dono (só administradores)' : (o.tipo === 'modelo' ? '📐 Modelo' : (o.publica ? '🌐 Pública' : '🔒 Particular'));
                    item.querySelector('.obra-meta').textContent = o.minha ? selo : (o.sem_dono ? selo : `${selo} · de ${o.dono || 'outro usuário'}`);

                    const btnVis = item.querySelector('.btn-vis-item');
                    const btnOwn = item.querySelector('.btn-own-item');
                    const btnDel0 = item.querySelector('.btn-del-item');
                    if (!o.minha) btnDel0.style.display = 'none';      // só o dono apaga
                    if (o.minha && o.tipo === 'modelo') {       // só o criador edita o modelo (TASK-058)
                        const btnEd = document.createElement('button');
                        btnEd.className = 'btn-secondary';
                        btnEd.style.cssText = 'background:#8957e5; border:none; padding: 4px 8px; border-radius: 4px; color: white;';
                        btnEd.textContent = 'Editar modelo';
                        btnEd.title = 'Abre o modelo nas tabelas (com os V) para editar; depois use "Salvar como modelo"';
                        btnEd.addEventListener('click', () => {
                            if (!btnLoad._snap) { showToast('Dados do modelo inválidos.'); return; }
                            if (!confirm('As tabelas atuais serão substituídas pelo modelo (Ctrl+Z desfaz). Continuar?')) return;
                            restoreObraSnapshot(btnLoad._snap);
                            definirModeloEditando({ id: o.id, nome: o.nome, parametros: [] });
                            modalLoadObra.classList.add('hidden');
                            showToast('Modelo aberto para edição. Altere e clique em "Salvar como modelo".');
                        });
                        btnVis.insertAdjacentElement('beforebegin', btnEd);
                    }
                    if (o.minha && o.tipo !== 'modelo') {
                        btnVis.style.display = '';
                        btnVis.textContent = o.publica ? 'Tornar particular' : 'Tornar pública';
                        btnVis.addEventListener('click', async () => {
                            const tornarPublica = !o.publica;
                            const proj = localStorage.getItem('projeto_selecionado_codigo') || '229';
                            if (tornarPublica && !confirm(`Todos os usuários do projeto ${proj} poderão ver e copiar esta obra. Deseja torná-la pública?`)) return;
                            btnVis.disabled = true;
                            try {
                                const r = await fetch(`/api/obras/${encodeURIComponent(o.id)}/visibilidade`, {
                                    method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ publica: tornarPublica })
                                });
                                if (!r.ok) throw new Error(await lerDetalhe(r));
                                const resp = await r.json();
                                o.publica = tornarPublica;
                                showToast(resp.aviso || (tornarPublica ? 'Obra agora é pública no projeto.' : 'Obra agora é particular.'));
                                renderObras();
                            } catch (err) {
                                showToast('Erro ao alterar a visibilidade: ' + err.message);
                                btnVis.disabled = false;
                            }
                        });
                    }
                    if (o.sem_dono) {
                        btnOwn.style.display = '';
                        btnOwn.addEventListener('click', async () => {
                            btnOwn.disabled = true;
                            try {
                                const r = await fetch(`/api/obras/${encodeURIComponent(o.id)}/assumir`, { method: 'POST' });
                                if (!r.ok) throw new Error(await lerDetalhe(r));
                                o.sem_dono = false; o.minha = true; o.publica = false; o.dono = '';
                                showToast('Obra assumida: agora é particular sua.');
                                renderObras();
                            } catch (err) {
                                showToast('Erro ao assumir a obra: ' + err.message);
                                btnOwn.disabled = false;
                            }
                        });
                    }

                    // Guarda o snap diretamente no elemento via propriedade JS (sem HTML encoding)
                    const btnLoad = item.querySelector('.btn-load-item');
                    const btnAdd  = item.querySelector('.btn-add-item');
                    const btnSub  = item.querySelector('.btn-sub-item');
                    const btnDel  = item.querySelector('.btn-del-item');

                    btnLoad._snap = snap;
                    btnAdd._snap  = snap;
                    btnSub._snap  = snap;

                    const snapDe = async (b) => (o.tipo === 'modelo' ? pedirParametrosModelo(o) : b._snap);

                    btnLoad.addEventListener('click', async () => {
                        const s = await snapDe(btnLoad);
                        if (!s) { if (o.tipo !== 'modelo') showToast('Dados da obra inválidos.'); return; }
                        restoreObraSnapshot(s);
                        modalLoadObra.classList.add('hidden');
                        showToast('Obra carregada!');
                    });

                    btnAdd.addEventListener('click', async () => {
                        const s = await snapDe(btnAdd);
                        if (!s) { if (o.tipo !== 'modelo') showToast('Dados da obra inválidos.'); return; }
                        if (adicionarObraAoProjeto(s)) {
                            modalLoadObra.classList.add('hidden');
                            showToast('Obra adicionada!');
                        }
                    });

                    btnSub.addEventListener('click', async () => {
                        const s = await snapDe(btnSub);
                        if (!s) { if (o.tipo !== 'modelo') showToast('Dados da obra inválidos.'); return; }
                        if (subtrairObraDoProjeto(s)) {
                            modalLoadObra.classList.add('hidden');
                            showToast('Obra subtraída!');
                        }
                    });

                    const btnCmp = item.querySelector('.btn-cmp-item');
                    btnCmp._snap = snap;
                    btnCmp.addEventListener('click', async () => {
                        const s = await snapDe(btnCmp);
                        if (!s) { if (o.tipo !== 'modelo') showToast('Dados da obra inválidos.'); return; }
                        if (compararObraNoResumo(s, o.nome)) modalLoadObra.classList.add('hidden');
                    });

                    const btnExp = item.querySelector('.btn-exp-item');
                    btnExp.addEventListener('click', async () => {
                        btnExp.disabled = true;
                        try {
                            const projExp = localStorage.getItem('projeto_selecionado_codigo') || '229';
                            const r = await fetch(`/api/obras/${encodeURIComponent(o.id)}/exportar?projeto=${encodeURIComponent(projExp)}`);
                            if (!r.ok) throw new Error(await lerDetalhe(r));
                            const blob = new Blob([JSON.stringify(await r.json(), null, 2)], { type: 'application/json' });
                            const arquivo = `${String(o.nome || 'obra').replace(/[\\/:*?"<>|]+/g, '_').slice(0, 80)}.obra.json`;
                            const a = document.createElement('a');
                            a.href = URL.createObjectURL(blob);
                            a.download = arquivo;
                            document.body.appendChild(a);
                            a.click();
                            a.remove();
                            setTimeout(() => URL.revokeObjectURL(a.href), 1000);
                            showToast('Obra exportada.');
                        } catch (err) {
                            showToast('Erro ao exportar a obra: ' + err.message);
                        } finally {
                            btnExp.disabled = false;
                        }
                    });

                    btnDel.addEventListener('click', async () => {
                        btnDel.disabled = true;
                        try {
                            const rd = await fetch(`/api/obras/${o.id}`, { method: 'DELETE' });
                            if (!rd.ok) throw new Error('falha');
                            obrasCache = obrasCache.filter(x => x.id !== o.id);
                            item.remove();
                        } catch (err) {
                            btnDel.disabled = false;
                        }
                    });

                    listDiv.appendChild(item);
                });
            }
    }
    ['filtro-obra-visibilidade', 'filtro-obra-origem', 'filtro-obra-tipo'].forEach(id => document.getElementById(id).addEventListener('change', renderObras));

    async function abrirModalObras() {
        try {
            const projCodeObras = localStorage.getItem('projeto_selecionado_codigo') || '229';
            const res = await fetch(`/api/obras?projeto=${encodeURIComponent(projCodeObras)}`);
            if (!res.ok) throw new Error('Falha ao listar obras');
            const obras = await res.json();
            
            obrasCache = obras;
            renderObras();
            modalLoadObra.classList.remove('hidden');
        } catch (e) {
            console.error('Erro ao carregar obras:', e);
            showToast('Erro ao carregar obras');
        }
    }
    document.getElementById('btn-load-obra').addEventListener('click', abrirModalObras);

    // Exportar JSON da obra que está na tela (TASK-111): mesmo formato .obra.json da exportação de uma obra salva, sem precisar salvá-la antes.
    document.getElementById('btn-exportar-json').addEventListener('click', async () => {
        const btn = document.getElementById('btn-exportar-json');
        btn.disabled = true;
        try {
            const numero = (document.getElementById('input-num-obra').value || '').trim();
            const projeto = localStorage.getItem('projeto_selecionado_codigo') || '229';
            const r = await fetch(`/api/obras/exportar-atual?projeto=${encodeURIComponent(projeto)}`, {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ nome: numero || 'obra', dados_json: JSON.stringify(tableStates) })
            });
            if (!r.ok) throw new Error(await lerDetalhe(r));
            const blob = new Blob([JSON.stringify(await r.json(), null, 2)], { type: 'application/json' });
            const arquivo = `${String(numero || 'obra').replace(/[\\/:*?"<>|]+/g, '_').slice(0, 80)}.obra.json`;
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = arquivo;
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(a.href), 1000);
            showToast('Obra exportada (.obra.json).');
        } catch (err) {
            showToast('Erro ao exportar a obra: ' + err.message);
        } finally {
            btn.disabled = false;
        }
    });

    // Importar obra (TASK-056): arquivo .obra.json → cópia PARTICULAR do usuário, só se for do projeto selecionado.
    const inputImportarObra = document.getElementById('input-importar-obra');
    document.getElementById('btn-importar-obra').addEventListener('click', () => inputImportarObra.click());
    inputImportarObra.addEventListener('change', async () => {
        const arquivo = inputImportarObra.files[0];
        inputImportarObra.value = '';
        const msg = document.getElementById('importar-obra-msg');
        const dizer = (texto, erro) => { msg.style.color = erro ? '#f85149' : '#3fb950'; msg.textContent = texto; };
        msg.textContent = '';
        if (!arquivo) return;
        if (arquivo.size > 5 * 1024 * 1024) { dizer('Arquivo grande demais (máx. 5 MB).', true); return; }
        try {
            const projeto = localStorage.getItem('projeto_selecionado_codigo') || '229';
            const r = await fetch(`/api/obras/importar?projeto=${encodeURIComponent(projeto)}`, {
                method: 'POST', headers: { 'Content-Type': 'application/json' }, body: await arquivo.text()
            });
            const d = await r.json().catch(() => ({}));
            if (!r.ok) {
                const erros = d.detail && Array.isArray(d.detail.erros) ? d.detail.erros : [typeof d.detail === 'string' ? d.detail : 'Falha ao importar a obra.'];
                dizer(erros.join(' · '), true);
                return;
            }
            await abrirModalObras();
            dizer(`Importada: “${d.nome}” (${d.cabos} linha(s) de Cabos, ${d.outros} de Outros) — particular, só sua.`, false);
        } catch (e) {
            dizer('Não foi possível importar: ' + e.message, true);
        }
    });

    /* TASK-118: "Importar JSON" da barra de título — abre o arquivo .obra.json e pergunta direto: Carregar (substitui a obra da tela),
       Adicionar, Subtrair ou Comparar. O arquivo é validado em /api/obras/ler-arquivo.
       Modelo com variáveis V continua sendo importado como modelo salvo (precisa dos parâmetros) e aparece em Carregar Obra. */
    const inputImportarJson = document.getElementById('input-importar-json');
    function dialogoImportarJson(info) {
        return new Promise(resolve => {
            const fundo = document.createElement('div');
            fundo.className = 'modal-overlay';
            fundo.innerHTML = `<div class="modal-content" style="max-width:480px; width:92%;">
                <h3>Importar JSON</h3>
                <p style="margin:0 0 6px; font-size:0.86rem;"><b></b></p>
                <p style="margin:0 0 12px; font-size:0.78rem; color:#8b949e;"></p>
                <div style="display:flex; flex-direction:column; gap:6px; font-size:0.8rem;">
                    <button type="button" class="btn-primary" data-acao="carregar">Carregar — substitui as tabelas da tela</button>
                    <button type="button" class="btn-secondary" data-acao="adicionar">Adicionar — acrescenta as linhas no fim</button>
                    <button type="button" class="btn-secondary" data-acao="subtrair">Subtrair — acrescenta como retirada (cabos negativos e * nos ativos)</button>
                    <button type="button" class="btn-secondary" data-acao="comparar">Comparar — mostra as diferenças nas tabelas</button>
                </div>
                <div class="modal-actions" style="margin-top:12px;"><button type="button" class="btn-secondary" data-acao="cancelar">Cancelar</button></div>
            </div>`;
            fundo.querySelector('b').textContent = info.nome;
            fundo.querySelectorAll('p')[1].textContent = `${info.cabos} linha(s) de Cabos e ${info.outros} de Outros. Alterações nas tabelas podem ser desfeitas com Ctrl+Z.`;
            const fechar = r => { document.removeEventListener('keydown', tecla, true); fundo.remove(); resolve(r); };
            const tecla = e => { if (e.key === 'Escape') { e.stopPropagation(); fechar('cancelar'); } };
            fundo.addEventListener('click', e => { const a = e.target.closest && e.target.closest('[data-acao]'); if (a) fechar(a.dataset.acao); else if (e.target === fundo) fechar('cancelar'); });
            document.addEventListener('keydown', tecla, true);
            document.body.appendChild(fundo);
        });
    }
    if (inputImportarJson) inputImportarJson.addEventListener('change', async () => {
        const arquivo = inputImportarJson.files[0];
        inputImportarJson.value = '';
        if (!arquivo) return;
        if (arquivo.size > 5 * 1024 * 1024) { showToast('Arquivo grande demais (máx. 5 MB).'); return; }
        try {
            const projeto = localStorage.getItem('projeto_selecionado_codigo') || '229';
            const texto = await arquivo.text();
            const r = await fetch(`/api/obras/ler-arquivo?projeto=${encodeURIComponent(projeto)}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: texto });
            const d = await r.json().catch(() => ({}));
            if (!r.ok) {
                const erros = d.detail && Array.isArray(d.detail.erros) ? d.detail.erros : [typeof d.detail === 'string' ? d.detail : 'Falha ao ler o arquivo.'];
                alert('Não foi possível importar:\n' + erros.join('\n'));
                return;
            }
            if (d.tipo === 'modelo') {                                   // modelo: importa como modelo salvo (os parâmetros são pedidos ao usar)
                const r2 = await fetch(`/api/obras/importar?projeto=${encodeURIComponent(projeto)}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: texto });
                if (!r2.ok) { alert('Não foi possível importar o modelo:\n' + (await mensagensDeErro(r2)).join('\n')); return; }
                showToast(`Modelo “${d.nome}” importado. Use-o em Carregar Obra.`);
                await abrirModalObras();
                return;
            }
            const acao = await dialogoImportarJson(d);
            if (acao === 'carregar') { restoreObraSnapshot(d.dados); showToast(`Obra “${d.nome}” carregada.`); }
            else if (acao === 'adicionar') { if (adicionarObraAoProjeto(d.dados)) showToast(`Obra “${d.nome}” adicionada.`); }
            else if (acao === 'subtrair') { if (subtrairObraDoProjeto(d.dados)) showToast(`Obra “${d.nome}” subtraída.`); }
            else if (acao === 'comparar') compararObraNoResumo(d.dados, d.nome);
        } catch (e) {
            showToast('Não foi possível importar: ' + e.message);
        }
    });

    /**
     * Payload de cálculo do orçamento (o mesmo que "Montar Orçamento" grava em orcamentoPayload):
     * sincroniza a Totalizadora no modo padrão e a converte para os formatos que o backend lê.
     * Extraída do handler do botão (TASK-014) para a validação usar exatamente o mesmo payload.
     */
    async function obterPayloadCalculo() {
        // Se estivermos no modo padrão (Cabos e Postes), atualiza a totalizadora silenciosamente antes
        const radioPadrao = document.querySelector('input[name="view_mode"][value="padrao"]');
        if (radioPadrao && radioPadrao.checked) {
            await syncTotalizadora(false);
        }
        return payloadDaTotalizadora(tableStates.totalizadora.data);
    }

    function payloadDaTotalizadora(totalizadora) {
        const payloadCabos = [];
        const payloadOutros = [];
        totalizadora.forEach(item => {
            if (!item) return;
            
            let qStr = item.qtd;
            if (qStr === '' || qStr === null || qStr === undefined || parseFloat(qStr) === 0) {
                return; // Ignora itens com quantidade vazia ou zerada no cálculo do orçamento
            }

            let operacao = item.operacao || 'I';
            
            // Reconstruir o campo "ativo"
            let ativoFinal = item.ativo;
            
            if (typeof qStr === 'number' && qStr < 0) {
                qStr = '*' + Math.abs(qStr);
            } else if (typeof qStr === 'string' && qStr.startsWith('-')) {
                qStr = '*' + qStr.substring(1);
            }

            if (item.origem === 'CABOS') {
                // Backend espera: [Nome] [Fases] [Comprimento]. Enviamos o nome e a qtd (comprimento)
                ativoFinal = `${item.ativo} 1 ${item.qtd}`; 
            } else {
                // Ex: 4-TERRA3 ou *1-TERRA3 para negativos
                ativoFinal = `${qStr}-${item.ativo}`;
            }

            const obj = {
                entidade: item.obs || '0',
                operacao: operacao,
                ativo: ativoFinal
            };

            if (item.origem === 'CABOS') {
                payloadCabos.push(obj);
            } else {
                payloadOutros.push(obj);
            }
        });

        return { cabos: payloadCabos, outros: payloadOutros };
    }

    /* ═══════════════════════════════════════
       VALIDAÇÃO DAS PLANILHAS (TASK-014)
       Camadas 1 e 2 (contrato e regras de domínio) vêm de /api/validacao/planilhas e nunca dependem da
       IA; a camada 3 (/api/validacao/ia) é opcional e, se falhar, só vira um aviso no painel.
    ═══════════════════════════════════════ */
    const VALIDACAO_AUTO_CHAVE = 'validacao_auto';
    const ROTULO_SEVERIDADE = { erro: 'Erro', aviso: 'Aviso', info: 'Info' };
    const COR_SEVERIDADE = { erro: '#f85149', aviso: '#d29922', info: '#58a6ff' };
    const ORDEM_SEVERIDADE = { erro: 0, aviso: 1, info: 2 };
    let ultimoPayloadValidacao = { cabos: [], outros: [] };

    function validacaoAutomaticaLigada() {
        try { return localStorage.getItem(VALIDACAO_AUTO_CHAVE) === 'sim'; } catch (e) { return false; }
    }

    function definirValidacaoAutomatica(ligada) {
        try { localStorage.setItem(VALIDACAO_AUTO_CHAVE, ligada ? 'sim' : 'nao'); } catch (e) { /* sem storage: vale só nesta sessão */ }
        atualizarRotuloValidar();
    }

    /* Modos de validação (TASK-020): determinística (contrato + regras de domínio) e IA, ligáveis em separado.
       Preferência local (localStorage), como "Validar ao montar". Os controles ficam na gaveta (aba Execução,
       painel_regras.js, TASK-026); aqui só se lê/grava e se mostra o resumo ao lado do botão Validar. */
    const MODOS_CHAVE = 'validacao_modos';
    const MODOS_PADRAO = { det: true, dominio: true, ia: true, pularIa: true };
    function lerModos() {
        try { return { ...MODOS_PADRAO, ...(JSON.parse(localStorage.getItem(MODOS_CHAVE)) || {}) }; } catch (e) { return { ...MODOS_PADRAO }; }
    }
    function gravarModos(m) {
        try { localStorage.setItem(MODOS_CHAVE, JSON.stringify(m)); } catch (e) { /* vale só nesta sessão */ }
        atualizarRotuloValidar();
    }

    /** Resumo curto (rótulo) e longo (dica) do que o botão Validar roda com as preferências atuais. */
    function resumoDoQueRoda() {
        const m = lerModos();
        const curto = [];
        const longo = [];
        if (m.det) {
            curto.push(m.dominio ? 'Det + Regras' : 'Det');
            longo.push(m.dominio ? 'Determinística (contrato + regras de domínio)' : 'Determinística (só contrato)');
        } else longo.push('Determinística desligada');
        if (m.ia) {
            curto.push('IA');
            longo.push(m.det && m.pularIa ? 'IA (pulada se o contrato tiver erro)' : 'IA');
        } else longo.push('IA desligada');
        const auto = validacaoAutomaticaLigada();
        return {
            curto: (curto.length ? curto.join(' + ') : 'nada ligado') + (auto ? ' · ao montar' : ''),
            longo: `${longo.join(' · ')}. ${auto ? 'Também roda ao montar o orçamento.' : 'Não roda ao montar o orçamento.'}`
        };
    }

    function atualizarRotuloValidar() {
        const r = resumoDoQueRoda();
        const el = document.getElementById('validar-estado');
        if (el) { el.textContent = r.curto; el.title = `Validar roda: ${r.longo} (mude em Regras de validação › Execução)`; }
        const btn = document.getElementById('btn-validar');
        if (btn) btn.title = `Verificar as planilhas Cabos e Outros. Roda: ${r.longo}`;
    }

    // API usada pela aba Execução da gaveta (painel_regras.js): mesma lógica, sem segunda cópia.
    window.validacaoPrefs = { lerModos, gravarModos, validacaoAutomaticaLigada, definirValidacaoAutomatica, resumoDoQueRoda };
    atualizarRotuloValidar();
    const btnOpcoes = document.getElementById('btn-validacao-opcoes');
    if (btnOpcoes) btnOpcoes.addEventListener('click', () => { if (typeof window.rpAbrirNaAba === 'function') window.rpAbrirNaAba('execucao'); });

    function linhasParaValidacao(tabela, prefixo) {
        const linhas = [];
        tableStates[tabela].data.forEach((r, i) => {
            if (r && !r.congelada) linhas.push({ id: `${prefixo}-${i}`, operacao: r.operacao || '', ativo: r.ativo || '' });          // congelada: a validação não a vê
        });
        return linhas;
    }

    /** Linhas atuais das duas tabelas, para a pré-visualização dos ajustes (painel de regras). */
    window.resumoLinhasParaAjuste = () => ({ cabos: linhasParaValidacao('cabos', 'CABOS'), outros: linhasParaValidacao('outros', 'OUTROS') });

    function cabecalhoIA() {
        return {
            'Content-Type': 'application/json',
            'X-Gemini-Key': localStorage.getItem('gemini_api_key') || 'SAVED_IN_BACKEND'
        };
    }

    async function lerDetalhe(resp) {
        try {
            const j = await resp.json();
            if (j.detail && Array.isArray(j.detail.erros)) return j.detail.erros.join('; ');
            return typeof j.detail === 'string' ? j.detail : `HTTP ${resp.status}`;
        }
        catch (e) { return `HTTP ${resp.status}`; }
    }

    /** Código do projeto selecionado no Resumo ('' se nenhum). */
    function codigoDoProjetoSelecionado() {
        const sel = document.getElementById('select-projeto');
        return sel && sel.selectedIndex >= 0 ? (sel.options[sel.selectedIndex].dataset.codigo || '') : '';
    }

    /* ── TASK-044: seletor de Contexto (subconjunto de Ajustes dentro do projeto) ──
       Terceira camada sobre padrão+projeto: nenhuma troca aqui recalcula as tabelas na hora — só
       define o que o botão "Ajustar" e o editor de Ajustes (gaveta) vão usar na próxima chamada. */
    function _chaveContextoLocalStorage(projCode) {
        return `contexto_selecionado_${projCode || 'DEFAULT'}`;
    }

    /** Nome do contexto selecionado ('' tratado como "Nenhum") — lida por painel_ajustes.js e pelo fluxo de Ajustar. */
    window.contextoSelecionado = function () {
        const sel = document.getElementById('select-contexto');
        return sel ? (sel.value || null) : null;
    };

    function _atualizarBotoesContexto() {
        const sel = document.getElementById('select-contexto');
        const btnNovo = document.getElementById('btn-novo-contexto');
        const btnExcluir = document.getElementById('btn-excluir-contexto');
        const admin = localStorage.getItem('is_admin') === 'true';
        if (btnNovo) btnNovo.style.display = admin ? 'inline-block' : 'none';
        if (btnExcluir) btnExcluir.style.display = (admin && sel && sel.value) ? 'inline-block' : 'none';
    }

    /** Recarrega a lista de contextos do projeto dado e restaura a seleção salva (se ainda existir). */
    window.carregarContextosAjustes = async function (projCode) {
        const sel = document.getElementById('select-contexto');
        if (!sel) return;
        sel.innerHTML = '';
        sel.appendChild(new Option('Nenhum', ''));
        if (!projCode || projCode === 'DEFAULT') { _atualizarBotoesContexto(); return; }
        try {
            const res = await fetch(`/api/validacao/ajustes/contextos?projeto_codigo=${encodeURIComponent(projCode)}`);
            const data = res.ok ? await res.json() : { contextos: [] };
            (data.contextos || []).forEach(c => sel.appendChild(new Option(c, c)));
            const salvo = localStorage.getItem(_chaveContextoLocalStorage(projCode)) || '';
            sel.value = (data.contextos || []).includes(salvo) ? salvo : '';
        } catch (e) {
            console.error('Erro ao carregar contextos:', e);
        }
        _atualizarBotoesContexto();
    };

    window.abrirNovoContexto = async function () {
        const projCode = codigoDoProjetoSelecionado();
        if (!projCode || projCode === 'DEFAULT') { showToast('Selecione um projeto primeiro.'); return; }
        const nome = (prompt('Nome do novo contexto (ex.: "34,5kV"):') || '').trim();
        if (!nome) return;
        try {
            const res = await fetch('/api/validacao/ajustes/contextos', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ projeto_codigo: projCode, contexto: nome })
            });
            if (res.ok) {
                await window.carregarContextosAjustes(projCode);
                document.getElementById('select-contexto').value = nome;
                localStorage.setItem(_chaveContextoLocalStorage(projCode), nome);
                _atualizarBotoesContexto();
                showToast('✓ Contexto criado!');
                if (typeof window.ajCarregar === 'function') window.ajCarregar();
            } else {
                const data = await res.json().catch(() => ({}));
                showToast((data && data.detail) || 'Erro ao criar contexto.');
            }
        } catch (e) {
            showToast('Erro de conexão ao criar contexto.');
        }
    };

    window.excluirContextoAtual = async function () {
        const projCode = codigoDoProjetoSelecionado();
        const sel = document.getElementById('select-contexto');
        const contexto = sel ? sel.value : '';
        if (!projCode || !contexto) return;
        if (!confirm(`Excluir o contexto "${contexto}"? Os ajustes ligados/desligados só dentro dele serão perdidos.`)) return;
        try {
            const res = await fetch('/api/validacao/ajustes/contextos/excluir', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ projeto_codigo: projCode, contexto })
            });
            if (res.ok) {
                localStorage.removeItem(_chaveContextoLocalStorage(projCode));
                await window.carregarContextosAjustes(projCode);
                showToast('Contexto excluído.');
                if (typeof window.ajCarregar === 'function') window.ajCarregar();
            } else {
                showToast('Erro ao excluir contexto.');
            }
        } catch (e) {
            showToast('Erro de conexão ao excluir contexto.');
        }
    };

    const _selectContexto = document.getElementById('select-contexto');
    if (_selectContexto) {
        _selectContexto.addEventListener('change', () => {
            const projCode = codigoDoProjetoSelecionado();
            localStorage.setItem(_chaveContextoLocalStorage(projCode), _selectContexto.value);
            _atualizarBotoesContexto();
            if (typeof window.ajCarregar === 'function') window.ajCarregar();
            if (window.GridRProjeto && GridRProjeto.marcarSujo) GridRProjeto.marcarSujo(false);          // TASK-165: o contexto vai no .obra
        });
    }

    const nomesDeAjuste = {};   // id → nome, para os títulos das pré-visualizações

    /** Ajustes (receitas) cadastrados e LIGADOS para o projeto, mais o mapa regra → ajustes que a corrigem. Falha = sem ajustes. */
    async function carregarAjustesCadastrados() {
        try {
            const ctx = window.contextoSelecionado ? window.contextoSelecionado() : null;
            const qs = `projeto_codigo=${encodeURIComponent(codigoDoProjetoSelecionado() || 'DEFAULT')}` + (ctx ? `&contexto=${encodeURIComponent(ctx)}` : '');
            const r = await fetch(`/api/validacao/ajustes?${qs}`);
            if (r.ok) {
                const d = await r.json();
                d.ajustes.forEach(a => { nomesDeAjuste[a.id] = a.nome; });
                return { ajustes: d.ajustes.filter(a => a.ativa && !a.oculta), por_regra: d.por_regra || {} };
            }
        } catch (e) { /* sem ajustes: a validação segue normalmente */ }
        return { ajustes: [], por_regra: {} };
    }

    /** Roda os modos ligados: determinística (camadas 1-2) e, depois, a IA. Devolve { achados, ia, modos, falhou }.
     *  `falhou` = a determinística foi pedida e não pôde rodar. `modos` diz o que rodou (a tela mostra ao usuário). */
    /** `somenteContrato` (TASK-051): ignora a preferência salva e força só Camada 1 + ativo não
     * encontrado (sem regras de domínio nem IA) — usado pela checagem obrigatória antes de Ajustar/
     * Montar Orçamento, independente do que o usuário tiver configurado em "Validar ao montar". */
    async function executarValidacao(somenteContrato = false) {
        const prefs = somenteContrato ? { det: true, dominio: false, ia: false, pularIa: true } : lerModos();
        const modos = { det: prefs.det, dominio: prefs.det && prefs.dominio, ia: 'desligada' };
        if (!prefs.det && !prefs.ia) return { achados: [], ia: null, modos, falhou: null, nada: true };
        let ajustes = { ajustes: [], por_regra: {} };

        const cabos = linhasParaValidacao('cabos', 'CABOS');
        const outros = linhasParaValidacao('outros', 'OUTROS');
        const selectProj = document.getElementById('select-projeto');
        const projeto = selectProj ? selectProj.value : '';
        const projetoCodigo = selectProj && selectProj.selectedIndex >= 0 ? (selectProj.options[selectProj.selectedIndex].dataset.codigo || '') : '';

        let achados = [];
        if (prefs.det) {
            // Payload da Totalizadora (após as regras de conversão) só para checar o ativo na base técnica.
            // Ids próprios para não colidirem com CABOS-<i>/OUTROS-<i> das tabelas.
            let payloadCalculo = null;
            try {
                const p = await obterPayloadCalculo();
                payloadCalculo = {
                    cabos: p.cabos.map((o, i) => ({ ...o, id: `TOTALIZADORA-CABOS-${i}` })),
                    outros: p.outros.map((o, i) => ({ ...o, id: `TOTALIZADORA-OUTROS-${i}` }))
                };
            } catch (e) { /* segue sem a checagem de base */ }
            ultimoPayloadValidacao = payloadCalculo || { cabos: [], outros: [] };

            try {
                const resp = await fetch('/api/validacao/planilhas', {
                    method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ cabos, outros, projeto, projeto_codigo: projetoCodigo || null,
                                           incluir_dominio: prefs.dominio, payload_calculo: payloadCalculo })
                });
                if (!resp.ok) throw new Error(await lerDetalhe(resp));
                achados = (await resp.json()).achados;
                // TASK-033: achados de vinculação cabo<->estrutura (TASK-032) — só existem no
                // frontend (o vínculo nunca vai ao backend), por isso são calculados aqui e
                // concatenados, em vez de vir de /api/validacao/planilhas.
                achados = achados.concat(avaliarVinculacaoLocal());
            } catch (e) {
                return { achados: [], ia: null, modos, falhou: `Não foi possível validar (${e.message}).` };
            }
            ajustes = await carregarAjustesCadastrados();
        }

        let ia = null;
        if (prefs.ia) {
            const erroDeContrato = achados.some(a => a.severidade === 'erro' && camadaDoAchado(a) === 'Contrato');
            if (prefs.det && prefs.pularIa && erroDeContrato) {
                modos.ia = 'pulada';  // contrato com erro: não gasta a chave de IA
            } else {
                modos.ia = 'rodou';
                ia = { status: 'erro', mensagem: 'sem resposta', achados: [] };
                try {
                    const resp = await fetch('/api/validacao/ia', {
                        method: 'POST', headers: cabecalhoIA(),
                        body: JSON.stringify({
                            cabos, outros, projeto_codigo: projetoCodigo || null,
                            achados_previos: achados.map(a => ({ linha_id: a.linha_id, regra_id: a.regra_id, mensagem: a.mensagem }))
                        })
                    });
                    ia = resp.ok ? await resp.json() : { status: 'erro', mensagem: await lerDetalhe(resp), achados: [] };
                } catch (e) {
                    ia = { status: 'erro', mensagem: e.message, achados: [] };
                }
            }
        }
        return { achados: achados.concat((ia && ia.achados) || []), ia, modos, ajustes, falhou: null };
    }

    /** Frase que diz ao usuário o que realmente rodou. */
    function descreverModos(modos) {
        const partes = [];
        if (modos.det) partes.push(modos.dominio ? 'Determinística (contrato + regras de domínio)' : 'Determinística (só contrato)');
        else partes.push('Determinística desligada');
        partes.push({ rodou: 'IA executada', pulada: 'IA não executada (erro de contrato)', desligada: 'IA desligada' }[modos.ia]);
        return partes.join(' · ');
    }

    function textoDaLinha(achado) {
        if (achado.linha_id === 'GERAL') return 'Planilha inteira';
        let m = /^(CABOS|OUTROS)-(\d+)$/.exec(achado.linha_id);
        if (m) {
            const linha = tableStates[m[1].toLowerCase()].data[Number(m[2])];
            return linha ? `${achado.linha_id} — ${linha.ativo}` : achado.linha_id;
        }
        m = /^TOTALIZADORA-(CABOS|OUTROS)-(\d+)$/.exec(achado.linha_id);
        if (m) {
            const item = ultimoPayloadValidacao[m[1].toLowerCase()][Number(m[2])];
            return item ? `Totalizadora (${m[1].toLowerCase()}) — ${item.ativo}` : achado.linha_id;
        }
        return achado.linha_id;
    }

    function camadaDoAchado(achado) {
        if (achado.origem === 'ia' || String(achado.regra_id).startsWith('IA:')) return 'IA';
        return String(achado.regra_id).startsWith('C2-') ? 'Domínio' : 'Contrato';
    }

    function botaoPainel(texto, classe) {
        const b = document.createElement('button');
        b.className = classe;
        b.textContent = texto;
        return b;
    }

    /** Ajustes cadastrados e ligados que corrigem a regra do achado (ids), na ordem do cadastro. */
    function ajustesDoAchado(res, achado) {
        const ids = (res.ajustes && res.ajustes.por_regra[achado.regra_id]) || [];
        return ids.map(id => res.ajustes.ajustes.find(a => a.id === id)).filter(Boolean);
    }

    function contagemPorSeveridade(achados) {
        const c = { erro: 0, aviso: 0, info: 0 };
        achados.forEach(a => { c[a.severidade] = (c[a.severidade] || 0) + 1; });
        return c;
    }

    function textoContagem(c) { return `${c.erro} erro(s), ${c.aviso} aviso(s), ${c.info} informação(ões)`; }

    /**
     * Mostra o painel. `confirmar`: pergunta "continuar mesmo assim?". `anterior`: contagem antes de um ajuste (mostra o
     * efeito). Resolve 'continuar' | 'corrigir' | 'fechar' | { tipo: 'ajustar', ids } | { tipo: 'ia' } — a decisão fica com o
     * chamador (cicloValidacao), que aplica o ajuste, revalida e reabre o painel.
     */
    function mostrarPainelValidacao(res, confirmar, anterior) {
        return new Promise(resolve => {
            const modal = document.getElementById('modal-validacao');
            const lista = document.getElementById('validacao-lista');
            const acoes = document.getElementById('validacao-acoes');
            const aviso = document.getElementById('validacao-aviso');
            const achados = res.achados.slice().sort((a, b) =>
                (ORDEM_SEVERIDADE[a.severidade] - ORDEM_SEVERIDADE[b.severidade]) || String(a.linha_id).localeCompare(String(b.linha_id), 'pt', { numeric: true }));
            const cont = contagemPorSeveridade(achados);

            document.getElementById('validacao-titulo').textContent = achados.length
                ? 'Problemas encontrados nas planilhas' : 'Nenhum problema encontrado';
            document.getElementById('validacao-resumo').textContent = (achados.length
                ? `${textoContagem(cont)}.` : 'As planilhas passaram nas verificações que rodaram.')
                + (anterior ? ` Antes do ajuste: ${textoContagem(anterior)}.` : '')
                + (res.modos ? ` Rodou: ${descreverModos(res.modos)}.` : '');

            const avisos = [];
            if (res.modos && res.modos.ia === 'pulada') avisos.push('A IA não foi executada porque o contrato das planilhas tem erro: corrija e valide de novo.');
            if (res.ia && res.ia.status !== 'ok') {
                const rotulos = { sem_chave: 'Revisão por IA não executada: sem chave de IA.', desativado: 'Revisão por IA desativada.',
                    indisponivel: 'Revisão por IA indisponível.', parcial: 'Revisão por IA incompleta: parte das linhas falhou.', erro: 'Revisão por IA falhou.' };
                avisos.push((rotulos[res.ia.status] || 'Revisão por IA não concluída.') + (res.ia.mensagem && res.ia.status !== 'desativado' ? ` (${res.ia.mensagem})` : ''));
            }
            if (res.ia && res.ia.truncado) avisos.push('A IA revisou só as primeiras linhas (limite por validação).');
            aviso.textContent = avisos.join(' ');
            aviso.style.display = avisos.length ? 'block' : 'none';

            const fechar = resultado => { modal.classList.add('hidden'); resolve(resultado); };
            const temIA = achados.some(a => camadaDoAchado(a) === 'IA');
            const temDet = achados.some(a => camadaDoAchado(a) !== 'IA');
            let filtro = 'todos';

            function desenhar() {
                lista.replaceChildren();
                if (temIA && temDet) {   // filtro Todos · Determinístico · IA
                    const barra = document.createElement('div');
                    barra.style.cssText = 'display:flex; gap:6px; margin-bottom:8px;';
                    [['todos', 'Todos'], ['det', 'Determinístico'], ['ia', 'IA']].forEach(([k, r]) => {
                        const b = botaoPainel(r, k === filtro ? 'btn-primary' : 'btn-secondary');
                        b.style.padding = '3px 10px'; b.style.fontSize = '0.75rem';
                        b.onclick = () => { filtro = k; desenhar(); };
                        barra.appendChild(b);
                    });
                    lista.appendChild(barra);
                }
                achados.filter(a => filtro === 'todos' || (filtro === 'ia') === (camadaDoAchado(a) === 'IA')).forEach(a => {
                    const item = document.createElement('div');
                    item.style.cssText = `padding: 8px 10px; margin-bottom: 6px; border-left: 3px solid ${COR_SEVERIDADE[a.severidade] || '#8b949e'}; background: rgba(255,255,255,0.03); border-radius: 4px;`;
                    const topo = document.createElement('div');
                    topo.style.cssText = 'display:flex; gap:8px; flex-wrap:wrap; align-items:baseline; font-size:0.75rem; color:#8b949e;';
                    const sev = document.createElement('strong');
                    sev.style.color = COR_SEVERIDADE[a.severidade] || '#8b949e';
                    sev.textContent = ROTULO_SEVERIDADE[a.severidade] || a.severidade;
                    const onde = document.createElement('span');
                    onde.textContent = textoDaLinha(a);
                    const regra = document.createElement('span');
                    regra.textContent = `${camadaDoAchado(a)} · ${a.regra_id}`;
                    topo.append(sev, onde, regra);
                    const msg = document.createElement('div');
                    msg.style.marginTop = '2px';
                    msg.textContent = a.mensagem;
                    item.append(topo, msg);
                    if (a.explicacao) {
                        const exp = document.createElement('div');
                        exp.style.cssText = 'margin-top:2px; font-size:0.78rem; color:#8b949e;';
                        exp.textContent = `Por quê: ${a.explicacao}`;
                        item.appendChild(exp);
                    }
                    if (a.sugestao) {
                        const sug = document.createElement('div');
                        sug.style.cssText = 'margin-top:2px; font-size:0.8rem; color:#3fb950;';
                        sug.textContent = `Sugestão: ${a.sugestao}`;
                        item.appendChild(sug);
                    }
                    ajustesDoAchado(res, a).forEach(aj => {
                        const b = botaoPainel(`Ajustar: ${aj.nome}`, 'btn-secondary');
                        b.style.cssText = 'margin-top:6px; margin-right:6px; padding:3px 10px; font-size:0.75rem; color:#3fb950; border-color:rgba(63,185,80,0.4);';
                        b.title = 'Mostra o que mudaria nas tabelas; nada é alterado sem o seu aceite';
                        b.onclick = () => fechar({ tipo: 'ajustar', ids: [aj.id] });
                        item.appendChild(b);
                    });
                    lista.appendChild(item);
                });
            }
            desenhar();

            acoes.replaceChildren();
            const idsComAjuste = [...new Set(achados.flatMap(a => ajustesDoAchado(res, a).map(x => x.id)))];
            const semAjuste = achados.filter(a => achadoCorrigivel(a) && !ajustesDoAchado(res, a).length);
            const extras = [];
            if (idsComAjuste.length) {
                const todos = botaoPainel('Ajustar tudo (determinístico)', 'btn-secondary');
                todos.title = 'Pré-visualiza todos os ajustes cadastrados que corrigem estes achados';
                todos.onclick = () => fechar({ tipo: 'ajustar', ids: idsComAjuste });
                extras.push(todos);
            }
            if (semAjuste.length && lerModos().ia) {
                const ia = botaoPainel('Pedir ajuste à IA', 'btn-secondary');
                ia.title = 'A IA propõe correções só para os achados sem ajuste cadastrado; você aceita linha a linha';
                ia.onclick = () => fechar({ tipo: 'ia' });
                extras.push(ia);
            }
            if (confirmar) {
                const nao = botaoPainel('Não, vou corrigir', 'btn-secondary');
                nao.onclick = () => fechar('corrigir');
                const sim = botaoPainel('Continuar mesmo assim', 'btn-primary');
                sim.onclick = () => fechar('continuar');
                acoes.append(nao, ...extras, sim);
            } else {
                const ok = botaoPainel('Fechar', 'btn-primary');
                ok.onclick = () => fechar('fechar');
                acoes.append(...extras, ok);
            }
            modal.classList.remove('hidden');
        });
    }

    /** Painel → ajuste/IA → revalida → painel de novo, até o usuário decidir. Devolve 'continuar' | 'corrigir' | 'fechar'. */
    async function cicloValidacao(res, confirmar) {
        let atual = res, anterior = null;
        for (;;) {
            const escolha = await mostrarPainelValidacao(atual, confirmar, anterior);
            if (typeof escolha === 'string') return escolha;
            let aplicou = false;
            if (escolha.tipo === 'ajustar') {
                aplicou = !!((await fluxoAjuste(escolha.ids)) || {}).aplicadas;
            } else if (escolha.tipo === 'ia') {
                const semAjuste = atual.achados.filter(a => achadoCorrigivel(a) && !ajustesDoAchado(atual, a).length);
                const r = await fluxoAjusteIA(semAjuste);
                if (r && r.cadastrar) return 'fechar';   // foi para o editor de ajustes: encerra o ciclo
                aplicou = !!(r && r.aplicadas);
            }
            if (!aplicou) continue;   // cancelou: volta ao painel, nada mudou
            anterior = contagemPorSeveridade(atual.achados);
            const nova = await executarValidacao();   // revalidação automática depois de aplicar
            if (nova.nada || nova.falhou) { showToast(nova.falhou || 'Ajuste aplicado. Valide de novo.'); return 'fechar'; }
            atual = nova;
            if (confirmar && !atual.achados.some(a => a.severidade === 'erro' || a.severidade === 'aviso')) {
                showToast('Ajuste aplicado e validação limpa; seguindo com o orçamento.');
                return 'continuar';
            }
        }
    }

    /** Checagem mínima e SEMPRE ativa (Camada 1 + ativo não encontrado) antes de Ajustar/Montar
     * Orçamento (TASK-051) — independente da preferência "Validar ao montar orçamento", que
     * continua controlando só a validação completa (regras de domínio/IA) em `validarAntesDeMontar`.
     * "Bloqueia mas permite continuar mesmo assim": reaproveita o mesmo painel/fluxo de sempre. */
    async function checarInconsistenciasBasicas() {
        const res = await executarValidacao(true);
        if (res.nada || res.falhou) return true;   // validação indisponível: não trava a ação
        if (!res.achados.some(a => a.severidade === 'erro' || a.severidade === 'aviso')) return true;
        const escolha = await cicloValidacao(res, true);
        return escolha === 'continuar';
    }

    /** Devolve true se o orçamento deve seguir. Só erro/aviso interrompem; info sozinho não pergunta. */
    async function validarAntesDeMontar() {
        const res = await executarValidacao();
        if (res.nada) return true;  // nenhum modo ligado: segue sem validar
        if (res.falhou) {
            showToast('Validação indisponível; seguindo sem validar.');  // a validação é um auxílio: não trava o orçamento
            return true;
        }
        if (!res.achados.some(a => a.severidade === 'erro' || a.severidade === 'aviso')) return true;
        const escolha = await cicloValidacao(res, true);   // 'corrigir' = o orçamento fica parado; ajuste e clique de novo
        return escolha === 'continuar';
    }

    const btnValidar = document.getElementById('btn-validar');
    if (btnValidar) {
        btnValidar.addEventListener('click', async () => {
            const textoOriginal = btnValidar.textContent.trim();
            btnValidar.disabled = true;
            btnValidar.textContent = 'Validando...';
            try {
                const res = await executarValidacao();
                if (res.nada) { showToast('Nada a validar: ligue Determinística ou IA.'); return; }
                if (res.falhou) { showToast(res.falhou); return; }
                await cicloValidacao(res, false);
            } finally {
                btnValidar.disabled = false;
                btnValidar.textContent = textoOriginal;
            }
        });
    }

    /* ═══════════════════════════════════════
       CORREÇÃO ASSISTIDA (TASK-015)
       A IA só PROPÕE (POST /api/validacao/corrigir); o usuário aceita linha a linha. Aplicar entra no histórico
       como um único passo (um Ctrl+Z desfaz o lote inteiro). A operação da linha nunca é alterada.
    ═══════════════════════════════════════ */
    function achadoCorrigivel(achado) {
        return /^(CABOS|OUTROS)-\d+$/.test(String(achado.linha_id));
    }

    function localizarLinha(linhaId) {
        const m = /^(CABOS|OUTROS)-(\d+)$/.exec(linhaId);
        if (!m) return null;
        const tabela = m[1].toLowerCase();
        const row = tableStates[tabela].data[Number(m[2])];
        return row ? { tabela, row } : null;
    }

    /** Aplica as correções aceitas. Pula a linha que mudou desde a proposta (o "antes" não confere mais). */
    function aplicarCorrecoes(aceitas) {
        const validas = aceitas.filter(c => {
            const alvo = localizarLinha(c.linha_id);
            return alvo && (alvo.row.ativo || '').trim() === c.antes.trim();
        });
        if (validas.length) {
            pushHistory();  // um único passo para o lote inteiro (Ctrl+Z desfaz tudo)
            const tabelas = new Set();
            validas.forEach(c => {
                const { tabela, row } = localizarLinha(c.linha_id);
                row.ativo = c.depois;
                if (row.entidade === '0') row.entidade = autoClassifyEntidade(c.depois);  // mesma reclassificação da edição manual
                tabelas.add(tabela);
            });
            if (tabelas.has('cabos')) recalcAllQtdAtivos();
            tabelas.forEach(t => { renderTable(t); refreshAllFilters(t); });
            buildAtivoSets();
            buildDataLists();
            atualizarResumoRedeUI();
        }
        return { aplicadas: validas.length, ignoradas: aceitas.length - validas.length };
    }

    function linhaAntesDepois(c) {
        const wrap = document.createElement('div');
        wrap.style.cssText = 'font-family: monospace; font-size: 0.8rem; margin-top: 4px;';
        const antes = document.createElement('div');
        antes.style.cssText = 'color:#f85149; text-decoration: line-through;';
        antes.textContent = c.antes;
        const depois = document.createElement('div');
        depois.style.color = '#3fb950';
        depois.textContent = c.depois;
        wrap.append(antes, depois);
        return wrap;
    }

    /** Mostra as propostas e resolve com { aplicadas, ignoradas } ou null (cancelou / nada a aplicar). */
    function mostrarPainelCorrecao(resultado) {
        return new Promise(resolve => {
            const modal = document.getElementById('modal-correcao');
            const lista = document.getElementById('correcao-lista');
            const acoes = document.getElementById('correcao-acoes');
            const resumo = document.getElementById('correcao-resumo');
            const caixas = [];
            lista.replaceChildren();
            acoes.replaceChildren();
            const fechar = valor => { modal.classList.add('hidden'); resolve(valor); };

            const problema = { sem_chave: 'Sem chave de IA: informe a sua na configuração de IA.', desativado: 'A correção por IA está desativada.',
                indisponivel: 'A correção por IA está indisponível.', erro: 'A correção por IA falhou.' }[resultado.status];
            if (problema || !resultado.correcoes.length) {
                resumo.textContent = (problema || 'A IA não encontrou correção segura para estes problemas.')
                    + (resultado.mensagem && problema ? ` (${resultado.mensagem})` : '');
            } else {
                resumo.textContent = `${resultado.correcoes.length} proposta(s). Nada é alterado sem o seu aceite; aplicar pode ser desfeito com Ctrl+Z.`
                    + (resultado.status === 'parcial' ? ' Parte das linhas falhou na IA.' : '');
            }

            resultado.correcoes.forEach(c => {
                const item = document.createElement('label');
                item.style.cssText = 'display:block; padding: 8px 10px; margin-bottom: 6px; background: rgba(255,255,255,0.03); border-radius: 4px; cursor: pointer;';
                const topo = document.createElement('div');
                topo.style.cssText = 'display:flex; gap:8px; align-items:center; font-size:0.75rem; color:#8b949e;';
                const caixa = document.createElement('input');
                caixa.type = 'checkbox';
                caixa.addEventListener('change', atualizarBotoes);
                caixas.push({ caixa, c });
                const onde = document.createElement('strong');
                onde.style.color = '#c9d1d9';
                onde.textContent = c.linha_id;
                topo.append(caixa, onde);
                item.append(topo, linhaAntesDepois(c));
                if (c.motivo) {
                    const m = document.createElement('div');
                    m.style.cssText = 'margin-top:4px; font-size:0.8rem; color:#8b949e;';
                    m.textContent = c.motivo;
                    item.appendChild(m);
                }
                (c.avisos || []).forEach(a => {
                    const w = document.createElement('div');
                    w.style.cssText = 'margin-top:2px; font-size:0.75rem; color:#d29922;';
                    w.textContent = `Atenção: ${a}`;
                    item.appendChild(w);
                });
                lista.appendChild(item);
            });

            if ((resultado.descartadas || []).length) {
                const det = document.createElement('details');
                det.style.cssText = 'margin-top: 8px; font-size: 0.75rem; color: #8b949e;';
                const sum = document.createElement('summary');
                sum.textContent = `${resultado.descartadas.length} proposta(s) da IA descartada(s) por não passar na validação`;
                det.appendChild(sum);
                resultado.descartadas.forEach(d => {
                    const p = document.createElement('div');
                    p.textContent = `${d.linha_id || '?'}: ${d.motivo}`;
                    det.appendChild(p);
                });
                lista.appendChild(det);
            }

            const cancelar = botaoPainel(resultado.correcoes.length ? 'Cancelar' : 'Fechar', 'btn-secondary');
            cancelar.onclick = () => fechar(null);
            const todas = botaoPainel('Aceitar todas', 'btn-secondary');
            todas.onclick = () => { caixas.forEach(x => { x.caixa.checked = true; }); atualizarBotoes(); };
            const aplicar = botaoPainel('Aplicar selecionadas', 'btn-primary');
            aplicar.onclick = () => {
                const r = aplicarCorrecoes(caixas.filter(x => x.caixa.checked).map(x => x.c));
                fechar(r);
            };
            function atualizarBotoes() {
                aplicar.disabled = !caixas.some(x => x.caixa.checked);
                aplicar.style.opacity = aplicar.disabled ? '0.45' : '1';
                aplicar.style.cursor = aplicar.disabled ? 'not-allowed' : 'pointer';
            }
            if (resultado.correcoes.length) acoes.append(cancelar, todas, aplicar); else acoes.append(cancelar);
            atualizarBotoes();
            modal.classList.remove('hidden');
        });
    }

    /** Pede as propostas à IA só para as linhas editáveis citadas nos achados e abre o painel de aceite. */
    async function fluxoCorrecao(achados) {
        const editaveis = achados.filter(achadoCorrigivel);
        if (!editaveis.length) {
            showToast('Nenhum problema aponta uma linha editável; ajuste as tabelas manualmente.');
            return false;
        }
        const selectProj = document.getElementById('select-projeto');
        const projetoCodigo = selectProj && selectProj.selectedIndex >= 0 ? (selectProj.options[selectProj.selectedIndex].dataset.codigo || '') : '';
        showToast('Pedindo correções à IA...');
        let resultado;
        try {
            const resp = await fetch('/api/validacao/corrigir', {
                method: 'POST', headers: cabecalhoIA(),
                body: JSON.stringify({
                    cabos: linhasParaValidacao('cabos', 'CABOS'), outros: linhasParaValidacao('outros', 'OUTROS'),
                    projeto_codigo: projetoCodigo || null,
                    achados: editaveis.map(a => ({ linha_id: a.linha_id, regra_id: a.regra_id, mensagem: a.mensagem, sugestao: a.sugestao || null }))
                })
            });
            resultado = resp.ok ? await resp.json()
                : { status: 'erro', mensagem: await lerDetalhe(resp), correcoes: [], descartadas: [] };
        } catch (e) {
            resultado = { status: 'erro', mensagem: e.message, correcoes: [], descartadas: [] };
        }
        const aplicado = await mostrarPainelCorrecao(resultado);
        if (aplicado) {
            showToast(aplicado.aplicadas
                ? `${aplicado.aplicadas} correção(ões) aplicada(s)${aplicado.ignoradas ? `; ${aplicado.ignoradas} ignorada(s) porque a linha mudou` : ''}. Ctrl+Z desfaz.`
                : 'Nenhuma correção aplicada: as linhas mudaram desde a proposta.');
        }
        return !!(aplicado && aplicado.aplicadas > 0);   // true = mudou as tabelas (o chamador revalida)
    }

    /* ═══════════════════════════════════════
       AJUSTES (TASK-024, ADR-006)
       Receitas cadastradas (TASK-023) rodam no backend (POST /api/validacao/ajustes/preview), que só devolve o DIFF.
       Aqui o usuário vê as mudanças (editar / inserir / excluir / reordenar), escolhe e aplica — sempre com
       pré-visualização, nada automático. Aplicar é UM passo de histórico (Ctrl+Z desfaz o lote). Mudança cuja linha
       foi alterada desde a pré-visualização é ignorada.
    ═══════════════════════════════════════ */
    const PREFIXO_TABELA = { cabos: 'CABOS', outros: 'OUTROS' };

    function textoOperacao(o) {
        const linha = x => `${x.operacao || '·'}  ${x.ativo || '(vazio)'}`;
        if (o.op === 'editar') return { antes: linha(o.antes), depois: linha(o.depois), onde: o.linha_id };
        if (o.op === 'inserir') return { antes: null, depois: linha(o.depois), onde: `Nova linha em ${o.tabela === 'cabos' ? 'Cabos' : 'Outros'}` };
        if (o.op === 'excluir') return { antes: linha(o.antes), depois: null, onde: o.linha_id };
        return { antes: null, depois: `${o.ordem_depois.length} linha(s) seriam reordenadas`, onde: o.tabela === 'cabos' ? 'Cabos' : 'Outros' };
    }

    /** Aplica as mudanças escolhidas (índices de r.operacoes) nas tabelas, num único passo de histórico. */
    function aplicarOperacoesAjuste(r, escolhidas) {
        const ops = r.operacoes.filter((o, i) => escolhidas.has(i));
        if (!ops.length) return { aplicadas: 0, ignoradas: 0 };
        pushHistory();   // um passo para o lote inteiro
        let aplicadas = 0, ignoradas = 0;
        const tocadas = new Set();
        ['cabos', 'outros'].forEach(tabela => {
            const doTabela = ops.filter(o => o.tabela === tabela);
            if (!doTabela.length) return;
            let lista = tableStates[tabela].data.map((row, i) => ({ id: `${PREFIXO_TABELA[tabela]}-${i}`, row })).filter(x => x.row);
            const achar = id => lista.find(x => x.id === id);
            const confere = (row, antes) => (row.ativo || '').trim() === (antes.ativo || '').trim() && (row.operacao || '') === antes.operacao;
            doTabela.filter(o => o.op === 'editar').forEach(o => {
                const x = achar(o.linha_id);
                if (!x || !confere(x.row, o.antes)) { ignoradas++; return; }
                x.row.ativo = o.depois.ativo;
                x.row.operacao = o.depois.operacao;
                if (x.row.entidade === '0') x.row.entidade = autoClassifyEntidade(o.depois.ativo);
                aplicadas++; tocadas.add(tabela);
            });
            doTabela.filter(o => o.op === 'excluir').forEach(o => {
                const x = achar(o.linha_id);
                if (!x || !confere(x.row, o.antes)) { ignoradas++; return; }
                lista = lista.filter(y => y !== x);
                aplicadas++; tocadas.add(tabela);
            });
            doTabela.filter(o => o.op === 'inserir').forEach(o => {
                const ent = autoClassifyEntidade(o.depois.ativo);
                const nova = { entidade: ent !== '0' ? ent : (tabela === 'cabos' ? 'CABO' : (o.depois.entidade || '0')), operacao: o.depois.operacao || 'I', ativo: o.depois.ativo };
                const ref = o.depois_de ? lista.findIndex(y => y.id === o.depois_de) : -1;
                lista.splice(o.depois_de && ref >= 0 ? ref + 1 : (o.depois_de ? lista.length : 0), 0, { id: o.linha_id, row: nova });
                aplicadas++; tocadas.add(tabela);
            });
            const mover = doTabela.find(o => o.op === 'mover');
            if (mover) {
                const pos = new Map(mover.ordem_depois.map((id, i) => [id, i]));
                let ultima = -1;
                const chaves = lista.map(x => { if (pos.has(x.id)) ultima = pos.get(x.id); return ultima + (pos.has(x.id) ? 0 : 0.5); });
                lista = lista.map((x, i) => ({ x, k: chaves[i], i })).sort((a, b) => (a.k - b.k) || (a.i - b.i)).map(y => y.x);
                aplicadas++; tocadas.add(tabela);
            }
            tableStates[tabela].data = lista.map(x => x.row);
        });
        if (tocadas.has('cabos')) recalcAllQtdAtivos();
        tocadas.forEach(t => { renderTable(t); refreshAllFilters(t); });
        buildAtivoSets();
        buildDataLists();
        atualizarResumoRedeUI();
        return { aplicadas, ignoradas };
    }

    /** Mostra o diff e deixa escolher; resolve { aplicadas, ignoradas } ou null (cancelou / nada a aplicar). */
    function mostrarPainelAjuste(r, nomes, extra) {
        return new Promise(resolve => {
            const modal = document.getElementById('modal-ajuste');
            const lista = document.getElementById('ajuste-lista');
            const acoes = document.getElementById('ajuste-acoes');
            const resumo = document.getElementById('ajuste-resumo');
            const fechar = v => { modal.classList.add('hidden'); resolve(v); };
            const caixas = [];
            lista.replaceChildren();
            acoes.replaceChildren();
            document.getElementById('ajuste-titulo').textContent = extra ? `Ajuste proposto pela IA: ${nomes}` : `Pré-visualização do ajuste: ${nomes}`;
            const n = r.resumo;
            const total = r.operacoes.length;
            resumo.textContent = total
                ? `${n.editar} editada(s), ${n.inserir} inserida(s), ${n.excluir} excluída(s)${n.mover ? ', ordem alterada' : ''}. Nada é alterado sem o seu aceite; aplicar pode ser desfeito com Ctrl+Z.`
                : 'Este ajuste não mudaria nada nas tabelas atuais.';

            if (extra) {   // proposta da IA: o que ela quer fazer (em português), aviso de exclusão e motivo
                if (extra.destrutivo) {
                    const av = document.createElement('div');
                    av.style.cssText = 'padding:8px 10px; margin-bottom:8px; border:1px solid #f85149; border-radius:4px; color:#ffa198; font-size:0.82rem;';
                    av.textContent = '⚠ Esta proposta da IA EXCLUI linhas ou itens das tabelas. Confira cada mudança abaixo (as exclusões estão em vermelho) e desmarque o que não quiser. Ctrl+Z desfaz.';
                    lista.appendChild(av);
                }
                const bloco = document.createElement('div');
                bloco.style.cssText = 'margin-bottom:8px; font-size:0.8rem; color:#8b949e;';
                extra.acoes.forEach(a => {
                    const l = document.createElement('div');
                    l.textContent = `• ${a.frase}${a.motivo ? ` — ${a.motivo}` : ''}`;
                    bloco.appendChild(l);
                });
                (extra.descartadas || []).forEach(d => {
                    const l = document.createElement('div');
                    l.style.color = '#d29922';
                    l.textContent = `• Proposta descartada${d.acao ? ` (${d.acao})` : ''}: ${d.motivo}`;
                    bloco.appendChild(l);
                });
                lista.appendChild(bloco);
            }
            r.operacoes.forEach((o, i) => {
                const t = textoOperacao(o);
                const item = document.createElement('label');
                item.style.cssText = 'display:block; padding: 8px 10px; margin-bottom: 6px; background: rgba(255,255,255,0.03); border-radius: 4px; cursor: pointer;'
                    + (o.op === 'excluir' ? ' border-left: 3px solid #f85149;' : '');
                const topo = document.createElement('div');
                topo.style.cssText = 'display:flex; gap:8px; align-items:center; font-size:0.75rem; color:#8b949e;';
                const caixa = document.createElement('input');
                caixa.type = 'checkbox'; caixa.checked = true;
                caixa.addEventListener('change', atualizar);
                caixas.push({ caixa, i });
                const onde = document.createElement('strong');
                onde.style.color = '#c9d1d9';
                onde.textContent = `${{ editar: 'Editar', inserir: 'Inserir', excluir: 'EXCLUIR', mover: 'Reordenar' }[o.op]} · ${t.onde}`;
                topo.append(caixa, onde);
                item.appendChild(topo);
                const wrap = document.createElement('div');
                wrap.style.cssText = 'font-family: monospace; font-size: 0.8rem; margin-top: 4px;';
                if (t.antes) { const d = document.createElement('div'); d.style.cssText = 'color:#f85149; text-decoration: line-through;'; d.textContent = t.antes; wrap.appendChild(d); }
                if (t.depois) { const d = document.createElement('div'); d.style.color = o.op === 'mover' ? '#8b949e' : '#3fb950'; d.textContent = t.depois; wrap.appendChild(d); }
                item.appendChild(wrap);
                lista.appendChild(item);
            });
            const notas = [...r.descartadas.map(d => `Descartado (${d.linha_id}): ${d.motivo}`), ...r.avisos.map(a => `Atenção: ${a}`)];
            if (notas.length) {
                const det = document.createElement('div');
                det.style.cssText = 'margin-top: 8px; font-size: 0.78rem; color: #d29922;';
                notas.forEach(x => { const p = document.createElement('div'); p.textContent = x; det.appendChild(p); });
                lista.appendChild(det);
            }

            const cancelar = botaoPainel(total ? 'Cancelar' : 'Fechar', 'btn-secondary');
            cancelar.onclick = () => fechar(null);
            const aplicar = botaoPainel('Aplicar selecionadas', 'btn-primary');
            aplicar.onclick = () => {
                const marcadas = new Set(caixas.filter(x => x.caixa.checked).map(x => x.i));
                const apagaAlgo = extra && [...marcadas].some(i => r.operacoes[i].op === 'excluir' || (r.operacoes[i].op === 'editar'
                    && extra.acoes.some(a => a.destrutiva && r.operacoes[i].acoes.includes(a.indice))));
                if (apagaAlgo && !confirm('Esta proposta da IA EXCLUI linhas ou itens. Você conferiu a pré-visualização e quer aplicar mesmo assim?\n(Ctrl+Z desfaz.)')) return;
                const res = aplicarOperacoesAjuste(r, marcadas);
                showToast(res.aplicadas
                    ? `${res.aplicadas} mudança(s) aplicada(s)${res.ignoradas ? `; ${res.ignoradas} ignorada(s) porque a linha mudou` : ''}. Ctrl+Z desfaz.`
                    : 'Nada aplicado: as linhas mudaram desde a pré-visualização.');
                fechar(res.aplicadas ? res : null);
            };
            function atualizar() {
                const marcadas = caixas.filter(x => x.caixa.checked).length;
                aplicar.disabled = !marcadas;
                aplicar.style.opacity = marcadas ? '1' : '0.45';
                aplicar.textContent = `Aplicar selecionadas (${marcadas})`;
            }
            if (extra && extra.acoes.length && localStorage.getItem('is_admin') === 'true' && typeof window.rpCadastrarComoAjuste === 'function') {
                const cad = botaoPainel('Cadastrar como ajuste', 'btn-secondary');
                cad.title = 'Abre o editor de ajustes (Regras de validação › Ajustes) com estas ações, para salvar como receita reutilizável';
                cad.onclick = () => { fechar({ cadastrar: true }); window.rpCadastrarComoAjuste(extra.acoes.map(a => a.acao), `Proposto pela IA: ${extra.acoes[0].frase.replace(/^Em [^:]+: /, '').slice(0, 70)}`); };
                acoes.appendChild(cad);
            }
            if (total) acoes.append(cancelar, aplicar); else acoes.append(cancelar);
            atualizar();
            modal.classList.remove('hidden');
        });
    }

    /** Pré-visualiza as receitas `ids` nas tabelas atuais e, se o usuário aceitar, aplica. Devolve o resultado ou null. */
    async function fluxoAjuste(ids) {
        let r;
        try {
            const resp = await fetch('/api/validacao/ajustes/preview', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ receitas: ids, cabos: linhasParaValidacao('cabos', 'CABOS'), outros: linhasParaValidacao('outros', 'OUTROS'),
                                       projeto_codigo: codigoDoProjetoSelecionado() || null,
                                       contexto: window.contextoSelecionado ? window.contextoSelecionado() : null })
            });
            if (!resp.ok) throw new Error(await lerDetalhe(resp));
            r = await resp.json();
        } catch (e) {
            showToast(`Não foi possível pré-visualizar o ajuste (${e.message}).`);
            return null;
        }
        const nomes = ids.length === 1 ? (nomesDeAjuste[ids[0]] || ids[0]) : `${ids.length} ajustes`;
        return mostrarPainelAjuste(r, nomes);
    }

    /** "Pedir ajuste à IA": a IA propõe AÇÕES (inclusive excluir, com aviso); o motor calcula o diff; o usuário confere no
     *  popup e aplica. Devolve o resultado do aplicar ou null. Só vai à IA o que os achados citam. */
    async function fluxoAjusteIA(achados) {
        const editaveis = achados.filter(achadoCorrigivel);
        if (!editaveis.length) { showToast('Nenhum problema aponta uma linha editável; ajuste as tabelas manualmente.'); return null; }
        showToast('Pedindo ajustes à IA...');
        let r;
        try {
            const resp = await fetch('/api/validacao/ajustes-ia', {
                method: 'POST', headers: cabecalhoIA(),
                body: JSON.stringify({
                    cabos: linhasParaValidacao('cabos', 'CABOS'), outros: linhasParaValidacao('outros', 'OUTROS'),
                    projeto_codigo: codigoDoProjetoSelecionado() || null,
                    achados: editaveis.map(a => ({ linha_id: a.linha_id, regra_id: a.regra_id, mensagem: a.mensagem, sugestao: a.sugestao || null }))
                })
            });
            r = resp.ok ? await resp.json() : { status: 'erro', mensagem: await lerDetalhe(resp), acoes: [], diff: null };
        } catch (e) {
            r = { status: 'erro', mensagem: e.message, acoes: [], diff: null };
        }
        const problema = { sem_chave: 'Sem chave de IA: informe a sua na configuração de IA.', desativado: 'O ajuste por IA está desativado.',
            indisponivel: 'O ajuste por IA está indisponível.', erro: 'O ajuste por IA falhou.' }[r.status];
        if (problema) { showToast(problema + (r.mensagem && r.status !== 'desativado' ? ` (${r.mensagem})` : '')); return null; }
        if (!r.diff) {
            showToast(r.descartadas && r.descartadas.length
                ? `A IA não propôs nenhum ajuste válido (${r.descartadas.length} descartado(s)).` : 'A IA não encontrou ajuste seguro para estes problemas.');
            return null;
        }
        return mostrarPainelAjuste(r.diff, `${r.acoes.length} ação(ões)`, { acoes: r.acoes, destrutivo: r.destrutivo, descartadas: r.descartadas });
    }

    /* ═══════════════════════════════════════
       BOTÃO "AJUSTAR" (TASK-029): roda TODOS os ajustes habilitados do projeto (POST /api/validacao/ajustes/preview-lote,
       uma única chamada) e mostra um cartão por ajuste que muda algo, com "Executar correção"; acima, "Executar todas as
       correções" (em cadeia, na ordem do cadastro, num só passo de histórico). Sempre com pré-visualização; depois de
       executar, tudo é recalculado. Exclusões pedem confirmação extra.
    ═══════════════════════════════════════ */
    async function carregarLoteDeAjustes() {
        try {
            const resp = await fetch('/api/validacao/ajustes/preview-lote', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ cabos: linhasParaValidacao('cabos', 'CABOS'), outros: linhasParaValidacao('outros', 'OUTROS'),
                                       projeto_codigo: codigoDoProjetoSelecionado() || null,
                                       contexto: window.contextoSelecionado ? window.contextoSelecionado() : null })
            });
            if (!resp.ok) throw new Error(await lerDetalhe(resp));
            return await resp.json();
        } catch (e) {
            showToast(`Não foi possível calcular os ajustes (${e.message}).`);
            return null;
        }
    }

    function resumoDoDiff(n) {
        const partes = [];
        if (n.editar) partes.push(`${n.editar} editada(s)`);
        if (n.inserir) partes.push(`${n.inserir} inserida(s)`);
        if (n.excluir) partes.push(`${n.excluir} excluída(s)`);
        if (n.mover) partes.push('ordem alterada');
        return partes.join(', ');
    }

    function mensagemDeAplicacao(res, o) {
        return `${o}: ${res.aplicadas} mudança(s) aplicada(s)${res.ignoradas ? `; ${res.ignoradas} ignorada(s) porque a linha mudou` : ''}. Ctrl+Z desfaz.`;
    }

    async function abrirPainelAjustar(mensagem) {
        const r = await carregarLoteDeAjustes();
        if (!r) return;
        const modal = document.getElementById('modal-ajustar');
        const lista = document.getElementById('ajustar-lista');
        const topo = document.getElementById('ajustar-topo');
        const resumo = document.getElementById('ajustar-resumo');
        const aviso = document.getElementById('ajustar-aviso');
        lista.replaceChildren();
        topo.replaceChildren();
        aviso.textContent = mensagem || '';
        aviso.style.display = mensagem ? 'block' : 'none';
        const fechar = () => modal.classList.add('hidden');
        const totalMudancas = r.itens.reduce((t, i) => t + Object.values(i.diff.resumo).reduce((a, b) => a + b, 0), 0);

        if (!r.total_habilitados) {
            resumo.textContent = 'Nenhum ajuste habilitado neste projeto. O administrador liga os ajustes em Regras de validação › Ajustes.';
        } else if (!r.itens.length) {
            resumo.textContent = `Nada a ajustar nas tabelas atuais (${r.total_habilitados} ajuste(s) habilitado(s), nenhum mudaria algo).`;
        } else {
            resumo.textContent = `${r.itens.length} ajuste(s) com correções (${totalMudancas} mudança(s))`
                + (r.sem_mudanca.length ? `; ${r.sem_mudanca.length} habilitado(s) não mudaria(m) nada` : '') + '. Nada é alterado sem você executar.';
            const todas = botaoPainel('Executar todas as correções', 'btn-primary');
            todas.style.cssText = 'width:100%; margin-bottom:10px;';
            if (!r.cadeia) {
                todas.disabled = true; todas.style.opacity = '0.5';
                todas.title = r.cadeia_erro || 'Indisponível: execute as correções uma a uma.';
            } else {
                todas.title = 'Executa todos os ajustes em sequência, na ordem do cadastro, num único passo de histórico';
                todas.onclick = async () => {
                    const n = r.cadeia.resumo;
                    if (r.cadeia.destrutivo && !confirm(`Executar todas as correções inclui EXCLUSÕES (${n.excluir} linha(s)/itens removidos). Você conferiu a pré-visualização? (Ctrl+Z desfaz.)`)) return;
                    const res = aplicarOperacoesAjuste(r.cadeia, new Set(r.cadeia.operacoes.map((o, i) => i)));
                    showToast(mensagemDeAplicacao(res, 'Todas as correções'));
                    await abrirPainelAjustar(mensagemDeAplicacao(res, 'Todas as correções'));
                };
            }
            topo.appendChild(todas);
        }
        (r.avisos || []).forEach(a => { const d = document.createElement('div'); d.style.cssText = 'font-size:0.78rem; color:#d29922;'; d.textContent = a; lista.appendChild(d); });
        r.problemas.forEach(p => {
            const d = document.createElement('div');
            d.style.cssText = 'font-size:0.78rem; color:#d29922; margin-bottom:6px;';
            d.textContent = `Ajuste "${p.nome}" ignorado (configuração inválida): ${p.erros.join('; ')}`;
            lista.appendChild(d);
        });

        r.itens.forEach(item => {
            const card = document.createElement('div');
            card.style.cssText = 'padding:10px 12px; margin-bottom:8px; background:rgba(255,255,255,0.03); border-radius:6px;'
                + (item.destrutivo ? ' border-left:3px solid #f85149;' : ' border-left:3px solid #3fb950;');
            const cab = document.createElement('div');
            cab.style.cssText = 'display:flex; gap:8px; align-items:center; flex-wrap:wrap;';
            const nome = document.createElement('strong');
            nome.textContent = item.nome;
            const sub = document.createElement('span');
            sub.style.cssText = 'font-size:0.78rem; color:#8b949e;';
            sub.textContent = resumoDoDiff(item.diff.resumo);
            cab.append(nome, sub);
            if (item.destrutivo) {
                const b = document.createElement('span');
                b.style.cssText = 'font-size:0.7rem; color:#f85149; border:1px solid #f85149; border-radius:10px; padding:0 6px;';
                b.textContent = 'EXCLUI';
                cab.appendChild(b);
            }
            const exec = botaoPainel('Executar correção', 'btn-secondary');
            exec.style.cssText = 'margin-left:auto; padding:4px 12px; font-size:0.8rem; color:#3fb950; border-color:rgba(63,185,80,0.5);';
            cab.appendChild(exec);
            card.appendChild(cab);
            item.frases.forEach(f => { const d = document.createElement('div'); d.style.cssText = 'font-size:0.78rem; color:#8b949e; margin-top:2px;'; d.textContent = f; card.appendChild(d); });

            const det = document.createElement('details');
            det.style.marginTop = '6px';
            const sum = document.createElement('summary');
            sum.style.cssText = 'font-size:0.78rem; color:#58a6ff; cursor:pointer;';
            sum.textContent = 'Ver e escolher as mudanças';
            det.appendChild(sum);
            const caixas = [];
            item.diff.operacoes.slice(0, 60).forEach((o, i) => {
                const t = textoOperacao(o);
                const linha = document.createElement('label');
                linha.style.cssText = 'display:block; padding:4px 6px; margin-top:4px; cursor:pointer; font-size:0.78rem;' + (o.op === 'excluir' ? ' border-left:2px solid #f85149;' : '');
                const cx = document.createElement('input');
                cx.type = 'checkbox'; cx.checked = true; cx.style.marginRight = '6px';
                caixas.push({ cx, i });
                const onde = document.createElement('span');
                onde.style.color = '#c9d1d9';
                onde.textContent = `${{ editar: 'Editar', inserir: 'Inserir', excluir: 'EXCLUIR', mover: 'Reordenar' }[o.op]} · ${t.onde}`;
                linha.append(cx, onde);
                const w = document.createElement('div');
                w.style.cssText = 'font-family:monospace; margin-left:22px;';
                if (t.antes) { const d = document.createElement('div'); d.style.cssText = 'color:#f85149; text-decoration:line-through;'; d.textContent = t.antes; w.appendChild(d); }
                if (t.depois) { const d = document.createElement('div'); d.style.color = o.op === 'mover' ? '#8b949e' : '#3fb950'; d.textContent = t.depois; w.appendChild(d); }
                linha.appendChild(w);
                det.appendChild(linha);
            });
            if (item.diff.operacoes.length > 60) {
                const mais = document.createElement('div');
                mais.style.cssText = 'font-size:0.75rem; color:#8b949e; margin-top:4px;';
                mais.textContent = `…e mais ${item.diff.operacoes.length - 60} mudança(s) (todas entram ao executar a correção).`;
                det.appendChild(mais);
            }
            card.appendChild(det);
            [...item.diff.descartadas.map(d => `Descartado (${d.linha_id}): ${d.motivo}`), ...item.diff.avisos.map(a => `Atenção: ${a}`)].forEach(x => {
                const d = document.createElement('div'); d.style.cssText = 'font-size:0.75rem; color:#d29922; margin-top:3px;'; d.textContent = x; card.appendChild(d);
            });

            exec.onclick = async () => {
                const todas = item.diff.operacoes.length <= 60;
                const marcadas = new Set(caixas.filter(x => x.cx.checked).map(x => x.i));
                if (!todas) item.diff.operacoes.forEach((o, i) => marcadas.add(i));   // acima do teto exibido: executa o ajuste inteiro
                if (!marcadas.size) { showToast('Nenhuma mudança marcada.'); return; }
                if (item.destrutivo && !confirm(`O ajuste "${item.nome}" EXCLUI linhas ou itens. Você conferiu a pré-visualização e quer executar? (Ctrl+Z desfaz.)`)) return;
                const res = aplicarOperacoesAjuste(item.diff, marcadas);
                const msg = res.aplicadas ? mensagemDeAplicacao(res, `Correção "${item.nome}"`) : 'Nada aplicado: as linhas mudaram desde a pré-visualização.';
                showToast(msg);
                await abrirPainelAjustar(msg);
            };
            lista.appendChild(card);
        });

        if (r.itens.length && r.sem_mudanca.length) {
            const d = document.createElement('div');
            d.style.cssText = 'font-size:0.75rem; color:#8b949e; margin-top:4px;';
            d.textContent = `Habilitados que não mudariam nada: ${r.sem_mudanca.map(x => x.nome).join(', ')}.`;
            lista.appendChild(d);
        }
        document.getElementById('ajustar-fechar').onclick = fechar;
        modal.classList.remove('hidden');
    }

    const btnAjustar = document.getElementById('btn-ajustar');
    if (btnAjustar) btnAjustar.addEventListener('click', async () => {
        // TASK-051: checagem obrigatória (contrato + ativo não encontrado) antes de abrir a gaveta.
        if (!(await checarInconsistenciasBasicas())) return;
        abrirPainelAjustar();
    });

    /** Montar Orçamento = "esta é a versão final": registra as correções da sessão para o aprendizado (como o Salvar obra; só administrador). */
    function registrarAprendizadoDoOrcamento() {
        try {
            if (localStorage.getItem('is_admin') !== 'true' || !(execucaoOrigemAutonomo || baselineManual)) return;
            const min = t => (tableStates[t].data || []).filter(Boolean).map(r => ({ entidade: r.entidade, operacao: r.operacao, ativo: r.ativo, _x: r._x, _y: r._y }));
            const corpo = { projeto: projetoAtualObras(), cabos: min('cabos'), outros: min('outros') };
            if (execucaoOrigemAutonomo) corpo.origem_execucao = execucaoOrigemAutonomo;
            else corpo.baseline_manual = { sessao: sessaoManual, ...baselineManual };
            fetch('/api/aprendizado/sessao', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(corpo) }).catch(() => {});
        } catch (e) { /* o aprendizado nunca atrapalha o orçamento */ }
    }

    const btnMontarOrcamento = document.getElementById('btn-montar-orcamento');
    const modalOrcamento = document.getElementById('modal-orcamento');
    const tbodyOrcamento = document.getElementById('resultado-orcamento-tbody');

    if (btnMontarOrcamento) {
        btnMontarOrcamento.addEventListener('click', async () => {
            // Feedback visual: desabilita botão e exibe "Calculando..."
            const textoOriginal = btnMontarOrcamento.textContent.trim();
            const revisao = revisaoObraAberta;
            btnMontarOrcamento.disabled = true;
            btnMontarOrcamento.textContent = 'Calculando...';

            try {
                // TASK-058: `V` de modelo que sobrou nas tabelas não vira orçamento
                const nV = contarVariaveisSobrando();
                if (nV > 0) {
                    alert(`Há ${nV} linha(s) com variável de modelo (V ou X(...)) sem valor. Gere a obra pelo modelo (Carregar/Adicionar/Subtrair) ou troque a variável por um valor antes de montar o orçamento.`);
                    return;
                }
                // TASK-051: checagem obrigatória (contrato + ativo não encontrado), sempre ativa,
                // independente da preferência abaixo.
                if (!(await checarInconsistenciasBasicas())) return;

                // Validação automática (opção local, desligada por padrão): se houver erro/aviso, pergunta
                // "continuar mesmo assim?" e só segue com o orçamento se o usuário confirmar.
                if (validacaoAutomaticaLigada() && !(await validarAntesDeMontar())) return;

                const selectProj = document.getElementById('select-projeto');
                const projVal = selectProj ? selectProj.value : "";
                const projCode = selectProj && selectProj.selectedIndex >= 0 ? selectProj.options[selectProj.selectedIndex].dataset.codigo : "";
                
                const { cabos: payloadCabos, outros: payloadOutros } = await obterPayloadCalculo();
                if (revisao !== revisaoObraAberta) {
                    showToast('A obra aberta mudou. Clique em Montar Orçamento para calcular a obra atual.');
                    return;
                }

                const payload = {
                    cabos: payloadCabos,
                    outros: payloadOutros,
                    projeto: projVal,
                    obra: window.ObraParams ? ObraParams.obraDoProjeto() : {}          // TASK-165: cabeçalho + contexto da obra aberta (parâmetros OBRA.*)
                };
                
                if(selectProj) {
                    localStorage.setItem('projeto_selecionado', projVal);
                    localStorage.setItem('projeto_selecionado_codigo', projCode);
                }
                
                registrarAprendizadoDoOrcamento();
                const texto = JSON.stringify(payload);
                localStorage.setItem('orcamentoPayload', texto); // abertura direta do Resultado
                window.__orcamentoDaSessao = texto;               // a janela integrada usa o pedido do seu Resumo
                AbrirNaApp.abrir('/resultado_orcamento', 'Resultado do orçamento', 'resultado');
            } finally {
                // Restaura o botão independentemente de sucesso ou erro
                btnMontarOrcamento.disabled = false;
                btnMontarOrcamento.textContent = textoOriginal;
            }
        });
    }

    /* ═══════════════════════════════════════
       LIMPAR TUDO (ao lado do Salvar no PDF)
    ═══════════════════════════════════════ */
    /* TASK-117: o Limpar Tudo também esvazia os modais que guardam dados da obra (Ramais, Conexões, geradores de Postes e Cabos).
       Vincular cabos guarda os cabos fictícios dentro das linhas de Outros (saem junto); os modais montáveis começam vazios a cada abertura. */
    const _camposConexoes = () => Array.from(document.querySelectorAll('#modal-conexoes input.con-mono, #modal-conexoes input.con-tri, #modal-conexoes input.con-glv'));
    const _chaveConexao = i => (i.className.match(/con-(mono|tri|glv)/) || [''])[0] + '|' + (i.dataset.id || '');
    function coletarModaisDaObra() {
        const conexoes = {};
        _camposConexoes().forEach(i => { if (String(i.value || '').trim() !== '') conexoes[_chaveConexao(i)] = i.value; });
        return { ramais: JSON.parse(JSON.stringify(window._ramaisData || [])), conexoes };
    }
    function contarModaisDaObra() {
        const m = coletarModaisDaObra();
        const g = (typeof GeradorPostes !== 'undefined' ? GeradorPostes.preenchidas() : 0) + (typeof GeradorCabos !== 'undefined' ? GeradorCabos.preenchidas() : 0);
        return { ramais: m.ramais.length, conexoes: Object.keys(m.conexoes).length, geradores: g };
    }
    function limparModaisDaObra() {
        window._ramaisData = [];
        if (typeof renderizarTabelaRamais === 'function') renderizarTabelaRamais([]);       // também zera os ativos de ramais e o badge
        if (typeof atualizarBadgeRamais === 'function') atualizarBadgeRamais();
        _camposConexoes().forEach(i => { i.value = ''; });
        if (typeof recalcularTotaisConexoes === 'function' && document.getElementById('conexoes-totalizador')) recalcularTotaisConexoes();
        if (typeof GeradorPostes !== 'undefined') GeradorPostes.limpar();
        if (typeof GeradorCabos !== 'undefined') GeradorCabos.limpar();
    }
    function restaurarModaisDaObra(m) {
        if (!m) return;
        window._ramaisData = JSON.parse(JSON.stringify(m.ramais || []));
        if (typeof renderizarTabelaRamais === 'function') renderizarTabelaRamais(window._ramaisData);
        if (typeof atualizarBadgeRamais === 'function') atualizarBadgeRamais();
        _camposConexoes().forEach(i => { const k = _chaveConexao(i); i.value = Object.prototype.hasOwnProperty.call(m.conexoes || {}, k) ? m.conexoes[k] : ''; });
        if (typeof recalcularTotaisConexoes === 'function' && document.getElementById('conexoes-totalizador')) recalcularTotaisConexoes();
    }
    window.limparModaisDaObra = limparModaisDaObra;

    const btnLimparTabelas = document.getElementById('btn-limpar-tabelas');
    if (btnLimparTabelas) {
        btnLimparTabelas.addEventListener('click', () => {
            const nCabos = tableStates.cabos.data.length, nOutros = tableStates.outros.data.length;
            const mod = contarModaisDaObra();
            const temModais = mod.ramais + mod.conexoes + mod.geradores > 0;
            if (nCabos + nOutros === 0 && !temModais) {
                showToast('As tabelas e os modais já estão vazios.');
                return;
            }
            const itens = [];
            if (nCabos + nOutros) itens.push(`• Tabelas Cabos (${nCabos} linha(s)) e Outros (${nOutros} linha(s)), com os cabos fictícios do Vincular cabos`);
            if (mod.ramais) itens.push(`• Modal Ramais (${mod.ramais} item(ns))`);
            if (mod.conexoes) itens.push(`• Modal Conexões (${mod.conexoes} campo(s) preenchido(s))`);
            if (mod.geradores) itens.push(`• Geradores de Postes e Cabos (${mod.geradores} linha(s) preenchida(s))`);
            const desfazer = nCabos + nOutros ? 'Esta ação pode ser desfeita com Ctrl+Z.' : 'Esta ação NÃO pode ser desfeita (as tabelas já estão vazias).';
            if (!confirm(`Tem certeza que deseja limpar TUDO?\n\n${itens.join('\n')}\n\n${desfazer}`)) return;

            const haviaTabelas = nCabos + nOutros > 0;
            if (haviaTabelas) {
                pushHistory(); // snapshot ANTES de limpar
                const atual = history[historyIdx];
                if (atual) atual.modais = coletarModaisDaObra();      // o Ctrl+Z devolve também Ramais e Conexões (fora da comparação de estados)
            }

            tableStates.cabos.data  = [];
            tableStates.outros.data = [];
            localStorage.removeItem('processar_dados');
            limparModaisDaObra();

            renderTable('cabos');
            renderTable('outros');
            buildAtivoSets();
            buildDataLists();
            refreshAllFilters('cabos');
            refreshAllFilters('outros');
            if (typeof updateCounters === 'function') updateCounters();

            showToast(haviaTabelas ? 'Tabelas e modais limpos! Use Ctrl+Z para desfazer.' : 'Modais limpos.');
        });
    }

    /* ═══════════════════════════════════════
       GERADOR DE CÓDIGOS - CABOS
    ═══════════════════════════════════════ */
    // TASK-068: o modal é montado por static/gerador_modal.js (painéis de atalhos, favoritos, linhas estilo Excel).
    const GeradorCabos = GeradorModal.criar('cabos', { overlay: 'modal-cabos-gerador', paineis: 'gm-paineis-cabos', tbody: 'body-modal-cabos' });
    window.abrirModalCabos = function() { GeradorCabos.abrir(); };
    window.adicionarLinhaCabos = function() { GeradorCabos.adicionarLinha(null); };

    window.inserirGeradorNaTabelaCabos = function() {
        const linhasModalCabos = GeradorCabos.linhas;
        if (!linhasModalCabos.length) return;
        
        pushHistory(); 
        
        const lenBefore = tableStates.cabos.data.length;
        
        linhasModalCabos.forEach(linha => {
            const ativoBase = (linha.ativo || '').trim();
            if (!ativoBase) return;
            
            const ativoFinal = `${ativoBase} ${linha.fase} ${linha.comp} m`;
            
            for (let i = 0; i < linha.qtd; i++) {
                tableStates.cabos.data.push({
                    entidade: 'CABO',
                    operacao: linha.op,
                    ativo: ativoFinal
                });
            }
        });
        
        const lenAfter = tableStates.cabos.data.length;
        for (let i = lenBefore; i < lenAfter; i++) {
            recalcRowAndAbove(i);
        }
        // Remove linhas com ativo vazio antes de renderizar
        tableStates.cabos.data = tableStates.cabos.data.filter(row => (row.ativo || '').trim() !== '');
        renderTable('cabos');
        document.getElementById('modal-cabos-gerador').classList.add('hidden');
        buildAtivoSets();
        buildDataLists();
    };

    /* ═══════════════════════════════════════
       POSTES E ESTRUTURAS — Modal Gerador
    ═══════════════════════════════════════ */
    const GeradorPostes = GeradorModal.criar('postes', { overlay: 'modal-postes-gerador', paineis: 'gm-paineis-postes', tbody: 'body-modal-postes' });
    window.abrirModalPostes = function() { GeradorPostes.abrir(); };
    window.adicionarLinhaPostes = function() { GeradorPostes.adicionarLinha(null); };

    window.inserirGeradorNaTabelaOutros = function() {
        const linhasModalPostes = GeradorPostes.linhas;
        if (!linhasModalPostes.length) return;

        pushHistory();

        linhasModalPostes.forEach(linha => {
            const ativoBase = (linha.ativo || '').trim();
            if (!ativoBase) return;

            for (let i = 0; i < linha.qtd; i++) {
                const nova = { entidade: '0', operacao: linha.op, ativo: ativoBase };
                reclassificarEntidade(nova);          // TASK-065: a entidade acompanha o ativo (antes ficava 'OUTRO')
                tableStates.outros.data.push(nova);
            }
        });

        renderTable('outros');
        document.getElementById('modal-postes-gerador').classList.add('hidden');
        buildAtivoSets();
        buildDataLists();
    };


    /* ═══════════════════════════════════════
       MODAIS MONTÁVEIS (TASK-077) — o motor/tela estão em modal_montavel_*.js; aqui só a escrita nas tabelas.
       Itens = [{ativo, qtd, operacao}]. Outros: uma linha "qtd-ativo" por item (mesmo formato de Ramais/Conexões).
       Cabos: o ativo repetido `qtd` vezes (mesmo formato do gerador de Cabos).
    ═══════════════════════════════════════ */
    window.ModalMontavelHost = {
        /** Linhas atuais da tabela como objetos { operacao, ativo, entidade } (para blocos grade com `carregar_de`). */
        ler(destino) {
            const dados = destino === 'cabos' ? tableStates.cabos.data : tableStates.outros.data;
            return dados.map(r => ({ operacao: r.operacao || '', ativo: r.ativo || '', entidade: r.entidade || '' }));
        },
        adicionar(destino, itens, nomeModal) {
            if (!itens || !itens.length) return;
            pushHistory();
            const origem = 'MODAL: ' + nomeModal;
            if (destino === 'cabos') {
                const antes = tableStates.cabos.data.length;
                itens.forEach(it => {
                    for (let i = 0; i < Math.max(1, Math.round(it.qtd)); i++) {
                        tableStates.cabos.data.push({ entidade: 'CABO', operacao: it.operacao || 'I', ativo: it.ativo, texto: origem });
                    }
                });
                for (let i = antes; i < tableStates.cabos.data.length; i++) recalcRowAndAbove(i);
                renderTable('cabos');
            } else {
                itens.forEach(it => {
                    tableStates.outros.data.push({ entidade: '0', operacao: it.operacao || 'I', ativo: it.qtd + '-' + it.ativo, texto: origem });
                });
                if (typeof recalcAllQtdAtivos === 'function') recalcAllQtdAtivos();
                renderTable('outros');
            }
            buildAtivoSets();
            buildDataLists();
            if (typeof updateCounters === 'function') updateCounters();
            if (typeof updateHistoryUI === 'function') updateHistoryUI();
            showToast(itens.length + ' item(ns) adicionados em ' + (destino === 'cabos' ? 'Cabos' : 'Outros') + '.');
        }
    };
    if (window.ModalMontavelUI) window.ModalMontavelUI.iniciar({ barra: document.getElementById('barra-modais-montaveis'), host: window.ModalMontavelHost });
    if (window.BarraModais) window.BarraModais.iniciar();      // TASK-117: aplica ocultos/ordem da barra deste projeto e monta a seção do painel Configurações

    // Mostrar badge na aba RAMAIS caso existam itens
    if (window._ramaisData && window._ramaisData.length > 0) {
        const badge = document.getElementById('badge-ramais');
        if (badge) {
            badge.style.display = 'inline-flex';
            badge.textContent = window._ramaisData.length;
        }
    }


/* ═══════════════════════════════════════════════════════
   MODAL RAMAIS — lógica 
═══════════════════════════════════════════════════════ */

window.abrirModalRamais = function() {
    const dados = window._ramaisData || [];
    renderizarTabelaRamais(dados);
    document.getElementById('modal-ramais').classList.remove('hidden');
};

function renderizarTabelaRamais(dados) {
    const tbody = document.getElementById('body-modal-ramais');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (dados.length === 0) {
        tbody.innerHTML = '<tr><td colspan="3" style="text-align:center; color:#8b949e; padding:20px;">Nenhum item de RAMAIS, IP ou Apoio relevante encontrado. Cole linhas aqui (Ctrl+V) ou use "+ Linha".</td></tr>';
        window.recalcularAtivosRamais();
        atualizarBadgeRamais();
        return;
    }

    dados.forEach((r, i) => {
        const tr = document.createElement('tr');
        tr.dataset.i = i;

        const tdTipo = document.createElement('td');
        const tipoColor = r.entidade === 'RAMAIS' ? '#58a6ff' : r.entidade === 'IP' ? '#d2991e' : '#3fb950';
        tdTipo.innerHTML = `<span style="background:${tipoColor}22; color:${tipoColor}; border:1px solid ${tipoColor}44; padding:2px 7px; border-radius:4px; font-size:0.72rem; font-weight:600;">${r.entidade}</span>`;
        tdTipo.style.cssText = 'padding:6px 8px; white-space:nowrap; cursor:pointer; user-select:none;';
        tdTipo.className = 'ram-tipo';
        tdTipo.title = 'Clique para selecionar a linha (Shift/Ctrl para várias). Ctrl+C copia, Delete remove.';

        const tdTexto = document.createElement('td');
        tdTexto.style.cssText = 'padding:4px 8px;';
        tdTexto.contentEditable = 'true';
        tdTexto.style.outline = 'none';
        tdTexto.className = 'ram-texto';
        tdTexto.textContent = typeof r._textoEditado === 'string' ? r._textoEditado : (r._textoOriginal || (r._raw && r._raw.texto) || r.texto || '');
        tdTexto.addEventListener('focus', () => tdTexto.style.background = 'rgba(88,166,255,0.06)');
        tdTexto.addEventListener('blur',  () => { tdTexto.style.background = ''; });

        const tdPag = document.createElement('td');
        tdPag.style.cssText = 'padding:6px 8px; text-align:center; color:#8b949e; font-size:0.78rem;';
        const pagVal = (r._raw && r._raw.pagina != null) ? r._raw.pagina : (r.pagina != null ? r.pagina : '-');
        tdPag.textContent = pagVal;

        tr.appendChild(tdTipo);
        tr.appendChild(tdTexto);
        tr.appendChild(tdPag);
        tbody.appendChild(tr);
    });

    window.recalcularAtivosRamais();
    atualizarBadgeRamais();
}

/* ── TASK-066: colar/copiar várias linhas, seleção, linhas novas e recálculo automático ── */
function atualizarBadgeRamais() {
    const badge = document.getElementById('badge-ramais');
    const n = (window._ramaisData || []).length;
    if (badge) { badge.style.display = n > 0 ? 'inline-flex' : 'none'; badge.textContent = n; }
}

function ramaisLinhaNova(texto) {
    return { entidade: 'RAMAIS', texto: texto, _textoOriginal: '', _textoEditado: texto, _raw: null, pagina: '-' };
}

/** Cola `texto` (várias linhas/colunas) a partir da linha `inicio`: sobrescreve e cria linhas novas (tipo RAMAIS). */
window.colarNosRamais = function(texto, inicio) {
    const G = window.GradeTabela;
    const parsed = G.parseColagem(texto, 1000);
    if (!parsed.linhas.length) return 0;
    const dados = window._ramaisData = window._ramaisData || [];
    parsed.linhas.forEach((cel, k) => {
        // com várias colunas coladas (tipo, texto, página) usa a coluna de texto; com uma só, ela é o texto
        const comTipo = cel.length >= 2 && /^(RAMAIS|IP|APOIO)$/i.test((cel[0] || '').trim());
        const t = (comTipo ? cel[1] : cel[0]).trim();
        const idx = inicio + k;
        if (idx < dados.length) dados[idx]._textoEditado = t;
        else dados.push(ramaisLinhaNova(t));
    });
    renderizarTabelaRamais(dados);
    return parsed.linhas.length;
};

window.adicionarLinhaRamais = function() {
    const dados = window._ramaisData = window._ramaisData || [];
    dados.push(ramaisLinhaNova(''));
    renderizarTabelaRamais(dados);
    const cel = document.querySelector('#body-modal-ramais tr:last-child .ram-texto');
    if (cel) cel.focus();
};

function ramaisLinhasSelecionadas() {
    return Array.from(document.querySelectorAll('#body-modal-ramais tr.ram-sel')).map(tr => parseInt(tr.dataset.i)).filter(i => !isNaN(i));
}

(function ligarModalRamais() {
    const modal = document.getElementById('modal-ramais');
    const tbody = document.getElementById('body-modal-ramais');
    if (!modal || !tbody) return;
    let ultima = null;

    // edição digitada: guarda e recalcula na hora
    tbody.addEventListener('input', e => {
        const td = e.target.closest ? e.target.closest('td.ram-texto') : null;
        if (!td) return;
        const dados = window._ramaisData || [], i = parseInt(td.parentElement.dataset.i);
        if (dados[i]) dados[i]._textoEditado = td.textContent;
        window.recalcularAtivosRamais();
    });

    // colar: sempre texto puro; várias linhas distribuem uma por linha da tabela
    tbody.addEventListener('paste', e => {
        const td = e.target.closest ? e.target.closest('td.ram-texto') : null;
        if (!td) return;
        const texto = (e.clipboardData || window.clipboardData).getData('text');
        e.preventDefault();
        if (!texto) return;
        if (window.GradeTabela.ehColagemMultipla(texto)) {
            const n = window.colarNosRamais(texto, parseInt(td.parentElement.dataset.i));
            if (window.showToast) window.showToast(`${n} linha(s) coladas. Ativos atualizados.`);
        } else {
            document.execCommand('insertText', false, texto.replace(/\r?\n/g, ' '));
        }
    });

    // Enter desce (na última linha cria outra)
    tbody.addEventListener('keydown', e => {
        const td = e.target.closest ? e.target.closest('td.ram-texto') : null;
        if (!td) return;
        if (e.key === 'Enter') {
            e.preventDefault();
            const prox = td.parentElement.nextElementSibling;
            if (prox && prox.querySelector('.ram-texto')) prox.querySelector('.ram-texto').focus();
            else window.adicionarLinhaRamais();
        }
    });

    // seleção de linhas pelo tipo
    tbody.addEventListener('mousedown', e => {
        const td = e.target.closest ? e.target.closest('td.ram-tipo') : null;
        if (!td) return;
        e.preventDefault();
        const tr = td.parentElement, i = parseInt(tr.dataset.i);
        const marcar = on => { tr.classList.toggle('ram-sel', on); tr.style.background = tr.classList.contains('ram-sel') ? 'rgba(56,139,253,0.2)' : ''; };
        if (e.shiftKey && ultima !== null) {
            Array.from(tbody.children).forEach(x => { x.classList.remove('ram-sel'); x.style.background = ''; });
            Array.from(tbody.children).forEach(x => {
                const k = parseInt(x.dataset.i);
                if (k >= Math.min(ultima, i) && k <= Math.max(ultima, i)) { x.classList.add('ram-sel'); x.style.background = 'rgba(56,139,253,0.2)'; }
            });
        } else if (e.ctrlKey || e.metaKey) { marcar(!tr.classList.contains('ram-sel')); ultima = i; }
        else {
            Array.from(tbody.children).forEach(x => { x.classList.remove('ram-sel'); x.style.background = ''; });
            marcar(true); ultima = i;
        }
        if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
    });

    // Ctrl+C copia só as linhas selecionadas; Delete remove
    const aberto = () => !modal.classList.contains('hidden');
    document.addEventListener('copy', e => {
        if (!aberto()) return;
        const sel = ramaisLinhasSelecionadas();
        if (!sel.length || (document.activeElement && document.activeElement.isContentEditable)) return;
        const dados = window._ramaisData || [];
        const linhas = sel.map(i => {
            const tr = tbody.querySelector(`tr[data-i="${i}"]`);
            const tx = tr ? tr.querySelector('.ram-texto').textContent.trim() : '';
            return `${dados[i] ? dados[i].entidade : ''}\t${tx}\t${tr ? tr.children[2].textContent : ''}`;
        });
        e.preventDefault();
        e.clipboardData.setData('text/plain', linhas.join('\n'));
        if (window.showToast) window.showToast(`${linhas.length} linha(s) copiadas.`);
    });
    document.addEventListener('keydown', e => {
        if (aberto() && (e.key === 'Delete' || e.key === 'Backspace') && !(document.activeElement && document.activeElement.isContentEditable)) {
            const sel = ramaisLinhasSelecionadas();
            if (!sel.length) return;
            e.preventDefault();
            window._ramaisData = (window._ramaisData || []).filter((_, i) => !sel.includes(i));
            renderizarTabelaRamais(window._ramaisData);
        }
    });
})();

window.recalcularAtivosRamais = function() {
    const isCorrosao = document.querySelector('input[name="ramais-corrosao"][value="sim"]')?.checked;
    const rows = document.querySelectorAll('#body-modal-ramais tr');
    
    window._ramaisAtivosInstalando = {};
    window._ramaisAtivosRemovendo = {};

    function addAtivo(dict, ativo, qty) {
        if (!dict[ativo]) dict[ativo] = 0;
        dict[ativo] += qty;
    }

    rows.forEach(tr => {
        const cells = tr.querySelectorAll('td');
        if (cells.length < 3) return;
        
        const texto = cells[1].textContent.trim();
        
        const matchTrocar = texto.match(/TROCAR\s+(\d+)\s+RS/i);
        const qtyTroca = matchTrocar ? parseInt(matchTrocar[1], 10) : (texto.match(/TROCAR\s+RS/i) ? 1 : 0);

        if (qtyTroca > 0) {
            const upText = texto.toUpperCase();
            
            if (upText.includes('RS M AC') || upText.includes('RS MAC')) {
                addAtivo(window._ramaisAtivosInstalando, 'MAC', qtyTroca * 20);
                addAtivo(window._ramaisAtivosRemovendo, 'MAC', qtyTroca * 15);
            } else if (upText.includes('RS M AM') || upText.includes('RS MAM')) {
                addAtivo(window._ramaisAtivosInstalando, 'MAM', qtyTroca * 20);
                addAtivo(window._ramaisAtivosRemovendo, 'MAM', qtyTroca * 15);
            } else if (upText.includes('RS T AM') || upText.includes('RS TAM')) {
                addAtivo(window._ramaisAtivosInstalando, 'TAM', qtyTroca * 20);
                addAtivo(window._ramaisAtivosRemovendo, 'TAM', qtyTroca * 15);
            } else if (upText.includes('RS M AA') || upText.includes('RS MAA')) {
                addAtivo(window._ramaisAtivosInstalando, 'MAM', qtyTroca * 20);
                addAtivo(window._ramaisAtivosRemovendo, 'MAA', qtyTroca * 1);
            } else if (upText.includes('RS T AA') || upText.includes('RS TAA')) {
                addAtivo(window._ramaisAtivosInstalando, 'TAM', qtyTroca * 20);
                addAtivo(window._ramaisAtivosRemovendo, 'MAA', qtyTroca * 2);
            }

            if (upText.includes('CP-REDE') || upText.includes('CP REDE') || upText.includes('CPREDE')) {
                addAtivo(window._ramaisAtivosInstalando, 'CPREDE', qtyTroca * 1);
            }
        }
    });

    const formatDict = (dict) => Object.keys(dict).sort().map(k => dict[k]+"-"+k).join(' ');
    
    const strInstalando = formatDict(window._ramaisAtivosInstalando);
    const strRemovendo = formatDict(window._ramaisAtivosRemovendo);

    const spanInstalando = document.getElementById('ramais-instalando');
    const spanRemovendo = document.getElementById('ramais-removendo');
    
    if (spanInstalando) spanInstalando.textContent = strInstalando || '-';
    if (spanRemovendo) spanRemovendo.textContent = strRemovendo || '-';
};

window.adicionarAtivosRamais = function() {
    const instDict = window._ramaisAtivosInstalando || {};
    const remDict = window._ramaisAtivosRemovendo || {};

    const instKeys = Object.keys(instDict);
    const remKeys = Object.keys(remDict);

    if (instKeys.length === 0 && remKeys.length === 0) {
        alert('Não há ativos calculados para adicionar.');
        return;
    }

    if (!confirm('Deseja adicionar esses ativos gerados à tabela OUTROS?')) {
        return;
    }

    if (typeof pushHistory === 'function') pushHistory();

    instKeys.forEach(k => {
        tableStates.outros.data.push({
            entidade: '0',
            operacao: 'I',
            ativo: instDict[k] + "-" + k,
            texto: 'RAMAIS (GERADO)'
        });
    });

    remKeys.forEach(k => {
        tableStates.outros.data.push({
            entidade: '0',
            operacao: 'R',
            ativo: remDict[k] + "-" + k,
            texto: 'RAMAIS (GERADO)'
        });
    });
    
    if (typeof recalcAllQtdAtivos === 'function') recalcAllQtdAtivos();
    if (typeof renderTable === 'function') renderTable('outros');
    if (typeof updateCounters === 'function') updateCounters();
    if (typeof updateHistoryUI === 'function') updateHistoryUI();
    if (typeof buildAtivoSets === 'function') buildAtivoSets();
    if (typeof buildDataLists === 'function') buildDataLists();
    
    document.getElementById('modal-ramais')?.classList.add('hidden');
    
    const btn = document.querySelector('#modal-ramais button.btn-primary');
    if (btn) {
        const orig = btn.textContent;
        btn.textContent = 'Adicionado!';
        setTimeout(() => { btn.textContent = orig; }, 1500);
    }
};

window.copyRamaisTable = function() {
    const rows = document.querySelectorAll('#body-modal-ramais tr');
    if (!rows.length) return;

    const lines = [];
    rows.forEach(tr => {
        const cells = tr.querySelectorAll('td');
        if (cells.length < 2) return;
        const texto = cells[1].textContent.trim();
        if (texto) lines.push(texto);
    });

    if (!lines.length) return;

    navigator.clipboard.writeText(lines.join('\n')).then(() => {
        const btn = document.getElementById('btn-copy-ramais');
        if (btn) {
            const orig = btn.textContent;
            btn.textContent = '✓ Copiado!';
            btn.style.color = '#3fb950';
            setTimeout(() => { btn.textContent = orig; btn.style.color = ''; }, 1800);
        }
    }).catch(() => alert('Não foi possível copiar. Use Ctrl+C manualmente.'));
};

/* ═══════════════════════════════════════
   TABELA TOTALIZADORA
═══════════════════════════════════════ */
window.toggleViewMode = function() {
    const isPadrao = document.querySelector('input[name="view_mode"][value="padrao"]').checked;
    const viewPadrao = document.getElementById('view-padrao');
    const viewTotalizadora = document.getElementById('view-totalizadora');
    const btnGerar = document.getElementById('btn-gerar-totalizadora');

    if (isPadrao) {
        viewPadrao.classList.remove('hidden');
        viewTotalizadora.classList.add('hidden');
        btnGerar.style.display = 'none';
    } else {
        viewPadrao.classList.add('hidden');
        viewTotalizadora.classList.remove('hidden');
        btnGerar.style.display = 'inline-block';
        
        carregarRegras();
        // Se a totalizadora estiver vazia, preencher automaticamente
        if (tableStates.totalizadora.data.length === 0) {
            syncTotalizadora(false);
        }
    }
};

/* ═══════════════════════════════════════
   TABELA DE REGRAS DE CONVERSÃO
═══════════════════════════════════════ */
window.toggleRegras = function() {
    const content = document.getElementById('regras-content');
    const icon = document.getElementById('regras-toggle-icon');
    if (content.classList.contains('hidden')) {
        content.classList.remove('hidden');
        icon.textContent = '▲ Ocultar';
    } else {
        content.classList.add('hidden');
        icon.textContent = '▼ Mostrar';
    }
};

window.carregarRegras = async function() {
    try {
        const projCode = localStorage.getItem('projeto_selecionado_codigo') || 'DEFAULT';
        const key = `regras_orcamento_${projCode}`;

        // 1. Tenta carregar da nuvem primeiro
        try {
            const res = await fetch(`/api/regras/conversao?projeto_codigo=${encodeURIComponent(projCode)}`);
            if (res.ok) {
                const data = await res.json();
                if (data.regras && data.regras.length > 0) {
                    tableStates.regras.data = data.regras;
                    // Sincroniza nuvem → localStorage
                    localStorage.setItem(key, JSON.stringify(data.regras));
                    renderRegrasTable();
                    refreshTotalizadoraProjectData();
                    return;
                }
            }
        } catch (netErr) {
            console.warn('Falha ao buscar regras da nuvem, usando localStorage.', netErr);
        }

        // 2. Fallback: localStorage
        const saved = localStorage.getItem(key);
        if (saved) {
            tableStates.regras.data = JSON.parse(saved);
        } else {
            const legacy = localStorage.getItem('regras_orcamento');
            if (legacy && projCode === 'DEFAULT') {
                tableStates.regras.data = JSON.parse(legacy);
            } else {
                tableStates.regras.data = [{
                    origem: 'CABOS', op_de: '', ativo_de: '', acao: 'ADICAO',
                    op_para: '', ativo_para: '', fator: 1.05,
                    arredondamento: 'NORMAL', val_min: '', val_max: ''
                }];
            }
        }
        renderRegrasTable();
        refreshTotalizadoraProjectData();
    } catch (e) {
        console.error("Erro ao carregar regras", e);
    }
};

async function refreshTotalizadoraProjectData() {
    if (tableStates.totalizadora && tableStates.totalizadora.data && tableStates.totalizadora.data.length > 0) {
        for(let i=0; i<tableStates.totalizadora.data.length; i++) {
            if (window.updateTotRow) {
                await window.updateTotRow(i, 'ativo', tableStates.totalizadora.data[i].ativo);
            }
        }
    }
}

window.salvarRegras = function() {
    const projCode = localStorage.getItem('projeto_selecionado_codigo') || 'DEFAULT';
    localStorage.setItem(`regras_orcamento_${projCode}`, JSON.stringify(tableStates.regras.data));
};

window.salvarRegrasNuvem = async function() {
    const projCode = localStorage.getItem('projeto_selecionado_codigo') || 'DEFAULT';
    salvarRegras(); // Salva localmente também
    try {
        const res = await fetch('/api/regras/conversao', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ projeto_codigo: projCode, regras: tableStates.regras.data })
        });
        if (res.ok) {
            showToast('✓ Regras salvas na nuvem!');
        } else {
            const data = await res.json().catch(() => ({}));
            showToast(data.detail || 'Erro ao salvar na nuvem.');
        }
    } catch(e) {
        showToast('Erro de conexão ao salvar.');
    }
};

window.carregarRegrasNuvem = async function() {
    const projCode = localStorage.getItem('projeto_selecionado_codigo') || 'DEFAULT';
    try {
        const res = await fetch(`/api/regras/conversao?projeto_codigo=${encodeURIComponent(projCode)}`);
        if (res.ok) {
            const data = await res.json();
            if (data.regras && data.regras.length > 0) {
                tableStates.regras.data = data.regras;
                salvarRegras(); // Sincroniza com localStorage
                renderRegrasTable();
                showToast('Regras carregadas da nuvem!');
                return;
            }
        }
    } catch(e) {
        console.warn('Falha ao carregar regras da nuvem, usando localStorage.', e);
    }
};

window.exportarRegrasCSV = function() {
    if (!tableStates.regras.data || tableStates.regras.data.length === 0) {
        alert("Nenhuma regra para exportar.");
        return;
    }
    const headers = ["ORIGEM", "OP_DE", "ATIVO_DE", "ACAO", "OP_PARA", "ATIVO_PARA", "FATOR", "ARREDONDAMENTO", "VAL_MIN", "VAL_MAX"];
    const rows = tableStates.regras.data.map(r => [
        r.origem || "", r.op_de || "", r.ativo_de || "", r.acao || "ADIÇÃO", r.op_para || "", r.ativo_para || "",
        r.fator !== undefined && r.fator !== null ? String(r.fator).replace(".", ",") : "",
        r.arredondamento || "", r.val_min || "", r.val_max || ""
    ]);
    
    let csvContent = headers.join(";") + "\n";
    rows.forEach(row => {
        csvContent += row.map(v => `"${v}"`).join(";") + "\n";
    });
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const projCode = localStorage.getItem('projeto_selecionado_codigo') || 'DEFAULT';
    link.setAttribute("href", url);
    link.setAttribute("download", `regras_${projCode}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
};

window.importarRegrasCSV = function(input) {
    const file = input.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        const lines = text.split("\n").filter(l => l.trim() !== "");
        if (lines.length < 2) {
            alert("O arquivo não possui regras válidas.");
            return;
        }
        
        const newRegras = [];
        for (let i = 1; i < lines.length; i++) {
            const cols = lines[i].split(";").map(c => c.replace(/^"|"$/g, "").trim());
            if (cols.length >= 9) {
                let hasAcao = cols.length >= 10;
                newRegras.push({
                    origem: cols[0],
                    op_de: cols[1],
                    ativo_de: cols[2],
                    acao: hasAcao ? (cols[3].toUpperCase().includes('SUBST') ? 'SUBST' : 'ADICAO') : 'ADICAO',
                    op_para: hasAcao ? cols[4] : cols[3],
                    ativo_para: hasAcao ? cols[5] : cols[4],
                    fator: parseFloat((hasAcao ? cols[6] : cols[5]).replace(",", ".")) || 1.0,
                    arredondamento: hasAcao ? cols[7] : (cols[6] || "NORMAL"),
                    val_min: hasAcao ? cols[8] : cols[7],
                    val_max: hasAcao ? cols[9] : cols[8]
                });
            }
        }
        
        if (newRegras.length > 0) {
            tableStates.regras.data = newRegras;
            salvarRegras();
            renderRegrasTable();
            alert("Regras importadas com sucesso!");
        } else {
            alert("Nenhuma regra pôde ser importada. Verifique o formato.");
        }
        input.value = ""; // limpa o input
    };
    reader.readAsText(file, "utf-8");
};

window.renderRegrasTable = function() {
    const tbody = document.getElementById('body-regras');
    if (!tbody) return;
    tbody.innerHTML = '';
    
    tableStates.regras.data.forEach((r, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <select class="tot-input" onchange="updateRegraRow(${index}, 'origem', this.value)">
                    <option value="" ${r.origem === '' ? 'selected' : ''}>Qqlr</option>
                    <option value="CABOS" ${r.origem === 'CABOS' ? 'selected' : ''}>CABOS</option>
                    <option value="OUTROS" ${r.origem === 'OUTROS' ? 'selected' : ''}>OUTROS</option>
                    <option value="VINCULO" ${r.origem === 'VINCULO' ? 'selected' : ''}>VINCULO</option>
                </select>
            </td>
            <td><input type="text" class="tot-input" placeholder="Qqlr" value="${r.op_de || ''}" onchange="updateRegraRow(${index}, 'op_de', this.value)"></td>
            <td><input type="text" class="tot-input" placeholder="${r.origem === 'VINCULO' ? 'Ex: CAA2_N4' : 'Qqlr'}" value="${r.ativo_de || ''}" onchange="updateRegraRow(${index}, 'ativo_de', this.value)"></td>
            <td>
                <select class="tot-input" onchange="updateRegraRow(${index}, 'acao', this.value)">
                    <option value="ADICAO" ${(!r.acao || r.acao === 'ADICAO' || r.acao === 'ADI\u00C7\u00C3O' || r.acao === 'ADIC\u00C3O') ? 'selected' : ''}>ADI\u00C7\u00C3O</option>
                    <option value="SUBST" ${(r.acao === 'SUBST' || r.acao === 'SUBSTITUI\u00C7\u00C3O' || r.acao === 'SUBSTITUICAO') ? 'selected' : ''}>SUBST.</option>
                </select>
            </td>
            <td><input type="text" class="tot-input" placeholder="Manter" value="${r.op_para || ''}" onchange="updateRegraRow(${index}, 'op_para', this.value)"></td>
            <td><input type="text" class="tot-input" placeholder="Manter" value="${r.ativo_para || ''}" onchange="updateRegraRow(${index}, 'ativo_para', this.value)"></td>
            <td><input type="number" step="0.01" class="tot-input" value="${r.fator || 1}" onchange="updateRegraRow(${index}, 'fator', this.value)"></td>
            <td>
                <select class="tot-input" onchange="updateRegraRow(${index}, 'arredondamento', this.value)">
                    <option value="NORMAL" ${r.arredondamento === 'NORMAL' ? 'selected' : ''}>NORMAL</option>
                    <option value="PARA CIMA" ${r.arredondamento === 'PARA CIMA' ? 'selected' : ''}>CIMA</option>
                    <option value="PARA BAIXO" ${r.arredondamento === 'PARA BAIXO' ? 'selected' : ''}>BAIXO</option>
                    <option value="INTEIRO" ${r.arredondamento === 'INTEIRO' ? 'selected' : ''}>INTEIRO</option>
                </select>
            </td>
            <td><input type="number" step="0.01" class="tot-input" placeholder="-" value="${r.val_min !== '' && r.val_min != null ? r.val_min : ''}" onchange="updateRegraRow(${index}, 'val_min', this.value)"></td>
            <td><input type="number" step="0.01" class="tot-input" placeholder="-" value="${r.val_max !== '' && r.val_max != null ? r.val_max : ''}" onchange="updateRegraRow(${index}, 'val_max', this.value)"></td>
            <td style="text-align: center;">
                <button class="btn-primary" style="background-color: transparent; border: none; color: #f85149; padding: 4px;" onclick="excluirRegra(${index})" title="Excluir">
                    <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"></path></svg>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
};

window.adicionarRegra = function() {
    tableStates.regras.data.push({
        origem: '', op_de: '', ativo_de: '', acao: 'ADICAO',
        op_para: '', ativo_para: '', fator: 1,
        arredondamento: 'NORMAL', val_min: '', val_max: ''
    });
    salvarRegras();
    renderRegrasTable();
};

window.excluirRegra = function(index) {
    tableStates.regras.data.splice(index, 1);
    salvarRegras();
    renderRegrasTable();
};

window.updateRegraRow = function(index, field, value) {
    if (tableStates.regras.data[index]) {
        if (field === 'fator' || field === 'val_min' || field === 'val_max') {
            if (value === '') {
                value = '';
            } else {
                value = parseFloat(value);
                if (isNaN(value)) value = field === 'fator' ? 1 : '';
            }
        } else if (field === 'origem' || field === 'op_de' || field === 'ativo_de' || field === 'op_para' || field === 'ativo_para') {
            value = value.toUpperCase().trim();
        } else if (field === 'acao') {
            // Normaliza para valores simples sem acento
            const v = (value || '').toUpperCase();
            if (v === 'SUBST' || v.includes('SUBSTITU')) value = 'SUBST';
            else value = 'ADICAO';
        }
        tableStates.regras.data[index][field] = value;
        salvarRegras();
    }
};

/* ── TASK-042: editor amigável das Regras de Conversão com origem=VINCULO (ativo composto → conector),
   acessível dentro do modal de vinculação. Mesmo array/endpoint da Tabela de Regras de Conversão
   (tableStates.regras.data, salvarRegrasNuvem) — só uma visão filtrada e com campos separados
   (cabo/estrutura) em vez do "ativo_de" bruto (ex.: "CAA2_U4"). ── */
function _splitAtivoVinculo(ativoDe) {
    const texto = (ativoDe || '').trim();
    const idx = texto.indexOf('_');
    if (idx === -1) return { cabo: texto, estrutura: '' };
    return { cabo: texto.slice(0, idx), estrutura: texto.slice(idx + 1) };
}

window.toggleVincConectores = function () {
    const content = document.getElementById('vinc-conectores-content');
    const icon = document.getElementById('vinc-conectores-toggle-icon');
    if (!content || !icon) return;
    if (content.classList.contains('hidden')) {
        content.classList.remove('hidden');
        icon.textContent = '▲ Ocultar';
        if (!tableStates.regras.data || tableStates.regras.data.length === 0) {
            carregarRegras().then(renderRegrasConectoresTable);
        } else {
            renderRegrasConectoresTable();
        }
    } else {
        content.classList.add('hidden');
        icon.textContent = '▼ Mostrar';
    }
};

window.renderRegrasConectoresTable = function () {
    const tbody = document.getElementById('body-vinc-conectores');
    if (!tbody) return;
    tbody.innerHTML = '';
    (tableStates.regras.data || []).forEach((r, index) => {
        if ((r.origem || '').toUpperCase() !== 'VINCULO') return;
        const partes = _splitAtivoVinculo(r.ativo_de);
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" class="tot-input" placeholder="Ex: CAA2" value="${_escapeHtmlVinculacao(partes.cabo)}" onchange="updateRegraConectorRow(${index}, 'cabo', this.value)"></td>
            <td><input type="text" class="tot-input" placeholder="Ex: U4" value="${_escapeHtmlVinculacao(partes.estrutura)}" onchange="updateRegraConectorRow(${index}, 'estrutura', this.value)"></td>
            <td><input type="text" class="tot-input" placeholder="Ex: CONECTOR-X" value="${_escapeHtmlVinculacao(r.ativo_para || '')}" onchange="updateRegraConectorRow(${index}, 'ativo_para', this.value)"></td>
            <td><input type="number" min="0" step="0.01" class="tot-input" style="width:80px;" value="${r.fator != null ? r.fator : 1}" onchange="updateRegraConectorRow(${index}, 'fator', this.value)"></td>
            <td style="text-align:center;">
                <button class="btn-primary" style="background-color: transparent; border: none; color: #f85149; padding: 4px;" onclick="excluirRegraConector(${index})" title="Excluir">
                    <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"></path></svg>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
};

window.updateRegraConectorRow = function (index, field, value) {
    const r = tableStates.regras.data[index];
    if (!r) return;
    if (field === 'cabo' || field === 'estrutura') {
        const partes = _splitAtivoVinculo(r.ativo_de);
        if (field === 'cabo') partes.cabo = (value || '').toUpperCase().trim();
        else partes.estrutura = (value || '').toUpperCase().trim();
        r.ativo_de = `${partes.cabo}_${partes.estrutura}`;
    } else if (field === 'ativo_para') {
        r.ativo_para = (value || '').toUpperCase().trim();
    } else if (field === 'fator') {
        const v = parseFloat(value);
        r.fator = isNaN(v) ? 1 : v;
    }
    salvarRegras();
};

window.adicionarRegraConector = function () {
    tableStates.regras.data.push({
        origem: 'VINCULO', op_de: '', ativo_de: '_', acao: 'ADICAO',
        op_para: '', ativo_para: '', fator: 1,
        arredondamento: 'NORMAL', val_min: '', val_max: ''
    });
    salvarRegras();
    renderRegrasConectoresTable();
};

window.excluirRegraConector = function (index) {
    tableStates.regras.data.splice(index, 1);
    salvarRegras();
    renderRegrasConectoresTable();
};

let orcamentoBaseData = [];

// TASK-123: avalia o ativo da Totalizadora SEPARANDO o nome dos parâmetros (static/ativo_avaliar.js). Nome inexistente → vermelho (naoEncontrado);
// nome que existe mas parâmetros com problema → âmbar (avisoParametros), com a mensagem na descrição.
let _idxAvaliar = null;
function avaliarTot(texto, origem) {
    const proj = localStorage.getItem('projeto_selecionado') || '';
    const rows = orcamentoBaseData;
    if (!_idxAvaliar || _idxAvaliar.rows !== rows || _idxAvaliar.n !== rows.length || _idxAvaliar.proj !== proj) {
        _idxAvaliar = { rows, n: rows.length, proj, idx: AtivoAvaliar.construirIndice(rows, proj) };
    }
    const r = AtivoAvaliar.avaliarAtivo(texto, origem, _idxAvaliar.idx);
    if (r.naoEncontrado) return { desc: '', naoEncontrado: true, avisoParametros: false };
    if (r.erroParametros) return { desc: r.erroParametros, naoEncontrado: false, avisoParametros: true };
    return { desc: r.desc, naoEncontrado: false, avisoParametros: false };
}
// TASK-122: ao digitar "(" na coluna Ativo de Outros, mostra os parâmetros aceitos do ativo (ou "ativo sem parâmetros declarados" / "ativo não encontrado")
if (window.AtivoDica) window.AtivoDica.iniciar({
    base: () => orcamentoBaseData,
    projeto: () => localStorage.getItem('projeto_selecionado') || '',
    carregarBase: () => carregarBaseOrcamento(),
    // TASK-181: de onde vem o valor que vale em cada parâmetro (manual, vínculo, da obra, padrão, automático) — o mesmo status do balão ao passar o mouse
    origens: (input, ativo) => {
        try {
            if (input.dataset.parTipo !== 'outros') return null;
            const idx = Number(input.dataset.parIdx), c = _contextoParametros();
            const sts = _statusParametrosDaLinha(tableStates.outros.data[idx], _parametrosVinculoPorLinha()[idx], c.assin, c.ctx, c.obra, _defineDoVinculo()[idx]);
            const s = sts.find(x => String(AtivoParametros.nomeBase(x.ativo)).toUpperCase() === ativo);
            if (!s) return null;
            const m = {};
            s.params.forEach(p => { m[String(p.nome).toUpperCase()] = { valor: p.valor, rotulo: AtivoStatus.rotuloOrigem(p.origem, p.detalhe) }; });
            return m;
        } catch (e) { return null; }
    }
});
// TASK-186: balão interativo dos ativos/parâmetros de Outros (Ctrl+Enter, duplo clique, mouse parado, "(" com «Ocultar parâmetros» ligado)
if (window.AtivoBalao && window.AtivoItens) window.AtivoBalao.iniciar({
    itens: input => {
        const idx = Number(input.dataset.parIdx), row = tableStates.outros.data[idx], itens = window.AtivoItens.itens(row ? row.ativo : '');
        let al = [], sts = [];
        try {
            const c = _contextoParametros(), pv = _parametrosVinculoPorLinha()[idx], dv = _defineDoVinculo()[idx];
            al = _statusParametrosDaLinha(row, pv, c.assin, c.ctx, c.obra, dv, true); sts = al.filter(Boolean);
        } catch (e) { al = []; sts = []; }
        // cada item do texto casa com o item interpretado NA MESMA POSIÇÃO (o interpretador mantém a ordem e converte nomes: 11300D → DT11/300, apelidos…); sem o mesmo número de itens, cai para o casamento pelo nome
        const porPosicao = al.length === itens.length;
        return itens.map((it, i) => ({ nome: it.nome, base: it.base, params: it.params, raw: it.raw,
            status: porPosicao ? al[i] : (sts.find(x => String(AtivoParametros.nomeBase(x.ativo)).toUpperCase() === String(AtivoParametros.nomeBase(it.base || it.nome)).toUpperCase()) || null) }));
    },
    dica: nome => { try { const c = window.AtivoDica && window.AtivoDica.contextoAtual(); return c ? window.AtivoDica.dica(nome, c) : null; } catch (e) { return null; } },
    origens: (input, ativo) => {
        try {
            const idx = Number(input.dataset.parIdx), c = _contextoParametros(), sts = _statusParametrosDaLinha(tableStates.outros.data[idx], _parametrosVinculoPorLinha()[idx], c.assin, c.ctx, c.obra, _defineDoVinculo()[idx]);
            const s = sts.find(x => String(AtivoParametros.nomeBase(x.ativo)).toUpperCase() === ativo);
            if (!s) return null;
            const m = {};
            s.params.forEach(p => { m[String(p.nome).toUpperCase()] = { valor: p.falta || p.invalido ? null : p.valor, rotulo: AtivoStatus.rotuloOrigem(p.origem, p.detalhe) }; });
            return m;
        } catch (e) { return null; }
    },
    gravar: (input, i, nomeParam, valor) => {
        const idx = Number(input.dataset.parIdx), row = tableStates.outros.data[idx], AD = window.AtivoDica;
        const c = AD.contextoAtual();
        // TASK-194: a base conhece «CON», não «6 CON» — a gravação procura a dica pelo nome SEM a quantidade (como a leitura já fazia) e mantém a quantidade no texto
        const novo = window.AtivoItens.gravarParametro(row.ativo, i, nomeParam, valor, nome => c && AD.dica(nome, c), AD.aplicarValor);
        if (novo !== null) input._gravarCheio(novo);
    },
    carregarBase: () => carregarBaseOrcamento(),
    mensagem: t => showToast(t)
});
// TASK-128: a base de orçamento é pedida JÁ ao abrir a tela (em segundo plano) — antes só no primeiro "(" digitado, e a resposta demora no início (sincroniza com a nuvem)
/** Prepara (em segundo plano) o que varre a base inteira — contexto de Outros, assinaturas e contexto da caixa de parâmetros — para o 1º uso já sair rápido. */
function aquecerCaches() { try { _contextoParametros(); if (window.AtivoDica && window.AtivoDica.aquecer) window.AtivoDica.aquecer(); } catch (e) { /* base ou convenção ainda carregando */ } }
let _basePromessa = null, _baseOutraVez = false;
function carregarBaseOrcamento(refazer) {
    // `refazer` (a base mudou agora): pedido durante um download em andamento (a base mudou agora há pouco): baixa de novo assim que terminar, em vez de ficar com a resposta antiga
    if (_basePromessa) { if (refazer) _baseOutraVez = true; return _basePromessa; }
    _basePromessa = fetch('/api/orcamento/dados').then(r => r.json()).then(d => { orcamentoBaseData = d.dados || []; if (window.aplicarCoresParametros) window.aplicarCoresParametros(); setTimeout(aquecerCaches, 30); }).catch(() => { /* sem base: sem dica */ }).finally(() => { _basePromessa = null; if (_baseOutraVez) { _baseOutraVez = false; carregarBaseOrcamento(); } });
    return _basePromessa;
}
carregarBaseOrcamento();
// a base foi salva/publicada/sincronizada na tela da Base de Orçamento (outra aba, janela ou moldura): recarrega na hora — o ativo novo passa a valer nas sugestões e nas cores
if (window.BaseSinal) window.BaseSinal.ouvir(origem => { if (origem !== 'local') carregarBaseOrcamento(true).then(() => { if (typeof showToast === 'function') showToast('Base de orçamento atualizada.'); }); });
// TASK-142: ao fechar a Tabela de orçamento (moldura do programa) a base é pedida de novo — ativo editado/sincronizado passa a valer na hora na tela de baixo
document.addEventListener('tela-app-fechada', e => { if (/\/orcamento(\?|$|#)/.test((e.detail && e.detail.url) || '')) carregarBaseOrcamento(); });
// versão online (a tabela abre em outra aba): ao voltar para esta aba depois de um tempo, também recarrega a base
(function () {
    let saiuEm = 0;
    document.addEventListener('visibilitychange', () => {
        if (document.hidden) { saiuEm = Date.now(); return; }
        if (saiuEm && Date.now() - saiuEm > 15000) carregarBaseOrcamento();
        saiuEm = 0;
    });
})();
// TASK-136: sugestão de ativos/apelidos do PROJETO ao digitar (Outros e Totalizadora); botão "💡 Sugestões" liga/desliga
if (window.AtivoSugestao) window.AtivoSugestao.iniciar({
    base: () => orcamentoBaseData,
    projeto: () => localStorage.getItem('projeto_selecionado') || '',
    carregarBase: () => carregarBaseOrcamento()
});

window.syncTotalizadora = async function(forceUpdate = true) {
    if (forceUpdate && tableStates.totalizadora.data.length > 0) {
        if (!confirm("Isso irá apagar todas as edições manuais feitas na Tabela Totalizadora e puxar os dados novamente das tabelas Cabos e Outros. Deseja continuar?")) {
            return;
        }
    }

    const btnGerar = document.getElementById('btn-gerar-totalizadora');
    const originalText = btnGerar ? btnGerar.textContent : '';
    if (btnGerar) btnGerar.textContent = 'Carregando Base...';

    await prepararTotalizadora();

    if (btnGerar) btnGerar.textContent = originalText;

    tableStates.totalizadora.data = construirTotalizadora(tableStates.cabos.data, tableStates.outros.data, true);
    renderTotalizadora();
    showToast("Tabela Totalizadora atualizada!");
};

async function prepararTotalizadora() {
    if (!tableStates.regras.data.length) await carregarRegras();
    if (!orcamentoBaseData.length) await carregarBaseOrcamento();
    if (window.ObraParams) await ObraParams.carregarTabela(localStorage.getItem('projeto_selecionado') || '');
}

function construirTotalizadora(cabosData, outrosData, emUso = false) {
    function buscarDescricao(ativoNome) {
        const ativoUpper = (ativoNome || '').trim().toUpperCase();
        const projName = (localStorage.getItem('projeto_selecionado') || '').toUpperCase();
        
        // TASK-079: compara pelo nome-base (N1(XLPE,DT) e N1(cabo=ALUM) → N1); prefere a linha sem parâmetros (comum/declaração)
        const nb = AtivoParametros.nomeBase(ativoUpper);
        const doProjeto = r => (r.projeto || '').trim().toUpperCase() === projName || (r.projeto || '').trim() === '';
        const candidatas = orcamentoBaseData.filter(r =>
            ((r.ativo || '') !== '' && (AtivoParametros.nomeBase(r.ativo) === nb || AtivoParametros.semColchetes(r.ativo).trim().toUpperCase() === ativoUpper) || (r.codigo || '').toUpperCase() === ativoUpper) && doProjeto(r));      // TASK-108: [CAZ7,94] casa "CAZ7,94"
        const row = candidatas.find(r => AtivoParametros.analisarCelulaBase(r.ativo || '').tipo !== 'grupo') || candidatas[0];
        if (row) {
            return { desc: row.desc_ativo || row.componente || row.desc_codigo || '', naoEncontrado: false };
        }
        return { desc: '', naoEncontrado: true };
    }

    const rawItems = [];
    let baseIdCounter = 1;

    // Processar Cabos
    for (let idx = 0; idx < cabosData.length; idx++) {
        const item = cabosData[idx];
        if (!item || item.congelada) continue;                                                   // linha congelada: fora da Totalizadora
        const currentBaseId = baseIdCounter++;
        if (AnotacaoCabos.ehNotacao(item.ativo)) {                                                // TASK-083: notação do LABEL → cabo (N vezes) + neutro à parte
            const nr = AnotacaoCabos.interpretar(item.ativo);
            if (nr.erros.length) {
                rawItems.push({ baseId: `TOT-${currentBaseId}`, obs: item.entidade, operacao: item.operacao || 'I', ativo: String(item.ativo).toUpperCase(),
                    qtd: 1, desc: nr.erros.join('; '), naoEncontrado: true, origem: 'CABOS' });
                continue;
            }
            nr.itens.forEach(it => {
                const resNr = buscarDescricao(it.ativo);
                let reps = it.fases;
                if (!it.neutro && item.qtdAtivos) {
                    const pq = parseFloat(item.qtdAtivos);
                    if (!isNaN(pq) && pq > 0) reps = Math.floor(pq);
                }
                for (let i = 0; i < reps; i++) {
                    rawItems.push({ baseId: `TOT-${currentBaseId}`, obs: item.entidade, operacao: item.operacao || 'I', ativo: it.ativo, qtd: it.comprimento,
                        desc: resNr.desc, naoEncontrado: resNr.naoEncontrado, origem: 'CABOS' });
                }
            });
            continue;
        }
        let q = 1;
        let at = item.ativo;
        const m1 = item.ativo.match(/([\d\.,]+)\s*m\s*\|\s*(.*)/i);
        if (m1) {
            q = parseFloat(m1[1].replace(',', '.'));
            at = m1[2].trim();
        } else {
            const m2 = item.ativo.match(/(.+?)\s+([\d\.,]+)\s*m$/i);
            if (m2) {
                q = parseFloat(m2[2].replace(',', '.'));
                at = m2[1].trim();
            } else {
                // Tenta herdar comprimento da próxima linha para casos standalone (ex: P 50)
                if (typeof isStandaloneLine === 'function' && isStandaloneLine(item.ativo)) {
                    if (idx + 1 < cabosData.length) {
                        const nextItem = cabosData[idx + 1];
                        if (nextItem && nextItem.ativo) {
                            const nextM1 = nextItem.ativo.match(/([\d\.,]+)\s*m\s*\|\s*(.*)/i);
                            const nextM2 = nextItem.ativo.match(/(.+?)\s+([\d\.,]+)\s*m$/i);
                            if (nextM1) q = parseFloat(nextM1[1].replace(',', '.'));
                            else if (nextM2) q = parseFloat(nextM2[2].replace(',', '.'));
                        }
                    }
                }
            }
        }
        
        let formatado = at.toUpperCase().replace(/\//g, "").replace("CAA ", "CAA").replace("CA ", "CA").replace("CU ", "CU").replace("CAZ ", "CAZ").replace(/\bP\s+/g, "P");
        let parts = formatado.trim().split(/\s+/);
        let baseAtivo = parts[0] || at;
        let resDesc = buscarDescricao(baseAtivo);

        let iteracoes = 1;
        if (item.qtdAtivos) {
            let parsedQtd = parseFloat(item.qtdAtivos);
            if (!isNaN(parsedQtd) && parsedQtd > 0) {
                iteracoes = Math.floor(parsedQtd); // Tratar como linhas separadas
            }
        }

        for (let i = 0; i < iteracoes; i++) {
            rawItems.push({
                baseId: `TOT-${currentBaseId}`,
                obs: item.entidade,
                operacao: item.operacao || 'I',
                ativo: baseAtivo,
                qtd: q,
                desc: resDesc.desc,
                naoEncontrado: resDesc.naoEncontrado,
                origem: 'CABOS'
            });
        }
    }

    // Processar Outros
    // TASK-079: ativos com parâmetros — o texto entre parênteses não é quebrado e o pedido é normalizado (N1 → N1(XLPE,DT) pelo padrão)
    const assinaturasParam = AtivoParametros.assinaturasDoProjeto(orcamentoBaseData, localStorage.getItem('projeto_selecionado') || '');
    const ctxOutros = _ctxOutros();
    const cfg = window.__vinculacaoConfig;
    const parametrosVinculo = cfg ? (emUso ? _parametrosVinculoPorLinha() :
        VinculacaoRegras.parametrosPorEstrutura(cabosData, outrosData, cfg, _prefixoCaboVinculo, _codigoEstruturaVinculo, _codigosEstruturaVinculo)) : {};
    const postesDosPontos = _postesDosPontos();
    outrosData.forEach((item, idxOutros) => {
        if (!item || item.congelada) return;                                                     // linha congelada: fora da Totalizadora
        const currentBaseId = baseIdCounter++;
        
        // TASK-081: notação compacta/apelidos/regra do M traduzidos aqui; o que o usuário digitou não muda
        const ctxLinha = _ctxComPosteDoPonto(parametrosVinculo[idxOutros] ? Object.assign({}, ctxOutros, { parametros_estrutura: parametrosVinculo[idxOutros] }) : ctxOutros, item, postesDosPontos);
        const interp = AnotacaoOutros.interpretarLinha(item.ativo, item.operacao || 'I', _classicoTotalizadora, ctxLinha);
        if (interp.erros.length) {
            rawItems.push({
                baseId: `TOT-${currentBaseId}`, obs: item.entidade, operacao: item.operacao || 'I', ativo: String(item.ativo || '').toUpperCase(), qtd: 1,
                desc: interp.erros.map(e => e.msg).join('; '), naoEncontrado: true, origem: 'OUTROS'
            });
            return;
        }

        interp.itens.forEach(it => {
            let aName = it.ativo;
            const q = it.qtd;
            const ped = AtivoParametros.resolverPedido(aName, assinaturasParam, window.ObraParams ? ObraParams.valores() : null);
            if (ped.com_assinatura || aName.indexOf('(') >= 0) aName = ped.erros.length ? aName.toUpperCase() : ped.texto;
            let resDesc = buscarDescricao(aName);
            // TASK-123: nome que existe com parâmetros com problema = âmbar (avisoParametros, mensagem na descrição); nome inexistente = vermelho (como antes)
            if (ped.erros.length) resDesc = resDesc.naoEncontrado ? { desc: ped.erros.join('; '), naoEncontrado: true } : { desc: ped.erros.join('; '), naoEncontrado: false, avisoParametros: true };

            rawItems.push({
                baseId: `TOT-${currentBaseId}`,
                obs: item.entidade,
                operacao: it.operacao || 'I',
                ativo: aName,
                qtd: q,
                desc: resDesc.desc,
                naoEncontrado: resDesc.naoEncontrado,
                ...(resDesc.avisoParametros ? { avisoParametros: true } : {}),
                origem: 'OUTROS'
            });
        });
    });

    // Ativos compostos do vínculo cabo<->estrutura (TASK-034): um pseudo-item por aresta do
    // vínculo (TASK-032), origem "VINCULO", para as Regras de Conversão poderem casar contra ele
    // (ex.: ativo_de=CAA2_N4). Só existe aqui dentro — nunca aparece na Totalizadora sem match.
    cabosData.forEach((cabo, idxCabo) => {
        if (!cabo || cabo.congelada || !VinculacaoRegras.participaVinculo(cabo)) return;                        // TASK-095: cabo R/*R não entra no vínculo; congelado só alimenta o vínculo feito à mão
        const ficticiasCabo = VinculacaoRegras.estruturasFicticiasDe(cabo);
        if ((!Array.isArray(cabo.vinculoEstruturas) || !cabo.vinculoEstruturas.length) && !ficticiasCabo.length) return;
        const info = extrairPrefixoEComprimentoCabo(cabo.ativo);
        const prefixoCabo = info ? info.prefixo : null;
        if (!prefixoCabo) return;
        ficticiasCabo.forEach((cod, k) => rawItems.push({            // TASK-095: estrutura fictícia do cabo → CABO_ESTRUTURA
            baseId: `TOT-VINC-${idxCabo}-F${k}`, obs: 'VINCULO', operacao: cabo.operacao || 'I', ativo: `${prefixoCabo}_${cod}`, qtd: 1, desc: '', naoEncontrado: false, origem: 'VINCULO'
        }));
        (cabo.vinculoEstruturas || []).forEach(idxEstrutura => {
            const estrutura = outrosData[VinculacaoRegras.linhaDe(idxEstrutura)];        // TASK-094: unidade = linha (+ posição/100)
            if (!estrutura || estrutura.congelada || !estrutura.ativo || !VinculacaoRegras.participaVinculo(estrutura)) return;
            const codigosLinha = _codigosEstruturaVinculo(estrutura.ativo), codigoUnidade = codigosLinha[VinculacaoRegras.slotDe(idxEstrutura)] || _codigoEstruturaVinculo(estrutura.ativo);
            rawItems.push({
                baseId: `TOT-VINC-${idxCabo}-${idxEstrutura}`,
                obs: 'VINCULO',
                operacao: cabo.operacao || 'I',
                ativo: `${prefixoCabo}_${codigoUnidade.toUpperCase()}`,
                qtd: 1,
                desc: '',
                naoEncontrado: false,
                origem: 'VINCULO'
            });
        });
    });

    // MOTOR DE REGRAS
    const newData = [];
    const regras = tableStates.regras.data;

    rawItems.forEach(item => {
        let matchEncontrado = false;
        let requiresSubstitution = false;
        
        regras.forEach(rule => {
            const ruleOrigem = (rule.origem || '').trim().toUpperCase();
            const ruleOp = (rule.op_de || '').trim().toUpperCase();
            const ruleAtivo = (rule.ativo_de || '').trim().toUpperCase();

            // Verifica se a regra se aplica ao item
            const checkMatch = (rulePattern, targetStr) => {
                if (rulePattern === '') return true;
                if (rulePattern === targetStr) return true;
                
                try {
                    let regexStr = rulePattern;
                    if (rulePattern.includes('%')) {
                        // Estilo LIKE: escapa caracteres especiais e converte % em .*
                        regexStr = '^' + rulePattern.replace(/[.+?^${}()|[\]\\]/g, '\\$&').replace(/%/g, '.*') + '$';
                    } else {
                        // Estilo Regex: exige match completo caso não ancorado (preserva retrocompatibilidade)
                        if (!regexStr.startsWith('^')) regexStr = '^' + regexStr;
                        if (!regexStr.endsWith('$')) regexStr = regexStr + '$';
                    }
                    return new RegExp(regexStr, 'i').test(targetStr);
                } catch (e) {
                    return false;
                }
            };

            let applies = true;
            if (!checkMatch(ruleOrigem, item.origem)) applies = false;
            if (!checkMatch(ruleOp, item.operacao)) applies = false;
            // TASK-079: a regra casa o texto completo (N1(XLPE,DT)) OU o nome-base (N1)
            if (!checkMatch(ruleAtivo, item.ativo) && !checkMatch(ruleAtivo, AtivoParametros.semColchetes(item.ativo)) && !(item.ativo.indexOf('(') >= 0 && checkMatch(ruleAtivo, AtivoParametros.nomeBase(item.ativo)))) applies = false;

            if (applies) {
                matchEncontrado = true;
                if (rule.acao === 'SUBST' || rule.acao === 'SUBSTITUIÇÃO' || rule.acao === 'SUBSTITUICAO') {
                    requiresSubstitution = true;
                }
                
                // Aplica a regra
                let newOp = rule.op_para ? rule.op_para.trim().toUpperCase() : item.operacao;
                let newAtivo = rule.ativo_para ? rule.ativo_para.trim().toUpperCase() : item.ativo;
                
                let newQtd = item.qtd;
                let fator = parseFloat(rule.fator);
                if (!isNaN(fator)) newQtd = newQtd * fator;
                
                const arr = rule.arredondamento;
                if (arr === 'PARA CIMA') newQtd = Math.ceil(newQtd);
                else if (arr === 'PARA BAIXO') newQtd = Math.floor(newQtd);
                else if (arr === 'INTEIRO') newQtd = Math.round(newQtd);
                else newQtd = Math.round(newQtd * 100) / 100; // NORMAL (duas casas)
                
                if (rule.val_min !== '' && rule.val_min != null) {
                    const min = parseFloat(rule.val_min);
                    if (!isNaN(min) && newQtd < min) newQtd = min;
                }
                if (rule.val_max !== '' && rule.val_max != null) {
                    const max = parseFloat(rule.val_max);
                    if (!isNaN(max) && newQtd > max) newQtd = max;
                }
                
                    let resNewDesc = buscarDescricao(newAtivo);
                    newData.push({
                        id: item.baseId,
                        obs: item.obs,
                        operacao: newOp,
                        ativo: newAtivo,
                        qtd: newQtd,
                        desc: resNewDesc.desc,
                        naoEncontrado: resNewDesc.naoEncontrado,
                        origem: item.origem
                    });
            }
        });
        
        // TASK-034: um ativo composto de VINCULO sem nenhuma regra de conversão correspondente
        // não gera nada na Totalizadora (não é um código de orçamento real, só uma chave interna
        // do vínculo) — diferente de CABOS/OUTROS, que mantêm o item original sem match.
        if (!matchEncontrado && item.origem === 'VINCULO') {
            return;
        }

        if (!matchEncontrado || !requiresSubstitution) {
            // Se nenhuma regra se aplicar OU se as regras aplicadas foram apenas de ADIÇÃO, mantemos o item original
            const itemClone = { ...item };
            itemClone.id = item.baseId;
            newData.push(itemClone);
        }
    });

    return newData;
}

window.renderTotalizadora = function() {
    const tbody = document.getElementById('body-totalizadora');
    const emptyMsg = document.getElementById('tot-empty-msg');
    
    tbody.innerHTML = '';
    
    if (tableStates.totalizadora.data.length === 0) {
        emptyMsg.style.display = 'block';
        return;
    }
    
    emptyMsg.style.display = 'none';

    tableStates.totalizadora.data.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.id = `tot-row-${index}`;
        if (row.naoEncontrado) {
            tr.style.backgroundColor = 'rgba(248, 81, 73, 0.15)';           // o NOME do ativo não existe na base
        } else if (row.avisoParametros) {
            tr.style.backgroundColor = 'rgba(210, 153, 34, 0.18)';          // o nome existe, mas os PARÂMETROS têm problema (a mensagem está na descrição)
        }
        
        tr.innerHTML = `
            <td style="font-weight: bold; color: var(--text-secondary); text-align: center;">${row.id}</td>
            <td><input type="text" class="tot-input" value="${row.obs || ''}" onchange="updateTotRow(${index}, 'obs', this.value)"></td>
            <td>
                <input type="text" class="tot-input tot-op-input" style="text-align: center;" maxlength="4" autocomplete="off" spellcheck="false" aria-label="Operação: I, *I, R, *R, M ou *M" title="Operação: I, *I, R, *R, M ou *M — digite ou escolha na lista" value="${row.operacao || ''}" onchange="validarOperacaoTot(this, ${index})">
            </td>
            <td><input type="text" class="tot-input tot-ativo-input" data-par-tipo="tot" data-par-idx="${index}" value="${row.ativo || ''}" onchange="updateTotRow(${index}, 'ativo', this.value)" onkeydown="handleTotKeydown(event, ${index}, 'ativo')" onpaste="handleTotPaste(event, ${index})"></td>
            <td><input type="number" step="0.01" class="tot-input" style="width: 70px; text-align: center;" value="${row.qtd !== undefined && row.qtd !== null ? row.qtd : ''}" onchange="updateTotRow(${index}, 'qtd', this.value)" onkeydown="handleTotKeydown(event, ${index}, 'qtd')" onpaste="handleTotPaste(event, ${index})"></td>
            <td><input type="text" id="tot-desc-${index}" class="tot-input" value="${row.desc || ''}" placeholder="Opcional" onchange="updateTotRow(${index}, 'desc', this.value)"></td>
            <td>
                <select class="tot-input" onchange="updateTotRow(${index}, 'origem', this.value)">
                    <option value="OUTROS" ${row.origem === 'OUTROS' ? 'selected' : ''}>OUTROS</option>
                    <option value="CABOS" ${row.origem === 'CABOS' ? 'selected' : ''}>CABOS</option>
                </select>
            </td>
            <td style="text-align: center;">
                <button class="btn-primary" style="background-color: transparent; border: none; color: #f85149; padding: 4px;" onclick="excluirTotRow(${index})" title="Excluir">
                    <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"></path></svg>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
};

window.validarOperacaoTot = function (inp, index) {              // TASK-182 C: operação digitada na Totalizadora (mesma validação das tabelas)
    const linha = tableStates.totalizadora.data[index], anterior = linha ? (linha.operacao || '') : '';
    const v = normalizarOperacao(inp.value);
    if (v === anterior) { inp.value = anterior; return; }
    if (!OPERACOES.includes(v)) { recusarOperacao(inp, anterior); return; }
    inp.value = v; inp.classList.remove('op-invalida');
    updateTotRow(index, 'operacao', v);
};
window.updateTotRow = async function(index, field, value) {
    if (tableStates.totalizadora.data[index]) {
        if (field === 'qtd') {
            const parsed = parseFloat(value);
            if (!isNaN(parsed)) {
                value = parsed;
            } else if (value === '' || value === '-' || value === '*') {
                // não força para 1 enquanto o usuário digita
            } else {
                value = 1;
            }
        }
        tableStates.totalizadora.data[index][field] = value;

        if (field === 'ativo' || field === 'origem') {
            // Garante que a base de orçamentos está carregada
            if (orcamentoBaseData.length === 0) {
                try {
                    const res = await fetch('/api/orcamento/dados');
                    const data = await res.json();
                    orcamentoBaseData = data.dados || [];
                } catch(e) {
                    console.error("Erro ao carregar base de orçamento", e);
                }
            }

            // TASK-123: o NOME é avaliado separado dos PARÂMETROS (antes comparava o texto inteiro "FLABIT(D,35)" e ficava vermelho)
            const lin = tableStates.totalizadora.data[index];
            const res = avaliarTot(field === 'ativo' ? value : lin.ativo, lin.origem);
            lin.desc = res.desc;
            lin.naoEncontrado = res.naoEncontrado;
            lin.avisoParametros = res.avisoParametros;

            const descInput = document.getElementById(`tot-desc-${index}`);
            if (descInput) descInput.value = res.desc;

            const tr = document.getElementById(`tot-row-${index}`);
            if (tr) tr.style.backgroundColor = res.naoEncontrado ? 'rgba(248, 81, 73, 0.15)' : (res.avisoParametros ? 'rgba(210, 153, 34, 0.18)' : '');
        }
    }
};

window.excluirTotRow = function(index) {
    tableStates.totalizadora.data.splice(index, 1);
    renderTotalizadora();
};

window.adicionarLinhaTotalizadora = function() {
    let nextIdNumber = 1;
    tableStates.totalizadora.data.forEach(r => {
        const m = r.id.match(/TOT-(\d+)/);
        if (m) {
            const num = parseInt(m[1], 10);
            if (num >= nextIdNumber) nextIdNumber = num + 1;
        }
    });

    tableStates.totalizadora.data.push({
        id: `TOT-${nextIdNumber}`,
        obs: '0',
        operacao: 'I',
        ativo: '',
        qtd: 1,
        desc: '',
        origem: 'OUTROS'
    });
    renderTotalizadora();
};

window.handleTotKeydown = function(e, index, field = 'ativo') {
    if (e.key === 'Enter') {
        e.preventDefault();
        
        if (field === 'qtd' && index + 1 < tableStates.totalizadora.data.length) {
            // Apenas move o foco para a quantidade da linha de baixo
            const nextRowTr = document.getElementById('body-totalizadora').children[index + 1];
            if (nextRowTr) {
                const nextInput = nextRowTr.querySelector('input[type="number"]');
                if (nextInput) {
                    nextInput.focus();
                    nextInput.select();
                }
            }
            return;
        }

        let nextIdNumber = 1;
        tableStates.totalizadora.data.forEach(r => {
            const m = r.id.match(/TOT-(\d+)/);
            if (m) {
                const num = parseInt(m[1], 10);
                if (num >= nextIdNumber) nextIdNumber = num + 1;
            }
        });
        
        const newRow = {
            id: `TOT-${nextIdNumber}`,
            obs: '0',
            operacao: 'I',
            ativo: '',
            qtd: 1,
            desc: '',
            origem: 'OUTROS'
        };
        
        tableStates.totalizadora.data.splice(index + 1, 0, newRow);
        renderTotalizadora();
        
        setTimeout(() => {
            const newRowTr = document.getElementById('body-totalizadora').children[index + 1];
            if (newRowTr) {
                const selector = field === 'ativo' ? '.tot-ativo-input' : 'input[type="number"]';
                const newInputs = newRowTr.querySelectorAll(selector);
                if (newInputs && newInputs.length > 0) newInputs[0].focus();
            }
        }, 50);
    } else if (e.key === 'Delete') {
        if (field === 'qtd') {
            e.preventDefault();
            e.target.value = '';
            updateTotRow(index, 'qtd', '');
        } else if (e.target.value === '') {
            e.preventDefault();
            excluirTotRow(index);
        }
    } else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
        e.preventDefault();
        const nextIndex = e.key === 'ArrowUp' ? index - 1 : index + 1;
        if (nextIndex >= 0 && nextIndex < tableStates.totalizadora.data.length) {
            const nextRowTr = document.getElementById('body-totalizadora').children[nextIndex];
            if (nextRowTr) {
                const selector = field === 'ativo' ? '.tot-ativo-input' : 'input[type="number"]';
                const nextInput = nextRowTr.querySelector(selector);
                if (nextInput) {
                    nextInput.focus();
                    nextInput.select();
                }
            }
        }
    }
};

window.handleTotPaste = function(e, index) {
    const paste = (e.clipboardData || window.clipboardData).getData('text');
    if (!paste) return;
    
    const lines = paste.split(/\r?\n/).filter(line => line.trim() !== '');
    if (lines.length === 0) return;
    
    const hasMultipleColumns = lines[0].includes('\t');
    if (lines.length <= 1 && !hasMultipleColumns) return;
    
    e.preventDefault();
    
    let currentIndex = index;
    
    for (let i = 0; i < lines.length; i++) {
        let ativo = lines[i].trim();
        let qtd = "";
        
        if (ativo.includes('\t')) {
            const cols = lines[i].split('\t');
            ativo = cols[0].trim();
            qtd = cols.length > 1 ? cols[1].trim() : "";
        }
        
        if (i === 0) {
            e.target.value = ativo;
            updateTotRow(index, 'ativo', ativo);
            if (qtd !== "") {
                updateTotRow(index, 'qtd', qtd);
                // Also update the input visually if the event target was not the qtd field
                const rowTr = document.getElementById('body-totalizadora').children[index];
                if (rowTr) {
                    const qtdInput = rowTr.querySelector('input[type="number"]');
                    if (qtdInput) qtdInput.value = qtd;
                }
            }
        } else {
            let nextIdNumber = 1;
            tableStates.totalizadora.data.forEach(r => {
                const m = r.id.match(/TOT-(\d+)/);
                if (m) {
                    const num = parseInt(m[1], 10);
                    if (num >= nextIdNumber) nextIdNumber = num + 1;
                }
            });
            
            const newRow = {
                baseId: `TOT-${nextIdNumber}`,
                id: `TOT-${nextIdNumber}`,
                obs: '0',
                operacao: 'I',
                ativo: ativo,
                qtd: qtd !== "" ? qtd : "",
                desc: '',
                origem: 'OUTROS'
            };
            
            tableStates.totalizadora.data.splice(currentIndex + 1, 0, newRow);
            currentIndex++;
        }
    }
    
    renderTotalizadora();
    
    // Processa a busca das descrições das linhas inseridas
    for (let i = 0; i < lines.length; i++) {
        const rowAtivo = tableStates.totalizadora.data[index + i].ativo;
        updateTotRow(index + i, 'ativo', rowAtivo);
    }
};


    /* TASK-106: aberto pelo plugin do nanoCAD ("Abrir na aplicação"): /resumo?obra=<id>&projeto=<código>&projeto_nome=<nome>.
       Seleciona o projeto, carrega a obra nas tabelas (como "Carregar") e tira os parâmetros da barra (F5 não recarrega por cima). */
    /* TASK-163: projeto entregue pelo plugin do nanoCAD (/resumo?plugin_pendente=<id>). O plugin grava o projeto de trabalho `<desenho>.obra`
       na pasta do desenho e entrega o CAMINHO; aqui ele é aberto como em 📁 Projeto → Abrir (static/projeto.js: abrirDoPlugin). Nada vai para o banco.
       Plugin antigo (só a obra, sem caminho): as linhas aparecem na tela, sem gravar no banco. */
    function mostrarTabelasDoResumo(...tabelas) {
        const padrao = document.querySelector('input[name="view_mode"][value="padrao"]');
        if (padrao && !padrao.checked) {
            padrao.checked = true;
            window.toggleViewMode();
        }
        window.ResumoLayout?.mostrarTabelas(...tabelas);
    }

    async function receberAnotacoesDoPlugin(url, lote, entrega) {
        if (entrega.estado !== 'aguardando') return;
        const env = entrega.envelope;
        let estado = 'erro';
        if (!window.GridRProjeto?.ativo()) {
            estado = 'abra_resumo';
            showToast('Abra a obra de destino no OrçaCore e reenvie as anotações.');
        } else if (String(env.projeto) !== localStorage.getItem('projeto_selecionado_codigo')) {
            estado = 'outro_projeto';
            showToast('O projeto selecionado difere do projeto enviado pelo nanoCAD.');
        } else {
            const recibo = 'orcacore:anotacoes:' + lote;
            if (sessionStorage.getItem(recibo) || adicionarObraAoProjeto(env.dados, true)) {
                sessionStorage.setItem(recibo, '1');
                estado = 'adicionado';
                window.GridRProjeto.marcarSujo(true);
                mostrarTabelasDoResumo('cabos', 'outros');
                updateCounters();
                updateHistoryUI();
                showToast('Anotações acrescentadas à obra aberta. Ctrl+Z desfaz.');
            }
        }
        await fetch(url + '/confirmar-anotacoes?estado=' + estado, { method: 'POST' });
    }

    async function receberObraDoPlugin(pend, anotacoes) {
        if (!pend) return;
        if (!localStorage.getItem('auth_token')) {
            if (anotacoes) return;
            window.location.href = '/login?next=' + encodeURIComponent(window.location.pathname + window.location.search);
            return;
        }
        try {
            if (!anotacoes) window.history.replaceState(null, '', '/resumo?plugin_pendente=' + encodeURIComponent(pend));
            const url = `/api/plugin/pendente/${encodeURIComponent(pend)}`;
            const r = await fetch(url);
            const d = await r.json().catch(() => ({}));
            if (!r.ok) { showToast(d.detail || `Não consegui receber o projeto do nanoCAD (HTTP ${r.status}).`); return; }
            if (d.modo === 'anotacoes') {
                await receberAnotacoesDoPlugin(url, pend, d);
                return;
            }
            if (anotacoes) return;
            const limparPendente = () => {
                try { window.history.replaceState(null, '', window.location.pathname); } catch (e) { /* navegador antigo */ }
            };
            const proj = d.projeto || {};
            const selecionarProjeto = () => {
                if (proj.codigo) localStorage.setItem('projeto_selecionado_codigo', proj.codigo);
                if (proj.nome) {
                    localStorage.setItem('projeto_selecionado', proj.nome);
                    const sel = document.getElementById('select-projeto');
                    if (sel) sel.value = proj.nome;
                }
            };
            if (d.caminho) {
                // a ponte do aplicativo (pywebview) chega logo depois da página: espera até ~10 s
                for (let k = 0; k < 40 && !(window.GridRProjeto && window.GridRProjeto.temPonte && window.GridRProjeto.temPonte()); k++) await new Promise(ok => setTimeout(ok, 250));
                if (!(window.GridRProjeto && window.GridRProjeto.temPonte && window.GridRProjeto.temPonte())) {
                    limparPendente();
                    showToast('O projeto do nanoCAD só abre no OrçaCore instalado. Arquivo: ' + d.caminho);
                    return;
                }
                const abriu = await window.GridRProjeto.abrirDoPlugin(async () => {
                    const resposta = await fetch(url + '/projeto');
                    const projeto = await resposta.json();
                    if (resposta.status !== 401) limparPendente();
                    return resposta.ok ? projeto : { ok: false, erro: projeto.detail };
                }, selecionarProjeto);
                if (abriu) mostrarTabelasDoResumo('cabos', 'outros');
                return;
            }
            limparPendente();
            const env = d.envelope || {};
            if (env.dados) {
                selecionarProjeto();
                restoreObraSnapshot(env.dados);
                mostrarTabelasDoResumo('cabos', 'outros');
                showToast('Obra do nanoCAD na tela (não foi gravada). Atualize o plugin para abrir como projeto.');
            }
        } catch (e) {
            showToast('Não consegui receber o projeto do nanoCAD: ' + e.message);
        }
    }
    // O desktop entrega na própria tela para preservar as edições e os diálogos do projeto.
    let recebendoPlugin = Promise.resolve();
    window.abrirPendenteNanoCAD = (pend, anotacoes = false) => {
        recebendoPlugin = recebendoPlugin.then(() => receberObraDoPlugin(pend, anotacoes));
    };
    window.abrirPendenteNanoCAD(new URLSearchParams(window.location.search).get('plugin_pendente'));

    (async function abrirObraDoLink() {
        const q = new URLSearchParams(window.location.search);
        const obraId = q.get('obra');
        if (!obraId) return;
        if (!localStorage.getItem('auth_token')) {
            window.location.href = '/login?next=' + encodeURIComponent(window.location.pathname + window.location.search);
            return;
        }
        const projeto = q.get('projeto') || '';
        const projetoNome = q.get('projeto_nome') || '';
        if (projeto) localStorage.setItem('projeto_selecionado_codigo', projeto);
        if (projetoNome) {
            localStorage.setItem('projeto_selecionado', projetoNome);
            const sel = document.getElementById('select-projeto');
            if (sel) sel.value = projetoNome;
        }
        try { history.replaceState(null, '', window.location.pathname); } catch (e) { /* navegador antigo: fica o link */ }
        try {
            const r = await fetch(`/api/obras/${encodeURIComponent(obraId)}?projeto=${encodeURIComponent(projeto)}`);
            if (!r.ok) { showToast(`Não consegui abrir a obra do nanoCAD (HTTP ${r.status}).`); return; }
            const obra = await r.json();
            const snap = typeof obra.dados_json === 'string' ? JSON.parse(obra.dados_json) : obra.dados_json;
            restoreObraSnapshot(snap);
            showToast(`Obra "${obra.nome}" aberta do nanoCAD. Ctrl+Z desfaz.`);
        } catch (e) {
            showToast('Não consegui abrir a obra do nanoCAD: ' + e.message);
        }
    })();

    /* TASK-154: PROJETO DE TRABALHO (.obra) — ganchos do Resumo para static/projeto.js (grava Cabos/Outros/Totalizadora no arquivo do projeto) */
    /* TASK-167: ligações com a tela (fora do trecho que os testes de paridade executam sem DOM) */
    (function ligarBalaoParametros() {
        let balao = null, timer = null, alvo = null;
        const esconder = () => { clearTimeout(timer); alvo = null; if (balao) balao.style.display = 'none'; };
        const mostrar = el => {
            let status = [];
            try { status = _statusDoAlvoDoBalao(el); } catch (x) { status = []; }
            if (alvo !== el) return;
            if (!status.length) { if (balao) balao.style.display = 'none'; return; }          // nada a mostrar: não deixa o balão de outra linha na tela
            if (!balao) { balao = document.createElement('div'); balao.className = 'par-balao'; document.body.appendChild(balao); }
            balao.innerHTML = _htmlBalaoParametros(status);
            const r = el.getBoundingClientRect();
            balao.style.display = 'block';
            balao.style.left = Math.max(8, Math.min(r.left, window.innerWidth - balao.offsetWidth - 8)) + 'px';
            balao.style.top = (r.bottom + balao.offsetHeight + 12 > window.innerHeight ? Math.max(8, r.top - balao.offsetHeight - 6) : r.bottom + 4) + 'px';
        };
        document.addEventListener('mouseover', e => {
            const el = e.target && e.target.closest ? e.target.closest('[data-par-tipo]') : null;
            if (!el || el === alvo) return;
            if (el.dataset.parTipo === 'outros' && window.AtivoBalao) return;                  // Outros: quem mostra (e edita) é o balão interativo (ativo_balao.js)
            esconder(); alvo = el;
            timer = setTimeout(() => mostrar(el), 250);
        });
        // Digitando no ativo (depois do "("): mostra já, sem esperar o mouse; some ao sair do campo ou fora do parêntese
        window.mostrarBalaoParametros = el => { clearTimeout(timer); alvo = el; mostrar(el); };
        window.esconderBalaoParametros = el => { if (!el || el === alvo) esconder(); };
        document.addEventListener('mouseout', e => { const el = e.target && e.target.closest ? e.target.closest('[data-par-tipo]') : null; if (el && el === alvo && !(e.relatedTarget && el.contains(e.relatedTarget))) esconder(); });
        ['mousedown', 'keydown', 'scroll'].forEach(ev => document.addEventListener(ev, esconder, true));
    })();

    /* Ocultar parâmetros ao sair da linha (preferência de tela, desligada por padrão): classe no body; o resto é CSS */
    const OCULTAR_PARAMS_CHAVE = 'ocultar_parametros_ao_sair';
    function ocultarParametrosLigado() { try { return localStorage.getItem(OCULTAR_PARAMS_CHAVE) === '1'; } catch (e) { return false; } }
    window.definirOcultarParametros = function (ligado) {
        try { localStorage.setItem(OCULTAR_PARAMS_CHAVE, ligado ? '1' : '0'); } catch (e) { /* vale só nesta sessão */ }
        document.body.classList.toggle('ocultar-parametros', !!ligado);
        if (window.AtivoBalao) window.AtivoBalao.fechar(false);
        try { renderTable('outros'); } catch (e) { /* tela ainda carregando */ }          // o campo do ativo passa a mostrar só os nomes (ou o texto inteiro)
    };
    (function ligarChaveOcultarParametros() {
        const chave = document.getElementById('cfg-ocultar-params');
        document.body.classList.toggle('ocultar-parametros', ocultarParametrosLigado());
        if (!chave) return;
        const pintar = () => chave.setAttribute('aria-checked', ocultarParametrosLigado() ? 'true' : 'false');
        pintar();
        chave.addEventListener('click', () => { window.definirOcultarParametros(!ocultarParametrosLigado()); pintar(); });         // botão (não checkbox): preferência de tela, não marca o projeto como alterado
    })();

    (function ligarChaveCoresParametros() {
        const chave = document.getElementById('cfg-colorir-params');
        document.body.classList.toggle('colorir-parametros', corParametrosLigada());
        if (!chave) return;
        const pintar = () => chave.setAttribute('aria-checked', corParametrosLigada() ? 'true' : 'false');
        pintar();
        chave.addEventListener('click', () => { window.definirCorParametros(!corParametrosLigada()); pintar(); });          // botão (não checkbox): é preferência de tela, não deve marcar o projeto como alterado
    })();

    (function ligarProjeto() {
        if (!window.GridRProjeto) return;
        const lista = k => tableStates[k].data.filter(Boolean);
        const sync = window.syncTotalizadora;
        if (typeof sync === 'function') {                                   // ir para a Totalizadora / recalculá-la também grava
            window.syncTotalizadora = async function () { const r = await sync.apply(this, arguments); GridRProjeto.marcarSujo(true); return r; };
        }
        document.querySelectorAll('input[name="view_mode"]').forEach(r => r.addEventListener('change', () => GridRProjeto.marcarSujo(true)));
        GridRProjeto.registrarResumo({
            coletar: () => ({ cabos: lista('cabos').map(deepClone), outros: lista('outros').map(deepClone), totalizadora: lista('totalizadora').map(deepClone) }),
            carregar: dados => restoreObraSnapshot(dados),
            mesclar: dados => adicionarObraAoProjeto(dados),
            vincularPorPoste: async () => {                                  // TASK-166: obra do plugin (com geometria) abre já vinculada pelo que chega em cada poste
                const cfg = await window.carregarRegrasVinculacaoResumo(localStorage.getItem('projeto_selecionado_codigo') || 'DEFAULT');
                if (!cfg || typeof VinculacaoRegras === 'undefined') return null;
                const r = VinculacaoRegras.vincularPorPoste(tableStates.cabos.data, tableStates.outros.data, cfg, _prefixoCaboVinculo, _codigoEstruturaVinculo, _codigosEstruturaVinculo, undefined, true);
                try { renderTable('outros'); atualizarResumoRedeUI(); } catch (e) { /* tela ainda carregando */ }
                return r;
            },
            temDados: () => GridRProjeto.temDados({ cabos: lista('cabos'), outros: lista('outros'), totalizadora: lista('totalizadora') })
        });
    })();

}); // end DOMContentLoaded

// ==========================================
// MODAL DE CONEXÕES
// ==========================================
const dadosCunha = [
    { tipo: "CAA 4 P/ C AA4", ativo_com: "C4", ativo_glv: "GLV2" },
    { tipo: "CAA 4 P/ C AA2", ativo_com: "C24", ativo_glv: "GLV2" },
    { tipo: "CAA 2 P/ C AA2", ativo_com: "C2", ativo_glv: "GLV2" },
    { tipo: "CAA 1/0 P/ C AA2", ativo_com: "C102", ativo_glv: "GLV10" },
    { tipo: "CAA 1/0 P/ C AA1/0", ativo_com: "C10", ativo_glv: "GLV10" },
    { tipo: "CAA 4/0 P/ CAA4", ativo_com: "C404", ativo_glv: "GLV40" },
    { tipo: "CAA 4/0 P/ CAA2", ativo_com: "C402", ativo_glv: "GLV40" },
    { tipo: "CAA 4/0 P/ CAA1/0", ativo_com: "C4010", ativo_glv: "GLV4010" },
    { tipo: "CAA 4/0 P/CAA4/0", ativo_com: "C40", ativo_glv: "GLV40" },
    { tipo: "CAA 336 P/ CAA4", ativo_com: "C3364", ativo_glv: "GLV336" },
    { tipo: "CAA 336 P/ CAA2", ativo_com: "C3362", ativo_glv: "GLV336" },
    { tipo: "CAA 336 P/ CAA1/0", ativo_com: "C33610", ativo_glv: "GLV336" },
    { tipo: "CAA 336 P/ CAA4/0", ativo_com: "C33640", ativo_glv: "GLV336" },
    { tipo: "CAA 336 P/ CAA336", ativo_com: "C336", ativo_glv: "GLV336" },
    { tipo: "CAA2 P/P50", ativo_com: "C502", ativo_glv: "GLV50" },
    { tipo: "P50 P/ P50", ativo_com: "C50", ativo_glv: "GLV50" },
    { tipo: "P50 P/ CAA2", ativo_com: "C502", ativo_glv: "GLV50" },
    { tipo: "P120 P/ P50", ativo_com: "C12050", ativo_glv: "GLV120" },
    { tipo: "P120 P/ CAA2", ativo_com: "C1202", ativo_glv: "GLV120" },
    { tipo: "P120 P/ P120", ativo_com: "C120", ativo_glv: "GLV120" },
    { tipo: "P185 P/ P50", ativo_com: "C18550", ativo_glv: "GLV185" },
    { tipo: "P185 P/ CAA2", ativo_com: "C1852", ativo_glv: "GLV185" },
    { tipo: "P185 P/ P120", ativo_com: "C185120", ativo_glv: "GLV185" },
    { tipo: "P185 P/ P185", ativo_com: "C185", ativo_glv: "GLV185" },
    { tipo: "aterramento temporario", ativo_com: "", ativo_glv: "" }
];

let conexoesRendered = false;

window.abrirModalConexoes = function() {
    document.getElementById('modal-conexoes').classList.remove('hidden');
    if (!conexoesRendered) {
        renderizarTabelaConexoes('tbody-conexoes-cunha', dadosCunha);
        // As outras ficam vazias por enquanto
        renderizarTabelaConexoes('tbody-conexoes-estrang', []);
        renderizarTabelaConexoes('tbody-conexoes-perf', []);
        conexoesRendered = true;
    }
    recalcularTotaisConexoes();
};

window.fecharModalConexoes = function() {
    document.getElementById('modal-conexoes').classList.add('hidden');
};

window.switchTabConexoes = function(tabName) {
    const tabs = ['cunha', 'estrang', 'perf'];
    tabs.forEach(t => {
        document.getElementById(`tab-con-${t}`).className = 'btn-secondary';
        document.getElementById(`tab-con-${t}`).style.background = '';
        document.getElementById(`panel-con-${t}`).style.display = 'none';
    });

    const activeTab = document.getElementById(`tab-con-${tabName}`);
    activeTab.className = 'btn-primary';
    activeTab.style.background = '#238636';
    document.getElementById(`panel-con-${tabName}`).style.display = 'block';
};

function renderizarTabelaConexoes(tbodyId, dados) {
    const tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    tbody.innerHTML = '';
    
    if (dados.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align: center; color: #8b949e; padding: 20px;">Sem dados por enquanto</td></tr>`;
        return;
    }

    dados.forEach((d, idx) => {
        const isAterr = d.tipo.toLowerCase().includes('aterramento');
        const tr = document.createElement('tr');
        
        // Define as cores de fundo especiais para aterramento temporário
        const styleMono = isAterr ? 'background-color: rgba(140, 227, 102, 0.4);' : '';
        const styleTri = isAterr ? 'background-color: rgba(255, 194, 32, 0.4);' : '';
        
        tr.innerHTML = `
            <td style="font-size: 0.8rem; padding: 4px 8px;">${d.tipo}</td>
            <td style="text-align: center; padding: 2px; ${styleMono}"><input type="number" min="0" class="modal-input con-mono" data-id="${tbodyId}-${idx}" data-ativo="${d.ativo_com}" style="width: 50px; text-align: center; padding: 2px;"></td>
            <td style="text-align: center; padding: 2px; ${styleTri}"><input type="number" min="0" class="modal-input con-tri" data-id="${tbodyId}-${idx}" data-ativo="${d.ativo_com}" style="width: 50px; text-align: center; padding: 2px;"></td>
            <td style="text-align: center; padding: 2px;"><input type="number" min="0" class="modal-input con-glv" data-id="${tbodyId}-${idx}" data-ativo="${d.ativo_glv}" style="width: 50px; text-align: center; padding: 2px;"></td>
        `;
        tbody.appendChild(tr);
    });

    tbody.querySelectorAll('input[type="number"]').forEach(input => {
        input.addEventListener('input', recalcularTotaisConexoes);
    });
}

function recalcularTotaisConexoes() {
    const totais = acumularTotaisConexoes();
    const painel = document.getElementById('conexoes-totalizador');
    
    let summary = [];
    let count = 0;
    
    for (const [ativo, qtd] of Object.entries(totais)) {
        if (qtd > 0) {
            summary.push(`<b>${ativo}</b>: ${qtd}`);
            count += qtd;
        }
    }
    
    if (count === 0) {
        painel.innerHTML = 'Nenhum item selecionado.';
    } else {
        painel.innerHTML = `Total a gerar: ${summary.join(' | ')}`;
    }
}

function acumularTotaisConexoes() {
    const totais = {};
    const add = (ativo, qtd) => {
        if (!ativo || qtd <= 0) return;
        if (!totais[ativo]) totais[ativo] = 0;
        totais[ativo] += qtd;
    };

    const container = document.getElementById('modal-conexoes');
    
    // MONO (Qtd * 1)
    container.querySelectorAll('.con-mono').forEach(input => {
        const val = parseInt(input.value) || 0;
        const ativo = input.getAttribute('data-ativo');
        add(ativo, val * 1);
    });

    // TRI (Qtd * 3)
    container.querySelectorAll('.con-tri').forEach(input => {
        const val = parseInt(input.value) || 0;
        const ativo = input.getAttribute('data-ativo');
        add(ativo, val * 3);
    });

    // GLV (Qtd * 1)
    container.querySelectorAll('.con-glv').forEach(input => {
        const val = parseInt(input.value) || 0;
        const ativo = input.getAttribute('data-ativo');
        add(ativo, val * 1);
    });
    
    return totais;
}

window.adicionarConexoesAosOutros = function() {
    const totais = acumularTotaisConexoes();
    let adicionados = 0;

    for (const [ativo, qtd] of Object.entries(totais)) {
        if (qtd > 0) {
            // Adiciona na Tabela Outros no formato "qtd-ativo"
            tableStates.outros.data.push({
                entidade: '',
                operacao: 'I',
                ativo: qtd + "-" + ativo
            });
            adicionados++;
        }
    }

    if (adicionados > 0) {
        renderOutrosTable();
        // Limpar inputs após adicionar
        document.querySelectorAll('#modal-conexoes input[type="number"]').forEach(i => i.value = '');
        recalcularTotaisConexoes();
        
        // Verifica se a tabela Outros está visível, se não, muda a view para PADRÃO para o usuário ver o resultado
        const radioPadrao = document.querySelector('input[name="view_mode"][value="padrao"]');
        if (radioPadrao && !radioPadrao.checked) {
            radioPadrao.checked = true;
            radioPadrao.dispatchEvent(new Event('change'));
        }
    }
    
    fecharModalConexoes();
};
