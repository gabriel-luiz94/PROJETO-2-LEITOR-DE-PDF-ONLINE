/**
 * resumo.js  –  Aba "Resumo do Orçamento"
 * Tabelas de Cabos e Demais Ativos lado a lado.
 * Colunas: Entidade (select), Operação (select validado), Ativo (input + autocomplete).
 * Funcionalidades: Add linha, Excluir linha, Copiar, Undo (Ctrl+Z), Redo (Ctrl+Y).
 */
document.addEventListener('DOMContentLoaded', () => {

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

        function atualizar(online) {
            if (online) {
                selo.style.display = 'none';
                return;
            }
            selo.style.display = 'inline-flex';
            selo.style.background = 'rgba(248,81,73,0.15)';
            selo.style.color = '#f85149';
            bola.style.background = '#f85149';
            texto.textContent = 'Sem conexão com a nuvem — trabalhando localmente';
        }

        try {
            const wsProtocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
            const socket = new WebSocket(`${wsProtocol}//${window.location.host}/ws`);
            socket.onmessage = (event) => {
                try {
                    const data = JSON.parse(event.data);
                    if (data.type === 'connectivity') atualizar(data.status === 'online');
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

    const tableStates = {
        cabos: { bodyId: 'body-cabos', data: [] },
        outros: { bodyId: 'body-outros', data: [] },
        totalizadora: { bodyId: 'body-totalizadora', data: [] },
        regras: { bodyId: 'body-regras', data: [] }
    };

    /* ─── Column filters ─── */
    // Selections: type → field → Set of selected values
    const filterSelections = {
        cabos:  { entidade: new Set(), operacao: new Set(), ativo: new Set() },
        outros: { entidade: new Set(), operacao: new Set(), ativo: new Set() }
    };
    const FILTER_FIELDS = ['entidade', 'operacao', 'ativo'];
    const FILTER_CONTAINERS = {
        cabos:  { entidade: 'rfc-cabos-entidade', operacao: 'rfc-cabos-operacao', ativo: 'rfc-cabos-ativo' },
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
    let ctxRow = null, ctxType = null;

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
    function renderTable(type) {
        const state = tableStates[type];
        const body = document.getElementById(state.bodyId);
        body.innerHTML = '';

        state.data.forEach((row, idx) => {
            const tr = document.createElement('tr');
            tr.dataset.index = idx;
            tr.dataset.type = type;

            /* ── Entidade ── */
            const tdEnt = document.createElement('td');
            const selEnt = document.createElement('select');
            selEnt.className = 'sel-entidade';
            selEnt.dataset.field = 'entidade';
            ENTIDADES.forEach(e => {
                const opt = document.createElement('option');
                opt.value = e; opt.textContent = e;
                if (e === row.entidade) opt.selected = true;
                selEnt.appendChild(opt);
            });
            selEnt.addEventListener('change', () => {
                const old = row.entidade;
                if (old !== selEnt.value) {
                    pushHistory();  // snapshot ANTES da mudança
                    row.entidade = selEnt.value;
                    refreshAllFilters(type);
                    atualizarResumoRedeUI();
                }
            });
            tdEnt.appendChild(selEnt);
            tr.appendChild(tdEnt);

            /* ── Operação ── */
            const tdOp = document.createElement('td');
            const selOp = document.createElement('select');
            selOp.className = 'sel-operacao';
            selOp.dataset.field = 'operacao';
            OPERACOES.forEach(o => {
                const opt = document.createElement('option');
                opt.value = o; opt.textContent = o;
                if (o === row.operacao) opt.selected = true;
                selOp.appendChild(opt);
            });
            selOp.addEventListener('change', () => {
                const old = row.operacao;
                if (old !== selOp.value) {
                    pushHistory();  // snapshot ANTES da mudança
                    row.operacao = selOp.value;
                    refreshAllFilters(type);
                    atualizarResumoRedeUI();
                }
            });
            tdOp.appendChild(selOp);
            tr.appendChild(tdOp);

            /* ── Ativo ── */
            const tdAt = document.createElement('td');
            tdAt.style.position = 'relative';
            const inpAt = document.createElement('input');
            inpAt.type = 'text';
            inpAt.className = 'inp-ativo';
            inpAt.dataset.field = 'ativo';
            inpAt.value = row.ativo;
            inpAt.autocomplete = 'off';
            inpAt.setAttribute('list', '');  // desabilita datalist nativo, usamos o nosso
            let prevAtivo = row.ativo;

            inpAt.addEventListener('input', () => {
                const oldVal = inpAt.value;
                const upperVal = oldVal.toUpperCase();
                if (oldVal !== upperVal) {
                    const cursor = inpAt.selectionStart;
                    inpAt.value = upperVal;
                    if (cursor !== null) inpAt.setSelectionRange(cursor, cursor);
                }
                row.ativo = inpAt.value;
                showAutocomplete(inpAt, type);
            });
            inpAt.addEventListener('keydown', e => {
                if (e.key === 'Delete' && inpAt.value.trim() === '') {
                    e.preventDefault();
                    deleteRow(type, idx);
                    return;
                }
                const consumed = handleAcKeydown(e, inpAt, type);
                if (e.key === 'Enter' && !consumed) {
                    e.preventDefault();
                    addRow(type, idx);
                }
            });
            inpAt.addEventListener('blur', () => {
                setTimeout(() => {
                    hideAutocomplete();
                    if (inpAt.value !== prevAtivo) {
                        const snapAtivo = prevAtivo;  // guarda o valor antes
                        prevAtivo = inpAt.value;
                        // Restaura temporariamente o valor antigo para capturar o snapshot correto
                        row.ativo = snapAtivo;
                        pushHistory();  // snapshot com valor ANTES
                        row.ativo = inpAt.value;  // aplica o novo valor
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
                    }
                }, 150);
            });
            inpAt.addEventListener('focus', () => {
                prevAtivo = inpAt.value;
                if (inpAt.value.length === 0) showAutocomplete(inpAt, type);
            });

            tdAt.appendChild(inpAt);
            tr.appendChild(tdAt);

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
                tr.appendChild(tdQtd);
            }

            /* ── Delete button ── */
            const tdDel = document.createElement('td');
            tdDel.className = 'col-del';
            const btnDel = document.createElement('button');
            btnDel.className = 'btn-row-del';
            btnDel.title = 'Excluir linha';
            btnDel.innerHTML = `<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/></svg>`;
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
                showCtxMenu(e.clientX, e.clientY);
            });

            body.appendChild(tr);
        });

        updateCounters();
        atualizarResumoRedeUI();

        // Após render, aplica filtros ativos sem rebuildar dropdowns
        applyFilters(type);
    }

    /* ═══════════════════════════════════════
       CRUD
    ═══════════════════════════════════════ */
    function deleteRow(type, idx) {
        pushHistory();  // snapshot ANTES da exclusão
        tableStates[type].data.splice(idx, 1);
        // TASK-032: ao remover uma linha de Outros, os vínculos cabo<->estrutura são baseados em
        // índice — reajusta as referências pra não apontarem para a linha errada depois do shift.
        if (type === 'outros') _reajustarVinculosAposRemoverOutros(idx);
        renderTable(type);
        buildAtivoSets(); buildDataLists();
        refreshAllFilters(type);
    }

    function _reajustarVinculosAposRemoverOutros(idxRemovido) {
        tableStates.cabos.data.forEach(cabo => {
            if (!Array.isArray(cabo.vinculoEstruturas)) return;
            cabo.vinculoEstruturas = cabo.vinculoEstruturas
                .filter(i => i !== idxRemovido)
                .map(i => i > idxRemovido ? i - 1 : i);
        });
    }

    function addRow(type, afterIdx = -1) {
        pushHistory();  // snapshot ANTES da inserção
        let op = 'I';
        if (afterIdx >= 0 && tableStates[type].data[afterIdx]) {
            op = tableStates[type].data[afterIdx].operacao;
        } else if (tableStates[type].data.length > 0 && afterIdx === -1) {
            op = tableStates[type].data[tableStates[type].data.length - 1].operacao;
        }
        
        const newRow = { entidade: '0', operacao: op, ativo: '' };
        if (afterIdx === -1) tableStates[type].data.push(newRow);
        else tableStates[type].data.splice(afterIdx + 1, 0, newRow);
        renderTable(type);
        // Foca no input Ativo da nova linha
        const body = document.getElementById(tableStates[type].bodyId);
        const target = afterIdx === -1 ? body.lastElementChild : body.children[afterIdx + 1];
        if (target) {
            const inp = target.querySelector('.inp-ativo');
            if (inp) { setTimeout(() => inp.focus(), 30); }
        }
        refreshAllFilters(type);
    }

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
        FILTER_FIELDS.forEach(field => {
            buildFilterDropdown(FILTER_CONTAINERS[type][field], type, field);
        });
    }

    function applyFilters(type) {
        const state = tableStates[type];
        const body  = document.getElementById(state.bodyId);
        if (!body) return;
        const sels = filterSelections[type];
        Array.from(body.children).forEach(tr => {
            const idx = parseInt(tr.dataset.index);
            if (isNaN(idx)) return;
            const row = state.data[idx];
            if (!row) return;
            let visible = true;
            for (const field of FILTER_FIELDS) {
                const sel = sels[field];
                if (sel.size > 0 && !sel.has(row[field] || '')) {
                    visible = false; break;
                }
            }
            tr.style.display = visible ? '' : 'none';
        });
        updateCounters();
    }


    function deleteSelectedOrLast(type) {
        const body = document.getElementById(tableStates[type].bodyId);
        const selected = Array.from(body.querySelectorAll('tr.row-selected'));
        if (selected.length > 0) {
            pushHistory();  // snapshot ANTES da exclusão em lote
            // Remove de trás pra frente para não deslocar índices
            const idxs = selected.map(tr => parseInt(tr.dataset.index)).sort((a,b) => b-a);
            idxs.forEach(i => tableStates[type].data.splice(i, 1));
            renderTable(type);
            buildAtivoSets(); buildDataLists();
            refreshAllFilters(type);
        } else {
            // Remove a última linha
            if (tableStates[type].data.length > 0) {
                pushHistory();  // snapshot ANTES da exclusão
                tableStates[type].data.pop();
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
    function showCtxMenu(x, y) {
        ctxMenu.style.left = Math.min(x, window.innerWidth - 180) + 'px';
        ctxMenu.style.top  = Math.min(y, window.innerHeight - 140) + 'px';
        ctxMenu.classList.add('visible');
    }
    function hideCtxMenu() { ctxMenu.classList.remove('visible'); ctxRow = null; ctxType = null; }
    document.addEventListener('mousedown', e => { if (!ctxMenu.contains(e.target)) hideCtxMenu(); });

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
    document.getElementById('ctx-del').addEventListener('click', () => {
        if (!ctxRow || !ctxType) return;
        deleteRow(ctxType, parseInt(ctxRow.dataset.index));
        hideCtxMenu();
    });

    /* Seleção de linha com clique */
    document.addEventListener('click', e => {
        const tr = e.target.closest('tr[data-index]');
        if (!tr) return;
        if (e.target.tagName === 'SELECT' || e.target.tagName === 'INPUT' || e.target.closest('.btn-row-del')) return;
        if (!e.ctrlKey && !e.shiftKey) {
            document.querySelectorAll('tr.row-selected').forEach(r => r.classList.remove('row-selected'));
        }
        tr.classList.toggle('row-selected');
    });

    /* ═══════════════════════════════════════
       AUTOCOMPLETE
    ═══════════════════════════════════════ */
    function buildAtivoSets() {
        ativoSets.cabos.clear();
        ativoSets.outros.clear();
        tableStates.cabos.data.forEach(r => { if (r.ativo) ativoSets.cabos.add(r.ativo.trim()); });
        tableStates.outros.data.forEach(r => { if (r.ativo) ativoSets.outros.add(r.ativo.trim()); });
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
        const q = input.value.trim().toUpperCase();
        const pool = Array.from(type === 'cabos' ? ativoSets.cabos : ativoSets.outros);
        acItems = q ? pool.filter(v => v.toUpperCase().includes(q) && v.toUpperCase() !== q) : pool.slice(0, 20);
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
            if (e.key === 'ArrowDown') { showAutocomplete(input, type); e.preventDefault(); }
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
    const mesmoEstado = (a, b) => JSON.stringify(a) === JSON.stringify(b);

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
        renderTable('cabos');
        renderTable('outros');
        buildAtivoSets();
        buildDataLists();
        refreshAllFilters('cabos');
        refreshAllFilters('outros');
        updateHistoryUI();
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
            const visible = body ? Array.from(body.children).filter(tr => tr.style.display !== 'none').length : tableStates[type].data.length;
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
        txt = txt.replace(/-/g, ' ').replace(/\*/g, '-');
        const tokens = txt.split(/\s+/).filter(Boolean);
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
     * Extrai o prefixo do ativo e o comprimento (metros) de uma linha da tabela Cabos, no formato
     * "ATIVO FASE COMPRIMENTO". Mesma normalização de calcularQtdAtivos() (linha ~903): funde
     * prefixos partidos por espaço ("CAA 2" → "CAA2") antes de separar os tokens, senão o
     * comprimento e a fase seriam lidos errado nesse formato. O comprimento retornado é o valor
     * bruto (último token) — nunca multiplicado por qtdAtivos/fases (RULES.md Regra 8, confirmado
     * pelo usuário em TASK-008-29-09-2026.md).
     */
    function extrairPrefixoEComprimentoCabo(ativoRaw) {
        if (!ativoRaw || !ativoRaw.trim()) return null;
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

    /**
     * Calcula as 4 contagens da "Resumo da rede": postes, rede de média, rede de baixa e
     * equipamentos — todas considerando só linhas com operação "I" (instalando).
     */
    function calcularResumoRede() {
        const resumo = { postes: 0, redeMedia: 0, redeBaixa: 0, equipamentos: 0 };

        tableStates.outros.data.forEach(row => {
            if ((row.operacao || '').toUpperCase() !== 'I') return;
            extrairParesQtdAtivoOutros(row.ativo).forEach(({ ativo, qtd }) => {
                if (/^(DT|CV)/i.test(ativo)) {
                    resumo.postes += qtd;
                } else if (/^(TR|CFU|CFA)/i.test(ativo)) {
                    resumo.equipamentos += qtd;
                }
            });
        });

        tableStates.cabos.data.forEach(row => {
            if ((row.operacao || '').toUpperCase() !== 'I') return;
            const info = extrairPrefixoEComprimentoCabo(row.ativo);
            if (!info) return;
            if (/^(CAA|CAL|P)/i.test(info.prefixo)) {
                resumo.redeMedia += info.comprimento;
            } else if (/^(M2X|M3X)/i.test(info.prefixo)) {
                resumo.redeBaixa += info.comprimento;
            }
        });

        return resumo;
    }

    function formatarNumeroResumoRede(n) {
        const arredondado = Math.round(n * 100) / 100;
        return arredondado % 1 === 0 ? String(arredondado) : arredondado.toFixed(2).replace('.', ',');
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

    window.__regrasVinculacao = [];
    window.carregarRegrasVinculacaoResumo = async function (projetoCodigo) {
        if (!projetoCodigo) return;
        try {
            const res = await fetch(`/api/regras-vinculacao?projeto_codigo=${encodeURIComponent(projetoCodigo)}`);
            const data = res.ok ? await res.json() : { regras: [] };
            window.__regrasVinculacao = data.regras || [];
        } catch (e) {
            console.error('Erro ao carregar regras de vinculação:', e);
        }
    };

    /* ── TASK-042: editor de regras_vinculacao (tipo/qtd/compatibilidade) dentro do modal ── */
    window.toggleVincRegras = function () {
        const content = document.getElementById('vinc-regras-content');
        const icon = document.getElementById('vinc-regras-toggle-icon');
        if (!content || !icon) return;
        if (content.classList.contains('hidden')) {
            content.classList.remove('hidden');
            icon.textContent = '▲ Ocultar';
            renderRegrasVinculacaoTable();
        } else {
            content.classList.add('hidden');
            icon.textContent = '▼ Mostrar';
        }
    };

    window.renderRegrasVinculacaoTable = function () {
        const tbody = document.getElementById('body-vinc-regras');
        if (!tbody) return;
        const regras = window.__regrasVinculacao || [];
        tbody.innerHTML = '';
        regras.forEach((r, index) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><input type="text" class="tot-input" maxlength="1" style="width:50px; text-align:center;" value="${_escapeHtmlVinculacao(r.tipo_estrutura || '')}" onchange="updateRegraVinculacaoRow(${index}, 'tipo_estrutura', this.value)"></td>
                <td><input type="number" min="1" step="1" class="tot-input" style="width:80px;" value="${r.qtd_cabos != null ? r.qtd_cabos : ''}" onchange="updateRegraVinculacaoRow(${index}, 'qtd_cabos', this.value)"></td>
                <td>
                    <select class="tot-input" onchange="updateRegraVinculacaoRow(${index}, 'compatibilidade', this.value)">
                        <option value="LIVRE" ${r.compatibilidade === 'LIVRE' ? 'selected' : ''}>LIVRE</option>
                        <option value="MESMO_TIPO_FASE_OPERACAO" ${r.compatibilidade === 'MESMO_TIPO_FASE_OPERACAO' ? 'selected' : ''}>MESMO TIPO/FASE/OPERAÇÃO</option>
                    </select>
                </td>
                <td><input type="text" class="tot-input" value="${_escapeHtmlVinculacao(r.descricao || '')}" onchange="updateRegraVinculacaoRow(${index}, 'descricao', this.value)"></td>
                <td style="text-align:center;">
                    <button class="btn-primary" style="background-color: transparent; border: none; color: #f85149; padding: 4px;" onclick="excluirRegraVinculacao(${index})" title="Excluir">
                        <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"></path></svg>
                    </button>
                </td>
            `;
            tbody.appendChild(tr);
        });
    };

    window.updateRegraVinculacaoRow = function (index, field, value) {
        const regras = window.__regrasVinculacao || [];
        if (!regras[index]) return;
        if (field === 'tipo_estrutura') {
            value = (value || '').trim().replace(/[^0-9]/g, '').slice(0, 1);
        } else if (field === 'qtd_cabos') {
            value = parseInt(value, 10);
            if (isNaN(value) || value < 1) value = 1;
        }
        regras[index][field] = value;
    };

    window.adicionarRegraVinculacao = function () {
        if (!window.__regrasVinculacao) window.__regrasVinculacao = [];
        window.__regrasVinculacao.push({ tipo_estrutura: '', qtd_cabos: 1, compatibilidade: 'LIVRE', descricao: '' });
        renderRegrasVinculacaoTable();
    };

    window.excluirRegraVinculacao = function (index) {
        if (!window.__regrasVinculacao) return;
        window.__regrasVinculacao.splice(index, 1);
        renderRegrasVinculacaoTable();
    };

    window.salvarRegrasVinculacaoNuvem = async function () {
        const projCode = localStorage.getItem('projeto_selecionado_codigo') || 'DEFAULT';
        try {
            const res = await fetch('/api/regras-vinculacao', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ projeto_codigo: projCode, regras: window.__regrasVinculacao || [] })
            });
            if (res.ok) {
                showToast('✓ Regras de vinculação salvas!');
            } else {
                const data = await res.json().catch(() => ({}));
                const erros = data && data.detail && data.detail.erros;
                showToast(Array.isArray(erros) ? erros.join(' ') : ((data && data.detail) || 'Erro ao salvar regras de vinculação.'));
            }
        } catch (e) {
            showToast('Erro de conexão ao salvar regras de vinculação.');
        }
    };

    function _estruturasDisponiveisVinculo() {
        return tableStates.outros.data
            .map((r, i) => ({ r, i }))
            .filter(x => x.r.entidade === 'ESTRUTURA' || x.r.entidade === 'POSTE');
    }

    function _qtdEstruturasParaOperacaoVinculo(operacao) {
        const op = (operacao || '').toUpperCase();
        return (op === 'M' || op === '*M') ? 1 : 2;
    }

    function _codigoEstruturaVinculo(ativoEstrutura) {
        // TASK-036: o ativo de uma linha Outros é "<qtd>-<código>" (ex.: "1-U4"), nunca o código puro.
        const texto = (ativoEstrutura || '').trim();
        const semQtd = texto.match(/^[*-]?[0-9]+(?:[.,][0-9]+)?[Xx-](.+)$/);
        return semQtd ? semQtd[1] : texto;
    }

    function _tipoEstruturaVinculo(ativoEstrutura) {
        // Pegar o primeiro dígito da string toda pegava o dígito da QUANTIDADE, não o do tipo —
        // isola o código (depois do prefixo de quantidade) antes de procurar o dígito do tipo.
        const m = _codigoEstruturaVinculo(ativoEstrutura).match(/\d/);
        return m ? m[0] : null;
    }

    function _regraVinculacaoParaTipo(tipo) {
        const regras = window.__regrasVinculacao || [];
        return regras.find(r => r.tipo_estrutura === tipo) || null;
    }

    function _prefixoCaboVinculo(ativo) {
        const info = extrairPrefixoEComprimentoCabo(ativo);
        return info ? info.prefixo : null;
    }

    window.abrirModalVinculacao = function () {
        const modal = document.getElementById('modal-vinculacao');
        if (!modal) return;
        modal.classList.remove('hidden');
        const resultadoEl = document.getElementById('vinculacao-auto-resultado');
        if (resultadoEl) resultadoEl.textContent = '';
        _renderVinculacaoLista();
    };

    function _renderVinculacaoLista() {
        const container = document.getElementById('vinculacao-lista');
        if (!container) return;
        const estruturas = _estruturasDisponiveisVinculo();

        if (!tableStates.cabos.data.length) {
            container.innerHTML = '<div style="color:#8b949e; padding:10px;">Nenhum cabo na tabela.</div>';
            return;
        }
        if (!estruturas.length) {
            container.innerHTML = '<div style="color:#8b949e; padding:10px;">Nenhuma estrutura/poste na tabela Outros ainda.</div>';
            return;
        }

        let html = '';
        tableStates.cabos.data.forEach((cabo, idxCabo) => {
            const qtd = _qtdEstruturasParaOperacaoVinculo(cabo.operacao);
            const vinculo = Array.isArray(cabo.vinculoEstruturas) ? cabo.vinculoEstruturas : [];
            html += `<div style="display:flex; align-items:center; gap:10px; padding:8px 6px; border-bottom:1px solid #21262d;">
                <div style="flex:1; min-width:0;">
                    <div style="font-weight:600;">${_escapeHtmlVinculacao(cabo.ativo || '(sem ativo)')}</div>
                    <div style="font-size:0.72rem; color:#8b949e;">operação ${_escapeHtmlVinculacao(cabo.operacao || '')} — exige ${qtd} estrutura(s)</div>
                </div>
                <div style="display:flex; gap:6px; flex-wrap:wrap;">`;
            for (let slot = 0; slot < qtd; slot++) {
                const valorAtual = vinculo[slot];
                const valido = valorAtual !== undefined && estruturas.some(e => e.i === valorAtual);
                html += `<select class="vinculacao-select" data-cabo="${idxCabo}" data-slot="${slot}" onchange="_atualizarVinculoSelect(this)" style="background:#0d1117; border:1px solid ${valorAtual !== undefined && !valido ? '#f85149' : '#30363d'}; border-radius:6px; color:white; padding:5px 7px; font-size:0.78rem; min-width:150px;">
                    <option value="">— selecione —</option>`;
                estruturas.forEach(e => {
                    const sel = valido && valorAtual === e.i ? 'selected' : '';
                    html += `<option value="${e.i}" ${sel}>${_escapeHtmlVinculacao(e.r.ativo || '(sem ativo)')}</option>`;
                });
                html += `</select>`;
            }
            html += `</div></div>`;
        });
        container.innerHTML = html;
    }

    window._atualizarVinculoSelect = function (selectEl) {
        const idxCabo = parseInt(selectEl.dataset.cabo, 10);
        const slot = parseInt(selectEl.dataset.slot, 10);
        const valor = selectEl.value === '' ? undefined : parseInt(selectEl.value, 10);

        const cabo = tableStates.cabos.data[idxCabo];
        if (!cabo) return;
        pushHistory();  // snapshot ANTES da mudança
        const vinculo = Array.isArray(cabo.vinculoEstruturas) ? cabo.vinculoEstruturas.slice() : [];
        while (vinculo.length <= slot) vinculo.push(undefined);
        vinculo[slot] = valor;
        cabo.vinculoEstruturas = vinculo.filter(v => v !== undefined);
        _renderVinculacaoLista();
    };

    /**
     * Vínculo automático por coordenada (opcional, TASK-032) — usa _x/_y (DXF sempre; PDF a
     * partir desta tarefa) para casar cada cabo com as estruturas mais próximas, respeitando a
     * quantidade/compatibilidade da regra do tipo de cada estrutura. Nunca sobrescreve um cabo
     * que já tenha vínculo (manual ou de uma passada anterior), e nunca "adivinha" quando não há
     * combinação válida — o cabo fica sem vínculo, destacado para correção manual.
     */
    window.tentarVincularAutomaticamente = function () {
        const resultadoEl = document.getElementById('vinculacao-auto-resultado');
        const estruturas = _estruturasDisponiveisVinculo();
        if (!estruturas.length) {
            if (resultadoEl) resultadoEl.textContent = 'Nenhuma estrutura/poste disponível para vincular.';
            return;
        }

        // Ocupação atual (vínculos manuais/já existentes contam para a capacidade de cada estrutura).
        const ocupacao = {};
        estruturas.forEach(e => { ocupacao[e.i] = []; });
        tableStates.cabos.data.forEach((cabo, idxCabo) => {
            (cabo.vinculoEstruturas || []).forEach(idxEst => {
                if (ocupacao[idxEst]) ocupacao[idxEst].push(idxCabo);
            });
        });

        let vinculados = 0, jaTinham = 0, semCandidato = 0, semCoordenada = 0;

        tableStates.cabos.data.forEach((cabo, idxCabo) => {
            if (Array.isArray(cabo.vinculoEstruturas) && cabo.vinculoEstruturas.length > 0) {
                jaTinham++;
                return;
            }
            if (cabo._x === undefined || cabo._y === undefined) {
                semCoordenada++;
                return;
            }

            const qtdNecessaria = _qtdEstruturasParaOperacaoVinculo(cabo.operacao);
            const prefixoCabo = _prefixoCaboVinculo(cabo.ativo);

            const candidatas = estruturas
                .filter(e => e.r._x !== undefined && e.r._y !== undefined)
                .map(e => ({ r: e.r, i: e.i, dist: Math.hypot(e.r._x - cabo._x, e.r._y - cabo._y) }))
                .sort((a, b) => a.dist - b.dist);

            const escolhidas = [];
            for (const cand of candidatas) {
                const tipo = _tipoEstruturaVinculo(cand.r.ativo);
                const regra = _regraVinculacaoParaTipo(tipo);
                const capacidade = regra ? regra.qtd_cabos : qtdNecessaria;
                const ocupados = ocupacao[cand.i] || [];
                if (ocupados.length + 1 > capacidade) continue;

                if (regra && regra.compatibilidade === 'MESMO_TIPO_FASE_OPERACAO' && ocupados.length > 0) {
                    const outroCabo = tableStates.cabos.data[ocupados[0]];
                    const prefixoOutro = _prefixoCaboVinculo(outroCabo.ativo);
                    if (prefixoOutro !== prefixoCabo || (outroCabo.operacao || '').toUpperCase() !== (cabo.operacao || '').toUpperCase()) {
                        continue;
                    }
                }
                escolhidas.push(cand);
                if (escolhidas.length === qtdNecessaria) break;
            }

            if (escolhidas.length === qtdNecessaria) {
                cabo.vinculoEstruturas = escolhidas.map(e => e.i);
                escolhidas.forEach(e => { ocupacao[e.i].push(idxCabo); });
                vinculados++;
            } else {
                semCandidato++;
            }
        });

        if (vinculados > 0) pushHistory();
        _renderVinculacaoLista();
        if (resultadoEl) {
            resultadoEl.textContent = `${vinculados} cabo(s) vinculado(s) automaticamente · ${jaTinham} já tinham vínculo (preservado) · ` +
                `${semCandidato + semCoordenada} ficaram sem vínculo (sem combinação confiável ou sem coordenada).`;
        }
    };

    /**
     * TASK-033: acha estruturas/postes sem o vínculo certo pro seu tipo (regras de vinculação,
     * TASK-032). Achados no mesmo formato de /api/validacao/planilhas, para aparecerem juntos no
     * modal de Validação — sem round-trip ao backend, já que o vínculo só existe no frontend.
     */
    function avaliarVinculacaoLocal() {
        const achados = [];
        const estruturas = _estruturasDisponiveisVinculo();
        if (!estruturas.length) return achados;

        const ocupacao = {};
        estruturas.forEach(e => { ocupacao[e.i] = []; });
        tableStates.cabos.data.forEach((cabo, idxCabo) => {
            (cabo.vinculoEstruturas || []).forEach(idxEst => {
                if (ocupacao[idxEst]) ocupacao[idxEst].push(idxCabo);
            });
        });

        // Exceção (só para zero vínculos): existe cabo M/*M ou ativo RETCA em qualquer lugar do projeto?
        const existeCaboMantendo = tableStates.cabos.data.some(c => {
            const op = (c.operacao || '').toUpperCase();
            return op === 'M' || op === '*M';
        });
        const existeRetca = tableStates.outros.data.some(r => /\bRETCA\b/i.test(r.ativo || ''));
        const temExcecaoZero = existeCaboMantendo || existeRetca;

        estruturas.forEach(e => {
            const ocupados = ocupacao[e.i] || [];
            const tipo = _tipoEstruturaVinculo(e.r.ativo);
            const regra = _regraVinculacaoParaTipo(tipo);
            if (!regra) return;  // sem regra cadastrada pro tipo: não dá pra avaliar

            if (ocupados.length === 0) {
                if (temExcecaoZero) return;
                achados.push({
                    linha_id: `OUTROS-${e.i}`, tabela: 'outros', regra_id: 'VINCULO-ESTRUTURA',
                    severidade: 'aviso',
                    mensagem: `Estrutura "${e.r.ativo}" sem nenhum cabo vinculado.`,
                    explicacao: `Tipo ${tipo} exige ${regra.qtd_cabos} cabo(s) vinculado(s); nenhum encontrado, e não há cabo M/*M nem RETCA no projeto para justificar.`
                });
                return;
            }

            if (ocupados.length !== regra.qtd_cabos) {
                achados.push({
                    linha_id: `OUTROS-${e.i}`, tabela: 'outros', regra_id: 'VINCULO-ESTRUTURA',
                    severidade: 'aviso',
                    mensagem: `Estrutura "${e.r.ativo}" com ${ocupados.length} cabo(s) vinculado(s), esperado ${regra.qtd_cabos}.`,
                    explicacao: `Tipo ${tipo} exige exatamente ${regra.qtd_cabos} cabo(s) vinculado(s).`
                });
                return;
            }

            if (regra.compatibilidade === 'MESMO_TIPO_FASE_OPERACAO' && ocupados.length > 1) {
                const primeiro = tableStates.cabos.data[ocupados[0]];
                const prefixoRef = _prefixoCaboVinculo(primeiro.ativo);
                const opRef = (primeiro.operacao || '').toUpperCase();
                const incompat = ocupados.slice(1).some(idx => {
                    const c = tableStates.cabos.data[idx];
                    return _prefixoCaboVinculo(c.ativo) !== prefixoRef || (c.operacao || '').toUpperCase() !== opRef;
                });
                if (incompat) {
                    achados.push({
                        linha_id: `OUTROS-${e.i}`, tabela: 'outros', regra_id: 'VINCULO-ESTRUTURA',
                        severidade: 'aviso',
                        mensagem: `Estrutura "${e.r.ativo}" com cabos vinculados incompatíveis entre si.`,
                        explicacao: `Tipo ${tipo} exige que os cabos vinculados sejam do mesmo tipo, fase e operação.`
                    });
                }
            }
        });

        return achados;
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
        // TASK-032: coordenada de origem (presente em obj._raw na primeira montagem da linha, e
        // em obj diretamente depois) e vínculo cabo<->estrutura — preservados através de
        // undo/redo e salvar/carregar obra, já que tudo passa por esta função.
        const x = obj._x !== undefined ? obj._x : (obj._raw && obj._raw._x);
        const y = obj._y !== undefined ? obj._y : (obj._raw && obj._raw._y);
        if (x !== undefined) c._x = x;
        if (y !== undefined) c._y = y;
        if (Array.isArray(obj.vinculoEstruturas)) c.vinculoEstruturas = obj.vinculoEstruturas.slice();
        return c;
    }

    /* ═══════════════════════════════════════
       CLASSIFICAÇÃO AUTOMÁTICA DE ENTIDADE
    ═══════════════════════════════════════ */
    // autoClassifyEntidade foi religada ao motor de regras do leitor (TASK-006), reusando a
    // mesma Tabela de Classificacao da aba Leitor. Operacao fixa em 'I' (neutro) porque esta
    // funcao so recebe o texto do ativo, sem operacao real conhecida -- nenhuma regra de
    // classificacao hoje exige um valor de operacao que 'I' nao satisfaca. O texto do ativo e
    // passado tambem como "texto bruto" porque e exatamente isso que o codigo original
    // verificava aqui (ver .ai/tasks/TASK-006-25-09-2026.md).
    function autoClassifyEntidade(ativoTexto) {
        if (!ativoTexto) return '0';
        const resultado = RegrasLeitorEngine.classificar(
            'I', ativoTexto, '', '', ativoTexto,
            window.__regrasLeitorClassificacao || []
        );
        return resultado.entidade;
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
    function recalcAllQtdAtivos() {
        const data = tableStates.cabos.data;
        for (let i = data.length - 1; i >= 0; i--) {
            const row = data[i];
            if (!row) continue;
            let faseFromNext = null;
            if (isStandaloneLine(row.ativo) && i + 1 < data.length) {
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
    function recalcRowAndAbove(index) {
        const data = tableStates.cabos.data;
        if (index < 0 || index >= data.length) return;

        function calc(i) {
            if (i < 0 || i >= data.length) return;
            const r = data[i];
            if (!r) return;
            let faseNext = null;
            if (isStandaloneLine(r.ativo) && i + 1 < data.length) {
                faseNext = extrairFase(data[i + 1].ativo);
            }
            const isNeg = r.qtdAtivos !== undefined && r.qtdAtivos !== null && r.qtdAtivos.toString().trim().startsWith('-');
            let nv = calcularQtdAtivos(r.ativo, faseNext);
            r.qtdAtivos = (isNeg && nv > 0) ? "-" + nv : nv;
        }

        calc(index);
        if (index - 1 >= 0 && isStandaloneLine(data[index - 1].ativo)) {
            calc(index - 1);
        }
    }

    // Ao desfocar um input de ativo, tenta auto-classificar se a Entidade for "0"
    document.addEventListener('blur', (e) => {
        if (e.target.classList.contains('inp-ativo')) {
            const tr = e.target.closest('tr');
            if (tr) {
                const selEnt = tr.querySelector('.sel-entidade');
                if (selEnt && selEnt.value === '0') {
                    const novaEntidade = autoClassifyEntidade(e.target.value);
                    if (novaEntidade !== '0') {
                        selEnt.value = novaEntidade;
                        // Aciona evento change para salvar no state e history
                        selEnt.dispatchEvent(new Event('change'));
                    }
                }
            }
        }
    }, true);


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

    /** Acrescenta as linhas da obra às tabelas atuais (um passo de histórico). */
    function adicionarObraAoProjeto(snap) {
        if (!snap) { showToast('Dados da obra inválidos.'); return false; }
        pushHistory();  // snapshot ANTES de adicionar
        ['cabos', 'outros'].forEach(type => {
            const linhas = linhasDoSnap(snap, type);
            if (!linhas.length) return;
            linhas.forEach(row => tableStates[type].data.push(JSON.parse(JSON.stringify(row))));
            if (type === 'cabos') recalcAllQtdAtivos();
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
        pushHistory();  // snapshot ANTES de subtrair
        ['cabos', 'outros'].forEach(type => {
            const linhas = linhasDoSnap(snap, type);
            if (!linhas.length) return;
            linhas.forEach(row => {
                const clonada = JSON.parse(JSON.stringify(row));
                if (type === 'cabos') clonada.qtdAtivos = "-" + Math.abs(clonada.qtdAtivos || 0);
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
            info.parametros.forEach(p => {
                const bloco = document.createElement('div');
                bloco.style.cssText = 'margin-bottom:10px;';
                const rot = document.createElement('label');
                rot.style.cssText = 'display:block; font-size:0.85rem; font-weight:600;';
                rot.textContent = p.rotulo;
                const onde = document.createElement('div');
                onde.style.cssText = 'font-size:0.7rem; color:#8b949e;';
                onde.textContent = (p.onde || []).join(' · ');
                const inp = document.createElement('input');
                inp.type = 'text'; inp.inputMode = 'decimal'; inp.className = 'modal-input';
                inp.placeholder = 'número maior que zero';
                if (p.padrao !== null && p.padrao !== undefined) inp.value = String(p.padrao).replace('.', ',');
                inp.dataset.chave = p.chave;
                bloco.append(rot, onde, inp);
                campos.appendChild(bloco);
                inputs.push(inp);
            });
            const fechar = v => { modal.classList.add('hidden'); resolve(v); };
            document.getElementById('btn-cancel-param-modelo').onclick = () => fechar(null);
            const btnGerar = document.getElementById('btn-gerar-param-modelo');
            const gerar = async () => {
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
                    btnGerar.disabled = false;
                }
            };
            btnGerar.onclick = gerar;
            inputs.forEach(i => { i.onkeydown = ev => { if (ev.key === 'Enter') gerar(); }; });
            modal.classList.remove('hidden');
            if (inputs[0]) inputs[0].focus();
        });
    }

    /** Há `V` de modelo sobrando nas tabelas? (espelha services/modelos_obra.py — Regra 4: mudou lá, mude aqui). */
    const RE_V_CABO = /^.*\S\s+V(\([^()\s]+\))?(\s+[mM])?\s*$/;
    const RE_V_OUTROS = /(^|\s)\*?V(\([^()\s]+\))?-\S/;
    function contarVariaveisSobrando() {
        let n = 0;
        (tableStates.cabos.data || []).forEach(r => { if (r && RE_V_CABO.test(String(r.ativo || '').trim())) n++; });
        (tableStates.outros.data || []).forEach(r => { if (r && RE_V_OUTROS.test(String(r.ativo || '').trim())) n++; });
        return n;
    }

    let execucaoOrigemAutonomo = null;

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

        pushHistory();  // snapshot ANTES de substituir
        // TASK-059: obra gerada pelo modo autônomo -> ao salvar, o servidor compara com o que o usuário deixou (aprendizado)
        execucaoOrigemAutonomo = (snap.autonomo && typeof snap.autonomo.execucao_id === 'string' && !snap.autonomo.revertida) ? snap.autonomo.execucao_id : null;
        tableStates.cabos.data  = (cabosData  || []).map(deepClone);
        tableStates.outros.data = (outrosData || []).map(deepClone);
        
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
    
    function addChatMessage(text, sender) {
        const div = document.createElement('div');
        div.className = `chat-message ${sender}`;
        div.innerHTML = text.replace(/\n/g, '<br>');
        aiChatMessages.appendChild(div);
        aiChatMessages.scrollTop = aiChatMessages.scrollHeight;
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
                pushHistory();
                tableStates[alvo].data.sort((a, b) => {
                    let valA = (a[action.coluna] || '').toString();
                    let valB = (b[action.coluna] || '').toString();
                    if (action.ordem === 'desc') return valB.localeCompare(valA);
                    return valA.localeCompare(valB);
                });
                renderTable(alvo);
                showToast(`Tabela ${alvo} ordenada!`);
            } 
            else if (action.acao_ui === 'filtrar') {
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
                FILTER_FIELDS.forEach(field => filterSelections[alvo][field].clear());
                applyFilters(alvo);
                refreshAllFilters(alvo);
                showToast(`Filtros limpos na tabela ${alvo}!`);
            }
        });
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
        btnSendChat.disabled = true;

        // Se for uma regra, salva no banco e informa a IA
        const isRule = text.toLowerCase().includes('lembre-se') || text.toLowerCase().includes('regra:');
        if (isRule) {
            try {
                await fetch('/api/regras', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ conteudo: text, projeto_codigo: localStorage.getItem('projeto_selecionado_codigo') || '229' })
                });
                showToast('Regra aprendida!');
            } catch (e) {
                console.error("Erro ao salvar regra", e);
            }
        }

        // Indicador carregando
        const loadingDiv = document.createElement('div');
        loadingDiv.className = 'chat-message ai';
        loadingDiv.innerHTML = '<span class="loading-dots">Gerando</span>';
        aiChatMessages.appendChild(loadingDiv);
        aiChatMessages.scrollTop = aiChatMessages.scrollHeight;

        // Montar contexto da tabela apenas quando o prompt pede análise/edição
        const palavrasContexto = ['analis', 'alterar', 'editar', 'excluir', 'remover', 'substituir',
            'tudo', 'todas', 'tabela', 'leia', 'leitura', 'completo', 'lista',
            'quantos', 'total', 'verifique', 'cheque', 'corrig'];
        const precisaContexto = palavrasContexto.some(kw => text.toLowerCase().includes(kw));
        let tableContext = "";
        if (precisaContexto) {
            tableContext = "\n\nESTADO ATUAL DA TABELA:\n";
            tableStates.cabos.data.forEach((r, i) => { if(r) tableContext += `[CABOS-${i}] | ${r.operacao} | ${r.ativo}\n`; });
            tableStates.outros.data.forEach((r, i) => { if(r) tableContext += `[OUTROS-${i}] | ${r.operacao} | ${r.ativo}\n`; });
        }

        try {
            const reqBody = {
                prompt: text,
                table_context: tableContext,
                history: chatHistory,
                provider: providerChat,
                openai_base_url: providerChat === 'openai' ? (localStorage.getItem('openai_base_url') || '') : '',
                projeto_codigo: localStorage.getItem('projeto_selecionado_codigo') || '229'
            };
            const response = await fetch('/api/gemini/chat', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    [camposChat.keyHeader]: apiKey,
                    [camposChat.modelHeader]: customModel
                },
                body: JSON.stringify(reqBody)
            });

            aiChatMessages.removeChild(loadingDiv);

            if (!response.ok) {
                if (response.status === 401) {
                    modalApikey.classList.remove('hidden');
                    addChatMessage("Chave da API não encontrada. Por favor, insira e tente novamente.", 'error');
                } else {
                    const err = await response.json();
                    addChatMessage(`Erro: ${err.detail || 'Falha na comunicação'}`, 'error');
                }
                aiChatInput.disabled = false;
                btnSendChat.disabled = false;
                return;
            }

            // Ler a resposta em stream
            const reader = response.body.getReader();
            const decoder = new TextDecoder();
            let replyText = "";
            
            // Cria a bolha da IA vazia
            const div = document.createElement('div');
            div.className = 'chat-message ai';
            aiChatMessages.appendChild(div);

            while (true) {
                const { done, value } = await reader.read();
                if (done) break;
                replyText += decoder.decode(value, { stream: true });
                div.innerHTML = replyText.replace(/\n/g, '<br>');
                aiChatMessages.scrollTop = aiChatMessages.scrollHeight;
            }
            
            // Finaliza decodificação
            replyText += decoder.decode();

            // Detectar se a resposta é vazia ou um erro do backend
            if (!replyText.trim()) {
                div.className = 'chat-message error';
                div.innerHTML = 'A IA não retornou resposta. Verifique a API Key e o modelo selecionado.';
                aiChatInput.disabled = false;
                btnSendChat.disabled = false;
                return;
            }
            if (replyText.trim().startsWith('[ERRO]')) {
                div.className = 'chat-message error';
                div.innerHTML = replyText.replace(/\n/g, '<br>');
                aiChatInput.disabled = false;
                btnSendChat.disabled = false;
                return;
            }

            // --- NOVO: Interceptar Ações de UI (JSON) ---
            const uiActionMatch = replyText.match(/\{[\s\S]*?"acao_ui"[\s\S]*?\}/);
            if (uiActionMatch) {
                try {
                    const action = JSON.parse(uiActionMatch[0]);
                    if (action.acao_ui === 'obra') {   // comando de obra (TASK-028): confirmação obrigatória, id validado
                        div.textContent = 'Comando de obra recebido. Confirme na janela para aplicar.';
                        chatHistory.push({ role: 'user', parts: [{ text: text }] });
                        chatHistory.push({ role: 'model', parts: [{ text: '*(Comando de obra)*' }] });
                        aiChatInput.disabled = false;
                        btnSendChat.disabled = false;
                        const resultado = await executarAcaoObra(action);
                        div.textContent = resultado.msg;
                        showToast(resultado.msg);
                        aiChatInput.focus();
                        return;
                    }
                    div.innerHTML = `<span style="color:#3fb950;">Ação de interface concluída: ${action.acao_ui}</span>`;
                    chatHistory.push({ role: 'user', parts: [{ text: text }] });
                    chatHistory.push({ role: 'model', parts: [{ text: "*(Ação de UI executada)*" }] });
                    executeUiAction(action);
                    aiChatInput.disabled = false;
                    btnSendChat.disabled = false;
                    aiChatInput.focus();
                    return;
                } catch(e) {
                    console.error("Erro ao parsear JSON de UI:", e);
                    div.innerHTML = `<span style="color:#ff0000;">Falha ao executar ação de interface. Verifique o console.</span><br>` + replyText.replace(/\n/g, '<br>');
                }
            } else {
                div.innerHTML = replyText.replace(/\n/g, '<br>');
            }
            // --------------------------------------------

            chatHistory.push({ role: 'user', parts: [{ text: text }] });
            chatHistory.push({ role: 'model', parts: [{ text: replyText }] });

            // Processar a tabela retornada, se houver
            const tableData = extractTableFromMarkdown(replyText);
            if (tableData.length > 0) {
                tableData.forEach(row => {
                    const cmd = row.comando.toUpperCase().replace(/\*/g, '').trim();
                    
                    if (cmd === 'EDITAR' || cmd === 'EXCLUIR' || cmd === 'REMOVER') {
                        const idParts = row.idStr.split('-');
                        if (idParts.length === 2) {
                            const type = idParts[0].toLowerCase();
                            const idx = parseInt(idParts[1], 10);
                            
                            if (tableStates[type] && tableStates[type].data[idx]) {
                                if (cmd === 'EXCLUIR' || cmd === 'REMOVER') {
                                    // Apenas anula (null) e depois limpa com filter, para não quebrar 
                                    // a ordem dos índices caso a IA mande excluir múltiplos itens
                                    tableStates[type].data[idx] = null;
                                } else {
                                    const isCabo = autoClassifyEntidade(row.ativosStr) === 'CABO';
                                    const novaEnt = autoClassifyEntidade(row.ativosStr);
                                    tableStates[type].data[idx].operacao = row.operacao.trim() || 'I';
                                    tableStates[type].data[idx].ativo = row.ativosStr;
                                    tableStates[type].data[idx].entidade = novaEnt !== '0' ? novaEnt : (isCabo ? 'CABO' : '0');
                                }
                            }
                        }
                        return; // Se for edição/exclusão, NUNCA adiciona uma nova linha (mesmo se o ID for inválido)
                    }
                    
                    // Roteamento baseado no formato estrito
                    let isCabo = false;
                    const ativoTest = row.ativosStr.trim().toUpperCase();
                    // Se termina em 'm' ou 'M' precedido por numero e espaço (ex: "35 m")
                    if (/\d+[\.,]?\d*\s*M$/.test(ativoTest)) {
                        isCabo = true;
                    } else if (/^\d+\s*-/.test(ativoTest)) { 
                        // Formato Quantidade-Ativo vai sempre para Outros
                        isCabo = false;
                    } else if (ativoTest.startsWith('DT') || ativoTest.startsWith('CV')) {
                        // Poste vai sempre para Outros
                        isCabo = false;
                    } else {
                        // Fallback original
                        isCabo = autoClassifyEntidade(row.ativosStr) === 'CABO';
                    }
                    
                    const type = isCabo ? 'cabos' : 'outros';
                    const novaEnt = autoClassifyEntidade(row.ativosStr);
                    tableStates[type].data.push({
                        entidade: novaEnt !== '0' ? novaEnt : (isCabo ? 'CABO' : '0'),
                        operacao: row.operacao.trim() || 'I', // Mantém asteriscos
                        ativo: row.ativosStr
                    });
                });
                
                // Limpar os nulos deixados por exclusões
                tableStates.cabos.data = tableStates.cabos.data.filter(r => r !== null);
                tableStates.outros.data = tableStates.outros.data.filter(r => r !== null);
                pushHistory();  // snapshot ANTES do render/atualização -- já inserimos, agora registramos a ação
                recalcAllQtdAtivos();
                renderTable('cabos');
                renderTable('outros');
                buildAtivoSets(); 
                buildDataLists();
                refreshAllFilters('cabos');
                refreshAllFilters('outros');
                showToast(`+${tableData.length} ativos inseridos!`);
            }
        } catch (error) {
            if (loadingDiv.parentNode) aiChatMessages.removeChild(loadingDiv);
            addChatMessage(`Erro de conexão: ${error.message}`, 'error');
        } finally {
            aiChatInput.disabled = false;
            btnSendChat.disabled = false;
            aiChatInput.focus();
        }
    }

    btnSendChat.addEventListener('click', sendToGemini);
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
        let params;
        try {
            const r = await fetch('/api/obras/modelo/detectar', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(dados) });
            if (!r.ok) throw new Error((await mensagensDeErro(r)).join(' '));
            params = (await r.json()).parametros;
        } catch (e) { showToast('Erro ao procurar variáveis: ' + e.message); return; }
        if (!params.length) {
            alert('Nenhuma variável V encontrada. Escreva V no lugar de uma quantidade — Cabos: "CAA 2 ABC V m"; Outros: "V-U4" ou "*V-CFU" — e tente de novo.');
            return;
        }
        document.getElementById('input-modelo-nome').value = modeloEditando ? modeloEditando.nome : '';
        document.getElementById('modelo-save-erro').textContent = '';
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
            const pad = document.createElement('input');
            pad.type = 'text'; pad.className = 'modal-input'; pad.placeholder = 'Valor padrão (opcional)'; pad.inputMode = 'decimal';
            const pa = antigos[p.chave] ? antigos[p.chave].padrao : p.padrao;
            pad.value = pa === null || pa === undefined ? '' : String(pa).replace('.', ',');
            bloco.append(onde, rot, pad);
            area.appendChild(bloco);
            return { chave: p.chave, rot, pad };
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
                parametros: linhas.map(l => ({ chave: l.chave, rotulo: l.rot.value.trim(), padrao: l.pad.value.trim().replace(',', '.') || null }))
            };
            try {
                const r = await fetch('/api/obras', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(corpo) });
                if (!r.ok) { erro.textContent = (await mensagensDeErro(r)).join(' '); return; }
                modalSaveModelo.classList.add('hidden');
                modeloEditando = null; atualizarAvisoModelo();
                showToast('Modelo salvo.');
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

    /**
     * Payload de cálculo do orçamento (o mesmo que "Montar Orçamento" grava em orcamentoPayload):
     * sincroniza a Totalizadora no modo padrão e a converte para os formatos que o backend lê.
     * Extraída do handler do botão (TASK-014) para a validação usar exatamente o mesmo payload.
     */
    async function obterPayloadCalculo() {
        let payloadCabos = [];
        let payloadOutros = [];

        // Se estivermos no modo padrão (Cabos e Postes), atualiza a totalizadora silenciosamente antes
        const radioPadrao = document.querySelector('input[name="view_mode"][value="padrao"]');
        if (radioPadrao && radioPadrao.checked) {
            await syncTotalizadora(false);
        }
        
        // Agora, INVARIAVELMENTE, constrói o payload a partir da Totalizadora
        tableStates.totalizadora.data.forEach(item => {
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
            if (r) linhas.push({ id: `${prefixo}-${i}`, operacao: r.operacao || '', ativo: r.ativo || '' });
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

    const btnMontarOrcamento = document.getElementById('btn-montar-orcamento');
    const modalOrcamento = document.getElementById('modal-orcamento');
    const tbodyOrcamento = document.getElementById('resultado-orcamento-tbody');

    if (btnMontarOrcamento) {
        btnMontarOrcamento.addEventListener('click', async () => {
            // Feedback visual: desabilita botão e exibe "Calculando..."
            const textoOriginal = btnMontarOrcamento.textContent.trim();
            btnMontarOrcamento.disabled = true;
            btnMontarOrcamento.textContent = 'Calculando...';

            try {
                // TASK-058: `V` de modelo que sobrou nas tabelas não vira orçamento
                const nV = contarVariaveisSobrando();
                if (nV > 0) {
                    alert(`Há ${nV} linha(s) com variável de modelo (V) sem valor. Gere a obra pelo modelo (Carregar/Adicionar/Subtrair) ou troque o V por um número antes de montar o orçamento.`);
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

                const payload = {
                    cabos: payloadCabos,
                    outros: payloadOutros,
                    projeto: projVal
                };
                
                if(selectProj) {
                    localStorage.setItem('projeto_selecionado', projVal);
                    localStorage.setItem('projeto_selecionado_codigo', projCode);
                }
                
                localStorage.setItem('orcamentoPayload', JSON.stringify(payload));
                window.open('/resultado_orcamento', '_blank');
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
    const btnLimparTabelas = document.getElementById('btn-limpar-tabelas');
    if (btnLimparTabelas) {
        btnLimparTabelas.addEventListener('click', () => {
            const totalLinhas = tableStates.cabos.data.length + tableStates.outros.data.length;
            if (totalLinhas === 0) {
                showToast('As tabelas já estão vazias.');
                return;
            }
            if (!confirm('Tem certeza que deseja limpar TODAS as tabelas (Cabos e Outros)?\nEsta ação pode ser desfeita com Ctrl+Z.')) return;

            pushHistory(); // snapshot ANTES de limpar

            tableStates.cabos.data  = [];
            tableStates.outros.data = [];
            localStorage.removeItem('processar_dados');

            renderTable('cabos');
            renderTable('outros');
            buildAtivoSets();
            buildDataLists();
            refreshAllFilters('cabos');
            refreshAllFilters('outros');

            showToast('Tabelas limpas! Use Ctrl+Z para desfazer.');
        });
    }

    /* ═══════════════════════════════════════
       GERADOR DE CÓDIGOS - CABOS
    ═══════════════════════════════════════ */
    let linhasModalCabos = [];
    let debounceTimeoutCabos = null;

    window.abrirModalCabos = function() {
        linhasModalCabos = [{ op: 'I', ativo: '', fase: 'ABC', comp: 80, qtd: 1, desc: '' }];
        renderizarTabelaModalCabos();
        document.getElementById('modal-cabos-gerador').classList.remove('hidden');
    };

    window.adicionarLinhaCabos = function() {
        linhasModalCabos.push({ op: 'I', ativo: '', fase: 'ABC', comp: 80, qtd: 1, desc: '' });
        renderizarTabelaModalCabos();
    };

    window.removerLinhaCabos = function(index) {
        linhasModalCabos.splice(index, 1);
        renderizarTabelaModalCabos();
    };

    async function buscarDescricaoAtivo(index) {
        const linha = linhasModalCabos[index];
        const ativoStr = (linha.ativo || '').trim();
        if (!ativoStr) {
            linha.desc = '';
            renderizarTabelaModalCabos();
            return;
        }

        const upperStr = ativoStr.toUpperCase();
        if (upperStr.includes('MULT') || upperStr.includes('MTX')) {
            linha.comp = 40;
        } else {
            linha.comp = 80;
        }

        const queryStr = ativoStr.replace(/\s+/g, '');

        try {
            const res = await fetch('/api/orcamento/search?q=' + encodeURIComponent(queryStr) + '&col=ativo');
            const data = await res.json();
            if (data.resultados && data.resultados.length > 0) {
                linha.desc = data.resultados[0].desc_ativo || 'Ativo encontrado';
            } else {
                linha.desc = 'Ativo não encontrado';
            }
        } catch (e) {
            linha.desc = 'Erro na busca';
        }
        
        // Atualiza o DOM diretamente para evitar recriação da tabela e perda de foco
        const tbody = document.getElementById('body-modal-cabos');
        if (tbody && tbody.children[index]) {
            const tr = tbody.children[index];
            const inComp = tr.children[3]?.querySelector('input');
            if (inComp) inComp.value = linha.comp;
            const tdDesc = tr.children[5];
            if (tdDesc) tdDesc.textContent = linha.desc;
        }
    }

    window.atualizarLinhaCabos = function(index, field, value) {
        linhasModalCabos[index][field] = value;
        
        if (field === 'ativo') {
            buscarDescricaoAtivo(index);
        }
    };

    function renderizarTabelaModalCabos() {
        const tbody = document.getElementById('body-modal-cabos');
        if(!tbody) return;
        tbody.innerHTML = '';

        linhasModalCabos.forEach((linha, index) => {
            const tr = document.createElement('tr');
            
            const tdOp = document.createElement('td');
            tdOp.innerHTML = `<select class="modal-input" style="width:100%; padding:4px;" onchange="atualizarLinhaCabos(${index}, 'op', this.value)">
                ${OPERACOES.map(op => `<option value="${op}" ${linha.op === op ? 'selected' : ''}>${op}</option>`).join('')}
            </select>`;
            
            const tdAtivo = document.createElement('td');
            tdAtivo.innerHTML = `<input type="text" class="modal-input" style="width:100%; padding:4px;" value="${linha.ativo}" onchange="atualizarLinhaCabos(${index}, 'ativo', this.value)" placeholder="Digite o ativo...">`;
            
            const tdFase = document.createElement('td');
            const fases = ['ABC', 'A', 'B', 'C', 'AC'];
            tdFase.innerHTML = `<select class="modal-input" style="width:100%; padding:4px;" onchange="atualizarLinhaCabos(${index}, 'fase', this.value)">
                ${fases.map(f => `<option value="${f}" ${linha.fase === f ? 'selected' : ''}>${f}</option>`).join('')}
            </select>`;
            
            const tdComp = document.createElement('td');
            tdComp.innerHTML = `<input type="number" class="modal-input" style="width:100%; padding:4px;" value="${linha.comp}" oninput="atualizarLinhaCabos(${index}, 'comp', parseFloat(this.value) || 0)">`;
            
            const tdQtd = document.createElement('td');
            tdQtd.innerHTML = `<input type="number" class="modal-input" style="width:100%; padding:4px;" value="${linha.qtd}" min="1" oninput="atualizarLinhaCabos(${index}, 'qtd', parseInt(this.value) || 1)">`;
            
            const tdDesc = document.createElement('td');
            tdDesc.style.fontSize = '0.75rem';
            tdDesc.style.color = '#8b949e';
            tdDesc.textContent = linha.desc;
            
            const tdDel = document.createElement('td');
            tdDel.innerHTML = `<button class="btn-danger-icon" onclick="removerLinhaCabos(${index})" title="Excluir" style="padding: 4px 8px; border: none; background: none; color: #f85149; cursor: pointer;">✖</button>`;
            
            tr.appendChild(tdOp);
            tr.appendChild(tdAtivo);
            tr.appendChild(tdFase);
            tr.appendChild(tdComp);
            tr.appendChild(tdQtd);
            tr.appendChild(tdDesc);
            tr.appendChild(tdDel);
            
            tbody.appendChild(tr);
        });
    }

    window.inserirGeradorNaTabelaCabos = function() {
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
       CABOS — Preset rápido de cabo
    ═══════════════════════════════════════ */
    window.adicionarCaboPreset = function(ativo, comp) {
        linhasModalCabos.push({ op: 'I', ativo: ativo, fase: 'ABC', comp: comp, qtd: 1, desc: '' });
        renderizarTabelaModalCabos();
        // Busca desc do ativo pra última linha adicionada
        const idx = linhasModalCabos.length - 1;
        buscarDescricaoAtivo(idx);
    };

    /* ═══════════════════════════════════════
       POSTES E ESTRUTURAS — Modal Gerador
    ═══════════════════════════════════════ */
    let linhasModalPostes = [];

    window.abrirModalPostes = function() {
        linhasModalPostes = [{ op: 'I', ativo: '', qtd: 1, desc: '' }];
        renderizarTabelaModalPostes();
        document.getElementById('modal-postes-gerador').classList.remove('hidden');
    };

    window.adicionarLinhaPostes = function() {
        linhasModalPostes.push({ op: 'I', ativo: '', qtd: 1, desc: '' });
        renderizarTabelaModalPostes();
    };

    window.removerLinhaPostes = function(index) {
        linhasModalPostes.splice(index, 1);
        renderizarTabelaModalPostes();
    };

    window.atualizarLinhaPostes = function(index, field, value) {
        linhasModalPostes[index][field] = value;
        if (field === 'ativo') {
            buscarDescricaoPoste(index);
        }
    };

    window.adicionarPostePreset = function(ativo) {
        linhasModalPostes.push({ op: 'I', ativo: ativo, qtd: 1, desc: '' });
        const idx = linhasModalPostes.length - 1;
        renderizarTabelaModalPostes();
        buscarDescricaoPoste(idx);
    };

    /* 
     * clicouPoste: postes sempre criam uma nova linha com o nome do poste.
     * Se a última linha estiver vazia, reutiliza ela. Se não, cria nova.
     */
    window.clicouPoste = function(nome) {
        const last = linhasModalPostes[linhasModalPostes.length - 1];
        const opAtual = last ? last.op : 'I';
        if (!last || last.ativo.trim() !== '') {
            // Última linha já tem conteúdo → cria nova linha
            linhasModalPostes.push({ op: opAtual, ativo: nome, qtd: 1, desc: '' });
        } else {
            // Linha em branco → preenche
            last.ativo = nome;
        }
        renderizarTabelaModalPostes();
        buscarDescricaoPoste(linhasModalPostes.length - 1);
    };

    /*
     * clicouEstrutura: SEMPRE acumula na linha atual, independente do conteúdo.
     * Ex: "DT11/300" + clique N3 → "DT11/300 1-N3"
     *     "DT11/300 1-N3" + clique N3 → "DT11/300 2-N3"
     *     "DT11/300 1-N3" + clique B2 → "DT11/300 1-N3 1-B2"
     * Nova linha apenas com clicouPoste() ou botão Adicionar Linha em Branco.
     */
    window.clicouEstrutura = function(nome) {
        if (linhasModalPostes.length === 0) {
            linhasModalPostes.push({ op: 'I', ativo: `1-${nome}`, qtd: 1, desc: '' });
            renderizarTabelaModalPostes();
            return;
        }
        
        const last = linhasModalPostes[linhasModalPostes.length - 1];
        const partes = last.ativo.trim() ? last.ativo.trim().split(/\s+/) : [];
        
        // Procura se esta estrutura já existe na linha no formato "X-nome"
        const escaped = nome.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(`^(\\d+)-${escaped}$`);
        
        let encontrado = false;
        for (let i = 0; i < partes.length; i++) {
            const m = partes[i].match(regex);
            if (m) {
                partes[i] = `${parseInt(m[1]) + 1}-${nome}`;
                encontrado = true;
                break;
            }
        }
        
        if (!encontrado) {
            partes.push(`1-${nome}`);
        }
        
        last.ativo = partes.join(' ');
        
        // Atualiza o campo no DOM diretamente (sem rerender para não perder foco)
        const tbody = document.getElementById('body-modal-postes');
        const lastIdx = linhasModalPostes.length - 1;
        if (tbody && tbody.children[lastIdx]) {
            const inputAtivo = tbody.children[lastIdx].children[1]?.querySelector('input');
            if (inputAtivo) inputAtivo.value = last.ativo;
        }
    };

    async function buscarDescricaoPoste(index) {
        const linha = linhasModalPostes[index];
        const ativoStr = (linha.ativo || '').trim();
        if (!ativoStr) {
            linha.desc = '';
            renderizarTabelaModalPostes();
            return;
        }

        const queryStr = ativoStr.replace(/\s+/g, '');
        try {
            const res = await fetch('/api/orcamento/search?q=' + encodeURIComponent(queryStr) + '&col=ativo');
            const data = await res.json();
            if (data.resultados && data.resultados.length > 0) {
                linha.desc = data.resultados[0].desc_ativo || 'Ativo encontrado';
            } else {
                linha.desc = 'Ativo não encontrado';
            }
        } catch (e) {
            linha.desc = 'Erro na busca';
        }

        const tbody = document.getElementById('body-modal-postes');
        if (tbody && tbody.children[index]) {
            const tr = tbody.children[index];
            const tdDesc = tr.children[3];
            if (tdDesc) tdDesc.textContent = linha.desc;
        }
    }

    function renderizarTabelaModalPostes() {
        const tbody = document.getElementById('body-modal-postes');
        if (!tbody) return;
        tbody.innerHTML = '';

        linhasModalPostes.forEach((linha, index) => {
            const tr = document.createElement('tr');

            const tdOp = document.createElement('td');
            tdOp.innerHTML = `<select class="modal-input" style="width:100%; padding:4px;" onchange="atualizarLinhaPostes(${index}, 'op', this.value)">
                ${OPERACOES.map(op => `<option value="${op}" ${linha.op === op ? 'selected' : ''}>${op}</option>`).join('')}
            </select>`;

            const tdAtivo = document.createElement('td');
            tdAtivo.innerHTML = `<input type="text" class="modal-input" style="width:100%; padding:4px;" value="${linha.ativo}" onchange="atualizarLinhaPostes(${index}, 'ativo', this.value)" placeholder="Digite o ativo (ex: PC 9 200)...">`;

            const tdQtd = document.createElement('td');
            tdQtd.innerHTML = `<input type="number" class="modal-input" style="width:100%; padding:4px;" value="${linha.qtd}" min="1" oninput="atualizarLinhaPostes(${index}, 'qtd', parseInt(this.value) || 1)">`;

            const tdDesc = document.createElement('td');
            tdDesc.style.fontSize = '0.75rem';
            tdDesc.style.color = '#8b949e';
            tdDesc.textContent = linha.desc;

            const tdDel = document.createElement('td');
            tdDel.innerHTML = `<button class="btn-danger-icon" onclick="removerLinhaPostes(${index})" title="Excluir" style="padding: 4px 8px; border: none; background: none; color: #f85149; cursor: pointer;">✖</button>`;

            tr.appendChild(tdOp);
            tr.appendChild(tdAtivo);
            tr.appendChild(tdQtd);
            tr.appendChild(tdDesc);
            tr.appendChild(tdDel);
            tbody.appendChild(tr);
        });
    }

    window.inserirGeradorNaTabelaOutros = function() {
        if (!linhasModalPostes.length) return;

        pushHistory();

        linhasModalPostes.forEach(linha => {
            const ativoBase = (linha.ativo || '').trim();
            if (!ativoBase) return;

            for (let i = 0; i < linha.qtd; i++) {
                tableStates.outros.data.push({
                    entidade: 'OUTRO',
                    operacao: linha.op,
                    ativo: ativoBase
                });
            }
        });

        renderTable('outros');
        document.getElementById('modal-postes-gerador').classList.add('hidden');
        buildAtivoSets();
        buildDataLists();
    };


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
        tbody.innerHTML = '<tr><td colspan="3" style="text-align:center; color:#8b949e; padding:20px;">Nenhum item de RAMAIS, IP ou Apoio relevante encontrado no processamento.</td></tr>';
        return;
    }

    dados.forEach((r) => {
        const tr = document.createElement('tr');

        const tdTipo = document.createElement('td');
        const tipoColor = r.entidade === 'RAMAIS' ? '#58a6ff' : r.entidade === 'IP' ? '#d2991e' : '#3fb950';
        tdTipo.innerHTML = `<span style="background:${tipoColor}22; color:${tipoColor}; border:1px solid ${tipoColor}44; padding:2px 7px; border-radius:4px; font-size:0.72rem; font-weight:600;">${r.entidade}</span>`;
        tdTipo.style.cssText = 'padding:6px 8px; white-space:nowrap;';

        const tdTexto = document.createElement('td');
        tdTexto.style.cssText = 'padding:4px 8px;';
        tdTexto.contentEditable = 'true';
        tdTexto.style.outline = 'none';
        tdTexto.textContent = r._textoOriginal || (r._raw && r._raw.texto) || r.texto || '';
        tdTexto.addEventListener('focus', () => tdTexto.style.background = 'rgba(88,166,255,0.06)');
        tdTexto.addEventListener('blur',  () => {
            tdTexto.style.background = '';
            window.recalcularAtivosRamais();
        });

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
}

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

window.syncTotalizadora = async function(forceUpdate = true) {
    if (tableStates.regras.data.length === 0) {
        carregarRegras();
    }

    if (forceUpdate && tableStates.totalizadora.data.length > 0) {
        if (!confirm("Isso irá apagar todas as edições manuais feitas na Tabela Totalizadora e puxar os dados novamente das tabelas Cabos e Outros. Deseja continuar?")) {
            return;
        }
    }

    const btnGerar = document.getElementById('btn-gerar-totalizadora');
    const originalText = btnGerar ? btnGerar.textContent : '';
    if (btnGerar) btnGerar.textContent = 'Carregando Base...';

    // Carregar a base técnica caso não tenha carregado
    if (orcamentoBaseData.length === 0) {
        try {
            const res = await fetch('/api/orcamento/dados');
            const data = await res.json();
            orcamentoBaseData = data.dados || [];
        } catch(e) {
            console.error("Erro ao carregar base de orçamento", e);
        }
    }

    if (btnGerar) btnGerar.textContent = originalText;

    function buscarDescricao(ativoNome) {
        const ativoUpper = (ativoNome || '').trim().toUpperCase();
        const projName = (localStorage.getItem('projeto_selecionado') || '').toUpperCase();
        
        const row = orcamentoBaseData.find(r => 
            ((r.ativo || '').toUpperCase() === ativoUpper || (r.codigo || '').toUpperCase() === ativoUpper) &&
            ((r.projeto || '').trim().toUpperCase() === projName || (r.projeto || '').trim() === '')
        );
        if (row) {
            return { desc: row.desc_ativo || row.componente || row.desc_codigo || '', naoEncontrado: false };
        }
        return { desc: '', naoEncontrado: true };
    }

    const rawItems = [];
    let baseIdCounter = 1;

    // Processar Cabos
    const cabosData = tableStates.cabos.data;
    for (let idx = 0; idx < cabosData.length; idx++) {
        const item = cabosData[idx];
        if (!item) continue;
        const currentBaseId = baseIdCounter++;
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
    tableStates.outros.data.forEach(item => {
        if (!item) return;
        const currentBaseId = baseIdCounter++;
        
        const parts = item.ativo.split(/\s+/);

        parts.forEach(p => {
            if (!p.trim()) return;
            let q = 1;
            let aName = p.trim();
            const m = p.match(/^([\*\-]?\d+(?:[.,]\d+)?)[Xx\-](.+)$/i);
            if (m) {
                let qStr = m[1];
                let isNegative = false;
                if (qStr.startsWith('*') || qStr.startsWith('-')) {
                    isNegative = true;
                    qStr = qStr.substring(1);
                }
                q = parseFloat(qStr.replace(',', '.'));
                if (isNegative) q = -q;
                aName = m[2].trim();
            }
            
            let resDesc = buscarDescricao(aName);

            rawItems.push({
                baseId: `TOT-${currentBaseId}`,
                obs: item.entidade,
                operacao: item.operacao || 'I',
                ativo: aName,
                qtd: q,
                desc: resDesc.desc,
                naoEncontrado: resDesc.naoEncontrado,
                origem: 'OUTROS'
            });
        });
    });

    // Ativos compostos do vínculo cabo<->estrutura (TASK-034): um pseudo-item por aresta do
    // vínculo (TASK-032), origem "VINCULO", para as Regras de Conversão poderem casar contra ele
    // (ex.: ativo_de=CAA2_N4). Só existe aqui dentro — nunca aparece na Totalizadora sem match.
    tableStates.cabos.data.forEach((cabo, idxCabo) => {
        if (!Array.isArray(cabo.vinculoEstruturas) || !cabo.vinculoEstruturas.length) return;
        const info = extrairPrefixoEComprimentoCabo(cabo.ativo);
        const prefixoCabo = info ? info.prefixo : null;
        if (!prefixoCabo) return;
        cabo.vinculoEstruturas.forEach(idxEstrutura => {
            const estrutura = tableStates.outros.data[idxEstrutura];
            if (!estrutura || !estrutura.ativo) return;
            rawItems.push({
                baseId: `TOT-VINC-${idxCabo}-${idxEstrutura}`,
                obs: 'VINCULO',
                operacao: cabo.operacao || 'I',
                ativo: `${prefixoCabo}_${_codigoEstruturaVinculo(estrutura.ativo).toUpperCase()}`,
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
            if (!checkMatch(ruleAtivo, item.ativo)) applies = false;

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

    tableStates.totalizadora.data = newData;
    renderTotalizadora();
    showToast("Tabela Totalizadora atualizada!");
};

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
            tr.style.backgroundColor = 'rgba(248, 81, 73, 0.15)';
        }
        
        tr.innerHTML = `
            <td style="font-weight: bold; color: var(--text-secondary); text-align: center;">${row.id}</td>
            <td><input type="text" class="tot-input" value="${row.obs || ''}" onchange="updateTotRow(${index}, 'obs', this.value)"></td>
            <td>
                <select class="tot-input" style="text-align: center;" onchange="updateTotRow(${index}, 'operacao', this.value)">
                    ${(typeof OPERACOES !== 'undefined' ? OPERACOES : ['I', 'R', 'M']).map(op => `<option value="${op}" ${row.operacao === op ? 'selected' : ''}>${op}</option>`).join('')}
                </select>
            </td>
            <td><input type="text" class="tot-input tot-ativo-input" value="${row.ativo || ''}" onchange="updateTotRow(${index}, 'ativo', this.value)" onkeydown="handleTotKeydown(event, ${index}, 'ativo')" onpaste="handleTotPaste(event, ${index})"></td>
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

        if (field === 'ativo') {
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
            
            // Tenta buscar a descrição baseada no ativo
            const ativoUpper = (value || '').trim().toUpperCase();
            const projName = (localStorage.getItem('projeto_selecionado') || '').toUpperCase();
            
            const row = orcamentoBaseData.find(r => 
                ((r.ativo || '').toUpperCase() === ativoUpper || 
                 (r.componente || '').toUpperCase() === ativoUpper ||
                 (r.codigo || '').toUpperCase() === ativoUpper) &&
                ((r.projeto || '').trim().toUpperCase() === projName || (r.projeto || '').trim() === '')
            );
            
            if (row) {
                let desc = '';
                if ((row.codigo || '').toUpperCase() === ativoUpper) {
                    desc = row.desc_codigo || '';
                } else {
                    desc = row.desc_ativo || '';
                }
                
                tableStates.totalizadora.data[index].desc = desc;
                tableStates.totalizadora.data[index].naoEncontrado = false;
                
                const descInput = document.getElementById(`tot-desc-${index}`);
                if (descInput) descInput.value = desc;
                
                const tr = document.getElementById(`tot-row-${index}`);
                if (tr) tr.style.backgroundColor = '';
            } else {
                tableStates.totalizadora.data[index].desc = '';
                tableStates.totalizadora.data[index].naoEncontrado = true;
                
                const descInput = document.getElementById(`tot-desc-${index}`);
                if (descInput) descInput.value = '';
                
                const tr = document.getElementById(`tot-row-${index}`);
                if (tr) tr.style.backgroundColor = 'rgba(248, 81, 73, 0.15)';
            }
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
