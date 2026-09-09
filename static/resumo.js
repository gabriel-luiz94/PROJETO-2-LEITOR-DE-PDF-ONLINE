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
       DADOS GLOBAIS
    ═══════════════════════════════════════ */
    const ENTIDADES = ['0', 'CABO', 'CHAVE', 'TRAFO', 'ESTRUTURA', 'APOIO', 'IP', 'POSTE', 'RAMAIS', 'CERCA'];
    const OPERACOES = ['I', '*I', 'R', '*R', 'M', '*M'];

    let extractedDataCache = [];
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

    /* ═══════════════════════════════════════
       LÓGICA DE NEGÓCIO (portada do resumo.js antigo)
    ═══════════════════════════════════════ */
    function isGray(hex) {
        if (!hex || hex.length < 7) return false;
        const r = parseInt(hex.substring(1, 3), 16);
        const g = parseInt(hex.substring(3, 5), 16);
        const b = parseInt(hex.substring(5, 7), 16);
        return Math.abs(r-g) < 5 && Math.abs(g-b) < 5 && r > 20 && r < 230;
    }

    function processAtivoFormula(col2) {
        if (!col2) return '';
        let normalized = col2.replace(/[\u00AD\u2010-\u2015\u2212]/g, '-').replace(/\u00A0/g, ' ');
        const cleanArrume = t => t.replace(/\s*-\s*/g, '-').trim().replace(/\s+/g, ' ');
        const tArrume = cleanArrume(normalized);
        const tUpperMatched = tArrume.toUpperCase();
        if (tUpperMatched.includes('AFASTADOR')) return '1-AF';
        let text1 = '';
        const instMatch = tUpperMatched.match(/INST\.?(?:AL(?:AR|A)?)?\s+0*(\d+)\s*(?:-| )?\s*(\d*SI\d*|\d*RA\d*|\d*BI\d*|\d*R\d*|\d*B\d*|\d*S\d*|\d*CE\d*|\d*N\d*|\d*U\d*|\d*T\d*|ISOL)/);
        if (instMatch) {
            text1 = instMatch[1] + '-' + instMatch[2].replace(/\s/g, '');
        } else if (tUpperMatched.includes(' METROS')) {
            const match = tUpperMatched.match(/(\d+(?:[.,]\d+)?)\s*METROS/);
            if (match) text1 = match[1] + '-ROCO';
            else text1 = tArrume.replace(/ METROS/gi, '-ROCO');
        } else if (/CAL[CÇ]ADA|RECAL|REC\.\s*CAL/i.test(tUpperMatched)) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            text1 = (matchQty ? matchQty[1] : '1') + '-RECAL';
        } else if (/\bIP\b/i.test(tUpperMatched)) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            text1 = (matchQty ? matchQty[1] : '1') + '-IP';
        } else if (tUpperMatched.includes('CONC') && tUpperMatched.includes('BASE')) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            text1 = (matchQty ? matchQty[1] : '1') + '-BASE';
        } else if (/\bCOMPRESSOR\b/i.test(tUpperMatched) || /\bCAVA\b/i.test(tUpperMatched)) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            text1 = (matchQty ? matchQty[1] : '1') + '-CAVA';
        } else {
            if (!tUpperMatched.includes('FIOS')) text1 = tArrume;
            else {
                const arrumeRaw = col2.trim().replace(/\s+/g, ' ');
                text1 = '1-' + arrumeRaw.substring(0, Math.max(0, arrumeRaw.length - 5)) + 'F ';
            }
        }
        let text2 = text1.replace(/TR\s*-\s*3\s*-\s*/g, '1-TR3').replace(/kVA/gi, '').replace(/TR\s*-\s*1\s*-\s*/g, '1-TR1').replace(/TR\s*-\s*2\s*-\s*/g, '1-TR2');
        const comboMatch = text2.match(/([13])\s*-\s*100A\s*-\s*(\d+(?:,\d+)?(?:H|K))/i);
        if (comboMatch) {
            const q = comboMatch[1], e = comboMatch[2].replace(',', '').replace(' ', '');
            return `${q}-CFU ${q}-EF${e}`;
        }
        if (text2.includes('-100A-')) text2 = text2.replace('-100A-', `-CFU ${text2.charAt(0)}-EF`);
        return text2.replace(/3CF-400A/g, '3-CFA').replace(/112,5/g, '112').replace(/0,5H/g, '05H').replace(/PODAS/g, 'PODA').replace(/PODA M/g, 'PODA').replace(/PODA G/g, 'PODA').replace(/PODA P/g, 'PODA').replace(/ PODA/g, '-PODA').replace(/3CL-300A/g, '3-CFU 3-CL').replace(/TR15/g, 'TR105').trim();
    }

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

        const displayColor = (item.cor || '#000000').toUpperCase();
        let textoAtivo = processAtivoFormula(item.texto || '');
        let opAuto = 'M';
        if (displayColor === '#FF0000') opAuto = 'I';
        else if (isGray(displayColor)) opAuto = 'R';

        const uAtivo = textoAtivo.toUpperCase();
        const tUpper = (item.texto || '').replace(/[\u00AD\u2010-\u2015\u2212]/g, '-').replace(/\u00A0/g, ' ').toUpperCase().trim();
        const itemLayer = (item.layer || '').trim().toUpperCase();
        let entAuto = '0';

        const isRedOrGray = (displayColor === '#FF0000') || isGray(displayColor);
        const isRed = displayColor === '#FF0000';

        const elosFusivel = ['0,5H','1H','2H','3H','5H','6K','8K','10K','12K','15K','25K','30K','40K'];
        if (elosFusivel.includes(tUpper) && isRedOrGray) {
            entAuto = 'CHAVE';
            const globalIdx = extractedDataCache.indexOf(item);
            if (globalIdx !== -1) {
                let qtyFound = '';
                const start = Math.max(0, globalIdx - 10), end = Math.min(extractedDataCache.length - 1, globalIdx + 10);
                for (let i = start; i <= end; i++) {
                    if (i === globalIdx) continue;
                    const neighbor = extractedDataCache[i];
                    if (neighbor.pagina !== item.pagina) continue;
                    const m = neighbor.texto.replace(/\u00A0/g, ' ').toUpperCase().match(/([13])\s*-\s*100A/);
                    if (m) { qtyFound = m[1]; break; }
                }
                if (qtyFound) textoAtivo = `${qtyFound}-EF${tUpper.replace(',','').replace(' ','')}`;
            }
        }

        entAuto = '0';

        if (!tUpper.includes('BLOCO')) {
            if (isRed && tUpper.includes('FLY')) {
                if (tUpper.includes('REF'))  { textoAtivo = '1-RFLY'; opAuto = 'I'; entAuto = 'ESTRUTURA'; }
                else if (tUpper.includes('DESF')) { textoAtivo = '1-FLY'; opAuto = 'R'; entAuto = 'ESTRUTURA'; }
                else if (tUpper.includes('INST')) { textoAtivo = '1-FLY'; opAuto = 'I'; entAuto = 'ESTRUTURA'; }
            }
        } else { textoAtivo = ''; opAuto = 'M'; entAuto = '0'; }

        const apoioMarkers = ['-ROCO','-RECAL','-BASE','-CAVA','PODA'];
        const foundApoioCount = apoioMarkers.filter(m => uAtivo.includes(m)).length;
        const hasApoioConflict = tUpper.includes('APOIOS') || tUpper.includes('LARGURA') || (tUpper.includes('BASE') && tUpper.includes('CALÇADA')) || foundApoioCount > 1;

        if (entAuto === '0') {
            const isBlack = !isRedOrGray;
            const isRetens = itemLayer === '01_RETENS' || itemLayer === '01_RETENS_LV' || itemLayer === '01_LV';

            if (isBlack && isRetens) {
                const trafoMatch = tUpper.match(/TR\s*-\s*([123])/);
                if (trafoMatch) {
                    const qty = trafoMatch[1];
                    entAuto = 'TRAFO';
                    textoAtivo = `1-RTR${qty}`;
                    opAuto = (itemLayer === '01_RETENS_LV' || itemLayer === '01_LV') ? '*I' : 'I';
                }
            }

            if (isBlack && isRetens && entAuto === '0') {
                const chaveMatch = tUpper.match(/^([123])\s*-\s*100\s*A/);
                if (chaveMatch) {
                    const qty = chaveMatch[1];
                    entAuto = 'CHAVE';
                    textoAtivo = `${qty}-RCFU`;
                    opAuto = (itemLayer === '01_RETENS_LV' || itemLayer === '01_LV') ? '*I' : 'I';
                }
            }

            if (entAuto === '0') {
                if (/\sRS\s+[MT]\s/i.test(item.texto || '')) entAuto = 'RAMAIS';
                else if (uAtivo.includes('-ROCO') && !hasApoioConflict) entAuto = 'APOIO';
                else if (uAtivo.includes('-IP') || /\bIP\b/i.test(item.texto || '')) entAuto = 'IP';
                else if (uAtivo === '1-AF') entAuto = 'ESTRUTURA';
                else if ((uAtivo.includes('-RECAL') || uAtivo.includes('-BASE') || uAtivo.includes('-CAVA')) && !hasApoioConflict) entAuto = 'APOIO';
                else if ((tUpper.includes('DT') || tUpper.includes('CV')) && tUpper.includes('/') && isRedOrGray) entAuto = 'POSTE';
                else if (!tUpper.includes('DT') && !tUpper.includes('CV') && !tUpper.includes('AWG') && !tUpper.includes('#') && (() => {
                    if (/\bM?\d+x\d+/.test(tUpper)) return true;
                    if (tUpper.includes('ABC') && /\d+\s*M$/.test(tUpper)) return true;
                    if (/^CU\s*\d/.test(tUpper) || /\bCU\s*\d/.test(tUpper.substring(0, 5))) return true;
                    if (/^CA\s+\d/.test(tUpper) || /^CA\d/.test(tUpper)) return true;
                    if (/\b(?:CAL|CAA|CAZ)(?:\s|\d)/.test(tUpper)) return true;
                    if (/\bP\s*(16|25|35|50|70|95|120|150|185|240)\b/.test(tUpper)) return true;
                    if (tUpper.includes('X1X') && /\d+\s*M$/.test(tUpper)) return true;
                    return false;
                })()) {
                    if (isRedOrGray) entAuto = 'CABO';
                    else if (!isRedOrGray && (itemLayer === '01_RETENS' || itemLayer === '01_RETENS_LV')) {
                        entAuto = 'CABO';
                        opAuto = itemLayer === '01_RETENS_LV' ? '*M' : 'M';
                    }
                }
                else if (tUpper.includes('FIOS')) entAuto = 'CERCA';
                else if ((uAtivo.includes('-CF') || uAtivo.includes('-EF')) && opAuto !== 'M') entAuto = 'CHAVE';
                else if (uAtivo.includes('-TR') && opAuto !== 'M') {
                    entAuto = 'TRAFO';
                    const match = textoAtivo.match(/(1-TR\d+)/i);
                    if (match) textoAtivo = match[1].toUpperCase();
                }
                else if (uAtivo.includes('PODA') && opAuto !== 'M' && !hasApoioConflict) entAuto = 'APOIO';
            }
        }

        if (['IP','APOIO','CERCA','RAMAIS'].includes(entAuto)) opAuto = 'I';
        if (entAuto === '0' && !tUpper.includes('APOIOS')) {
            const hasEstruturaPattern = /\b\d+\s*-\s*(\d+[A-Z]{1,2}\d*|\d*[A-Z]{1,2}\d+|ISOL)\b/i.test(uAtivo) || /\b\d+\s*-\s*(SI|RA|BI|CE|N|U|T|R|B|S)\d+/i.test(uAtivo);
            const words = textoAtivo.trim().split(/\s+/);
            const allWordsValid = words.length > 0 && words.every(w => /^\d+-\S+$/i.test(w));
            if (hasEstruturaPattern && allWordsValid && (opAuto === 'I' || opAuto === 'R')) entAuto = 'ESTRUTURA';
        }
        if (itemLayer === '01_LV' && opAuto && !opAuto.startsWith('*')) opAuto = '*' + opAuto;

        return { entidade: entAuto, operacao: opAuto, ativo: textoAtivo, _raw: item };
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
                    inpAt.setSelectionRange(cursor, cursor);
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
                            recalcAllQtdAtivos();
                            // Atualiza visualmente os inputs de qtdAtivos
                            const body = document.getElementById(state.bodyId);
                            if (body) {
                                Array.from(body.children).forEach(tr => {
                                    const trIdx = parseInt(tr.dataset.index);
                                    const qtdInp = tr.querySelector('.inp-qtd-ativos');
                                    if (qtdInp && state.data[trIdx]) {
                                        qtdInp.value = state.data[trIdx].qtdAtivos !== undefined ? state.data[trIdx].qtdAtivos : '';
                                    }
                                });
                            }
                        }
                        refreshAllFilters(type);
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

        // Após render, aplica filtros ativos sem rebuildar dropdowns
        applyFilters(type);
    }

    /* ═══════════════════════════════════════
       CRUD
    ═══════════════════════════════════════ */
    function deleteRow(type, idx) {
        pushHistory();  // snapshot ANTES da exclusão
        tableStates[type].data.splice(idx, 1);
        renderTable(type);
        buildAtivoSets(); buildDataLists();
        refreshAllFilters(type);
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

    function pushHistory() {
        // Descarta redo futuro
        if (historyIdx < history.length - 1) history = history.slice(0, historyIdx + 1);
        history.push(snapshotState());
        if (history.length > MAX_HISTORY) history.shift();
        historyIdx = history.length - 1;
        updateHistoryUI();
    }

    function undo() {
        if (historyIdx <= 0) return;
        historyIdx--;
        applyUndoRedoSnapshot(history[historyIdx]);
        showToast('Desfeito');
    }

    function redo() {
        if (historyIdx >= history.length - 1) return;
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
        btnU.disabled = historyIdx <= 0;
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

    /* ═══════════════════════════════════════
       UTILS
    ═══════════════════════════════════════ */
    function deepClone(obj) {
        const c = { entidade: obj.entidade || '0', operacao: obj.operacao || 'M', ativo: obj.ativo || '' };
        if (obj.qtdAtivos !== undefined) c.qtdAtivos = obj.qtdAtivos;
        return c;
    }

    /* ═══════════════════════════════════════
       CLASSIFICAÇÃO AUTOMÁTICA DE ENTIDADE
    ═══════════════════════════════════════ */
    function autoClassifyEntidade(ativoTexto) {
        if (!ativoTexto) return '0';
        const tUpper = ativoTexto.toUpperCase().trim();
        
        // Verifica primeiro se não tem conflitos de APOIO
        const apoioMarkers = ["-ROCO", "-RECAL", "-BASE", "-CAVA", "PODA"];
        const foundApoioCount = apoioMarkers.filter(m => tUpper.includes(m)).length;
        const hasApoioConflict = tUpper.includes("APOIOS") || tUpper.includes("LARGURA") || (tUpper.includes("BASE") && tUpper.includes("CALÇADA")) || foundApoioCount > 1;

        if (tUpper.includes("-RTR")) return "TRAFO";
        if (tUpper.includes("-RCFU")) return "CHAVE";
        if (/\sRS\s+[MT]\s/i.test(tUpper)) return "RAMAIS";
        if (tUpper.includes("-ROCO") && !hasApoioConflict) return "APOIO";
        if (tUpper.includes("-IP") || /\bIP\b/i.test(tUpper)) return "IP";
        if (tUpper === "1-AF") return "ESTRUTURA";
        if ((tUpper.includes("-RECAL")||tUpper.includes("-BASE")||tUpper.includes("-CAVA")) && !hasApoioConflict) return "APOIO";
        if ((tUpper.includes("DT") || tUpper.includes("CV")) && tUpper.includes("/")) return "POSTE";
        
        // Cabos
        if (!tUpper.includes("DT") && !tUpper.includes("CV") && !tUpper.includes("AWG") && !tUpper.includes("#")) {
            if (/\bM?\d+x\d+/.test(tUpper)) return "CABO";
            if (tUpper.includes("ABC") && /\d+\s*M$/.test(tUpper)) return "CABO";
            if (/^CU\s*\d/.test(tUpper) || /\bCU\s*\d/.test(tUpper.substring(0, 5))) return "CABO";
            if (/^CA\s+\d/.test(tUpper) || /^CA\d/.test(tUpper) || /\b(?:CAL|CAA|CAZ)\s*\d/.test(tUpper)) return "CABO";
            if (/\bP\s*(16|25|35|50|70|95|120|150|185|240)\b/.test(tUpper)) return "CABO";
            if (tUpper.includes("X1X") && /\d+\s*M$/.test(tUpper)) return "CABO";
        }
        
        if (tUpper.includes("FIOS")) return "CERCA";
        if (tUpper.includes("-CF") || tUpper.includes("-EF")) return "CHAVE";
        if (tUpper.includes("-TR")) return "TRAFO";
        if (tUpper.includes("PODA") && !hasApoioConflict) return "APOIO";
        
        // Estrutura
        const hasEstruturaPattern = /\b\d+\s*-\s*(\d+[A-Z]{1,2}\d*|\d*[A-Z]{1,2}\d+|ISOL)\b/i.test(tUpper) || 
                                   /\b\d+\s*-\s*(SI|RA|BI|CE|N|U|T|R|B|S)\d+/i.test(tUpper);
        if (hasEstruturaPattern) return "ESTRUTURA";

        return "0"; // Default
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
    
    // API Key config
    const modalApikey = document.getElementById('modal-apikey');
    const inputApikey = document.getElementById('input-apikey');
    const inputModel = document.getElementById('input-model');
    
    async function fetchModels() {
        const keyToUse = inputApikey.value.trim() || "SAVED_IN_BACKEND";
        try {
            const resp = await fetch('/api/gemini/models', {
                headers: { 'X-Gemini-Key': keyToUse }
            });
            if (resp.ok) {
                const data = await resp.json();
                inputModel.innerHTML = '<option value="">Automático (gemini-3.1-flash-lite)</option>';
                (data.models || []).forEach(m => {
                    const opt = document.createElement('option');
                    opt.value = m.id || m;
                    opt.textContent = m.label || m.id || m;
                    inputModel.appendChild(opt);
                });
                inputModel.value = localStorage.getItem('gemini_model') || '';
            }
        } catch (e) {
            console.error("Erro ao carregar modelos", e);
        }
    }

    inputApikey.addEventListener('blur', fetchModels);

    const btnConfigApi = document.getElementById('btn-config-api');
    if (btnConfigApi) {
        btnConfigApi.addEventListener('click', () => {
            inputApikey.value = localStorage.getItem('gemini_api_key') || '';
            inputModel.value = localStorage.getItem('gemini_model') || '';
            fetchModels();
            modalApikey.classList.remove('hidden');
        });
    }

    const btnSaveApiKey = document.getElementById('btn-save-apikey');
    if (btnSaveApiKey) {
        btnSaveApiKey.addEventListener('click', () => {
            localStorage.setItem('ai_provider', 'gemini');
            localStorage.setItem('gemini_api_key', inputApikey.value.trim());
            localStorage.setItem('gemini_model', inputModel.value.trim());
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

        const apiKey = localStorage.getItem('gemini_api_key') || 'SAVED_IN_BACKEND';
        const customModel = localStorage.getItem('gemini_model') || 'SAVED_IN_BACKEND';

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
                    body: JSON.stringify({ conteudo: text })
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
                provider: "gemini",
                openai_base_url: ""
            };
            const response = await fetch('/api/gemini/chat', {
                method: 'POST',
                headers: { 
                    'Content-Type': 'application/json',
                    'X-Gemini-Key': apiKey,
                    'X-Gemini-Model': customModel
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
        modalSaveObra.classList.remove('hidden');
    });

    document.getElementById('btn-confirm-save-obra').addEventListener('click', async () => {
        const nome = document.getElementById('input-obra-nome').value.trim();
        if (!nome) return alert('Digite um nome');
        
        const obra = {
            id: 'obra_' + Date.now(),
            nome: nome,
            data: new Date().toLocaleString(),
            dados_json: JSON.stringify(tableStates)
        };

        try {
            await fetch('/api/obras', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(obra)
            });
            modalSaveObra.classList.add('hidden');
            showToast('Obra salva com sucesso!');
        } catch (e) {
            showToast('Erro ao salvar no backend');
        }
    });

    document.getElementById('btn-load-obra').addEventListener('click', async () => {
        try {
            const res = await fetch('/api/obras');
            if (!res.ok) throw new Error('Falha ao listar obras');
            const obras = await res.json();
            
            const listDiv = document.getElementById('obras-list');
            listDiv.innerHTML = '';
            
            if (obras.length === 0) {
                listDiv.innerHTML = '<div style="color:#8b949e;padding:10px;">Nenhuma obra salva ainda.</div>';
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
                            <span class="obra-nome">${o.nome}</span>
                            <span class="obra-data">${o.data}</span>
                        </div>
                        <div class="obra-actions">
                            <button class="btn-primary btn-load-item"  style="background:#238636; border:none; padding: 4px 8px; border-radius: 4px; color: white;">Carregar</button>
                            <button class="btn-secondary btn-add-item"  style="background:#1f6feb; border:none; padding: 4px 8px; border-radius: 4px; color: white;">Adicionar</button>
                            <button class="btn-secondary btn-sub-item"  style="background:#d29922; border:none; padding: 4px 8px; border-radius: 4px; color: white;">Subtrair</button>
                            <button class="btn-secondary btn-del-item" data-id="${o.id}" style="background:#da3633; border:none; padding: 4px 8px; border-radius: 4px; color: white;">Excluir</button>
                        </div>
                    `;

                    // Guarda o snap diretamente no elemento via propriedade JS (sem HTML encoding)
                    const btnLoad = item.querySelector('.btn-load-item');
                    const btnAdd  = item.querySelector('.btn-add-item');
                    const btnSub  = item.querySelector('.btn-sub-item');
                    const btnDel  = item.querySelector('.btn-del-item');

                    btnLoad._snap = snap;
                    btnAdd._snap  = snap;
                    btnSub._snap  = snap;

                    btnLoad.addEventListener('click', () => {
                        if (!btnLoad._snap) { showToast('Dados da obra inválidos.'); return; }
                        restoreObraSnapshot(btnLoad._snap);
                        modalLoadObra.classList.add('hidden');
                        showToast('Obra carregada!');
                    });

                    btnAdd.addEventListener('click', () => {
                        if (!btnAdd._snap) { showToast('Dados da obra inválidos.'); return; }
                        const s = btnAdd._snap;
                        pushHistory();  // snapshot ANTES de adicionar
                        ['cabos', 'outros'].forEach(type => {
                            if (s[type] && s[type].data) {
                                s[type].data.forEach(row => {
                                    if (row) tableStates[type].data.push(JSON.parse(JSON.stringify(row)));
                                });
                                if (type === 'cabos') recalcAllQtdAtivos();
                                renderTable(type);
                                refreshAllFilters(type);
                            }
                        });
                        buildAtivoSets();
                        buildDataLists();
                        modalLoadObra.classList.add('hidden');
                        showToast('Obra adicionada!');
                    });

                    btnSub.addEventListener('click', () => {
                        if (!btnSub._snap) { showToast('Dados da obra inválidos.'); return; }
                        const s = btnSub._snap;
                        pushHistory();  // snapshot ANTES de subtrair
                        ['cabos', 'outros'].forEach(type => {
                            if (s[type] && s[type].data) {
                                s[type].data.forEach(row => {
                                    if (row) {
                                        let clonedRow = JSON.parse(JSON.stringify(row));
                                        if (type === 'cabos') {
                                            clonedRow.qtdAtivos = "-" + Math.abs(clonedRow.qtdAtivos || 0);
                                        } else {
                                            clonedRow.ativo = "*" + clonedRow.ativo;
                                        }
                                        tableStates[type].data.push(clonedRow);
                                    }
                                });
                                renderTable(type);
                                refreshAllFilters(type);
                            }
                        });
                        buildAtivoSets();
                        buildDataLists();
                        modalLoadObra.classList.add('hidden');
                        showToast('Obra subtraída!');
                    });

                    btnDel.addEventListener('click', async () => {
                        btnDel.disabled = true;
                        try {
                            await fetch(`/api/obras/${o.id}`, { method: 'DELETE' });
                            item.remove();
                        } catch (err) {
                            btnDel.disabled = false;
                        }
                    });

                    listDiv.appendChild(item);
                });
            }
            modalLoadObra.classList.remove('hidden');
        } catch (e) {
            console.error('Erro ao carregar obras:', e);
            showToast('Erro ao carregar obras');
        }
    });

    const btnMontarOrcamento = document.getElementById('btn-montar-orcamento');
    const modalOrcamento = document.getElementById('modal-orcamento');
    const tbodyOrcamento = document.getElementById('resultado-orcamento-tbody');

    if (btnMontarOrcamento) {
        btnMontarOrcamento.addEventListener('click', async () => {
            const selectProj = document.getElementById('select-projeto');
            const projVal = selectProj ? selectProj.value : "";
            const projCode = selectProj && selectProj.selectedIndex >= 0 ? selectProj.options[selectProj.selectedIndex].dataset.codigo : "";
            
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
            window.open('/static/resultado_orcamento.html', '_blank');
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
        
        recalcAllQtdAtivos();
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

        recalcAllQtdAtivos();
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
                addAtivo(window._ramaisAtivosInstalando, 'MAA', qtyTroca * 1);
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

window.carregarRegras = function() {
    try {
        const projCode = localStorage.getItem('projeto_selecionado_codigo') || 'DEFAULT';
        const key = `regras_orcamento_${projCode}`;
        const saved = localStorage.getItem(key);
        if (saved) {
            tableStates.regras.data = JSON.parse(saved);
        } else {
            const legacy = localStorage.getItem('regras_orcamento');
            if (legacy && projCode === 'DEFAULT') {
                tableStates.regras.data = JSON.parse(legacy);
            } else {
                tableStates.regras.data = [{
                    origem: 'CABOS', op_de: '', ativo_de: '',
                    op_para: '', ativo_para: '', fator: 1.05,
                    arredondamento: 'NORMAL', val_min: '', val_max: ''
                }];
            }
        }
        renderRegrasTable();
    } catch (e) {
        console.error("Erro ao carregar regras", e);
    }
};

window.salvarRegras = function() {
    const projCode = localStorage.getItem('projeto_selecionado_codigo') || 'DEFAULT';
    localStorage.setItem(`regras_orcamento_${projCode}`, JSON.stringify(tableStates.regras.data));
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
                    acao: hasAcao ? cols[3].toUpperCase() : 'ADIÇÃO',
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
                </select>
            </td>
            <td><input type="text" class="tot-input" placeholder="Qqlr" value="${r.op_de || ''}" onchange="updateRegraRow(${index}, 'op_de', this.value)"></td>
            <td><input type="text" class="tot-input" placeholder="Qqlr" value="${r.ativo_de || ''}" onchange="updateRegraRow(${index}, 'ativo_de', this.value)"></td>
            <td>
                <select class="tot-input" onchange="updateRegraRow(${index}, 'acao', this.value)">
                    <option value="ADIÇÃO" ${r.acao !== 'SUBSTITUIÇÃO' ? 'selected' : ''}>ADIÇÃO</option>
                    <option value="SUBSTITUIÇÃO" ${r.acao === 'SUBSTITUIÇÃO' ? 'selected' : ''}>SUBST.</option>
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
        origem: '', op_de: '', ativo_de: '', acao: 'ADIÇÃO',
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
        }
        tableStates.regras.data[index][field] = value;
        salvarRegras();
    }
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
        const row = orcamentoBaseData.find(r => (r.ativo || '').toUpperCase() === ativoUpper || (r.codigo || '').toUpperCase() === ativoUpper);
        if (row) {
            return row.desc_ativo || row.componente || row.desc_codigo || '';
        }
        return '';
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
        
        let formatado = at.toUpperCase().replace("CAA ", "CAA").replace("CA ", "CA").replace("CU ", "CU").replace("CAZ ", "CAZ").replace("P ", "P");
        let parts = formatado.trim().split(/\s+/);
        let baseAtivo = parts[0] || at;
        let desc = buscarDescricao(baseAtivo);

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
                desc: desc,
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
            const m = p.match(/^([\*\-]?\d+(?:\.\d+)?)[Xx\-](.+)$/i);
            if (m) {
                let qStr = m[1];
                let isNegative = false;
                if (qStr.startsWith('*') || qStr.startsWith('-')) {
                    isNegative = true;
                    qStr = qStr.substring(1);
                }
                q = parseFloat(qStr);
                if (isNegative) q = -q;
                aName = m[2].trim();
            }
            
            let desc = buscarDescricao(aName);

            rawItems.push({
                baseId: `TOT-${currentBaseId}`,
                obs: item.entidade,
                operacao: item.operacao || 'I',
                ativo: aName,
                qtd: q,
                desc: desc, 
                origem: 'OUTROS'
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
            let applies = true;
            if (ruleOrigem !== '' && ruleOrigem !== item.origem) applies = false;
            if (ruleOp !== '' && ruleOp !== item.operacao) applies = false;
            if (ruleAtivo !== '' && ruleAtivo !== item.ativo) applies = false;

            if (applies) {
                matchEncontrado = true;
                if (rule.acao === 'SUBSTITUIÇÃO') {
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
                
                newData.push({
                    id: item.baseId,
                    obs: item.obs,
                    operacao: newOp,
                    ativo: newAtivo,
                    qtd: newQtd,
                    desc: buscarDescricao(newAtivo),
                    origem: item.origem
                });
            }
        });
        
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
            const row = orcamentoBaseData.find(r => 
                (r.ativo || '').toUpperCase() === ativoUpper || 
                (r.componente || '').toUpperCase() === ativoUpper ||
                (r.codigo || '').toUpperCase() === ativoUpper
            );
            
            if (row) {
                let desc = '';
                if ((row.codigo || '').toUpperCase() === ativoUpper) {
                    desc = row.desc_codigo || '';
                } else {
                    desc = row.desc_ativo || '';
                }
                
                tableStates.totalizadora.data[index].desc = desc;
                
                // Atualiza visualmente o input de descrição sem dar re-render na tabela para não perder o foco atual do usuário
                const descInput = document.getElementById(`tot-desc-${index}`);
                if (descInput) {
                    descInput.value = desc;
                }
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
