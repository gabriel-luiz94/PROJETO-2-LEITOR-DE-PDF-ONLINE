document.addEventListener('DOMContentLoaded', () => {
    let extractedDataCache = [];
    const tableStates = {
        cabos: { bodyId: 'body-cabos', data: [], filters: { 0: new Set(), 1: new Set(), 2: new Set(), 3: new Set(), 4: new Set(), 5: new Set() } },
        outros: { bodyId: 'body-outros', data: [], filters: { 0: new Set(), 1: new Set(), 2: new Set(), 3: new Set(), 4: new Set(), 5: new Set() } }
    };

    const dataRaw = localStorage.getItem('processar_dados');
    if (!dataRaw) {
        alert("Nenhum dado encontrado.");
        return;
    }

    extractedDataCache = JSON.parse(dataRaw);
    tableStates.cabos.data = extractedDataCache.filter(d => d.entidade === "CABO");
    tableStates.outros.data = extractedDataCache.filter(d => d.entidade !== "CABO" && d.entidade !== "0");

    initTable('cabos');
    initTable('outros');

    function initTable(type) {
        renderTable(type);
        updateCounters();
        setupFilters(type);
    }

    function isGray(hex) {
        if (!hex || hex.length < 7) return false;
        const r = parseInt(hex.substring(1, 3), 16);
        const g = parseInt(hex.substring(3, 5), 16);
        const b = parseInt(hex.substring(5, 7), 16);
        return Math.abs(r-g)<5 && Math.abs(g-b)<5 && r>20 && r<230;
    }

    function processAtivoFormula(col2) {
        if (!col2) return "";
        let normalized = col2.replace(/[\u00AD\u2010-\u2015\u2212]/g, "-").replace(/\u00A0/g, " ");
        const cleanArrume = (t) => t.replace(/\s*-\s*/g, "-").trim().replace(/\s+/g, ' ');
        const tArrume = cleanArrume(normalized);
        const tUpperMatched = tArrume.toUpperCase();
        if (tUpperMatched.includes("AFASTADOR")) return "1-AF";
        let text1 = "";
        const instMatch = tUpperMatched.match(/INST\.?(?:AL(?:AR|A)?)?\s+0*(\d+)\s*(?:-| )?\s*(\d*SI\d*|\d*RA\d*|\d*BI\d*|\d*R\d*|\d*B\d*|\d*S\d*|\d*CE\d*|\d*N\d*|\d*U\d*|\d*T\d*|ISOL)/);
        if (instMatch) {
            text1 = instMatch[1] + "-" + instMatch[2].replace(/\s/g, "");
        } else if (tUpperMatched.includes(" METROS")) {
            const match = tUpperMatched.match(/(\d+(?:[.,]\d+)?)\s*METROS/);
            if (match) text1 = match[1] + "-ROCO";
            else text1 = tArrume.replace(/ METROS/gi, "-ROCO");
        } else if (/CAL[CÇ]ADA|RECAL|REC\.\s*CAL/i.test(tUpperMatched)) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            const qty = matchQty ? matchQty[1] : "1";
            text1 = qty + "-RECAL";
        } else if (/\bIP\b/i.test(tUpperMatched)) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            const qty = matchQty ? matchQty[1] : "1";
            text1 = qty + "-IP";
        } else if (tUpperMatched.includes("CONC") && tUpperMatched.includes("BASE")) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            const qty = matchQty ? matchQty[1] : "1";
            text1 = qty + "-BASE";
        } else if (/\bCOMPRESSOR\b/i.test(tUpperMatched) || /\bCAVA\b/i.test(tUpperMatched)) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            const qty = matchQty ? matchQty[1] : "1";
            text1 = qty + "-CAVA";
        } else {
            if (!tUpperMatched.includes("FIOS")) text1 = tArrume;
            else {
                const arrumeRaw = col2.trim().replace(/\s+/g, ' ');
                text1 = "1-" + arrumeRaw.substring(0, Math.max(0, arrumeRaw.length - 5)) + "F ";
            }
        }
        let text2 = text1.replace(/TR\s*-\s*3\s*-\s*/g, "1-TR3").replace(/kVA/gi, "").replace(/TR\s*-\s*1\s*-\s*/g, "1-TR1").replace(/TR\s*-\s*2\s*-\s*/g, "1-TR2");
        
        // Regra para padrão combinado: 3 - 100A - 3H
        const comboMatch = text2.match(/([13])\s*-\s*100A\s*-\s*(\d+(?:,\d+)?(?:H|K))/i);
        if (comboMatch) {
            const q = comboMatch[1];
            const e = comboMatch[2].replace(",", "").replace(" ", "");
            return `${q}-CFU ${q}-EF${e}`;
        }

        if (text2.includes("-100A-")) text2 = text2.replace("-100A-", `-CFU ${text2.charAt(0)}-EF`);
        return text2.replace(/3CF-400A/g, "3-CFA").replace(/112,5/g, "112").replace(/0,5H/g, "05H").replace(/PODAS/g, "PODA").replace(/PODA M/g, "PODA").replace(/PODA G/g, "PODA").replace(/PODA P/g, "PODA").replace(/ PODA/g, "-PODA").replace(/3CL-300A/g, "3-CFU 3-CL").replace(/TR15/g, "TR105").trim();
    }

    function updateRowLogic(item, tr) {
        const displayColor = item.cor || "#000000";
        let textoAtivo = processAtivoFormula(item.texto);
        let opAuto = "M";
        if (displayColor.toUpperCase() === "#FF0000") opAuto = "I";
        else if (isGray(displayColor)) opAuto = "R";
        
        const uAtivo = textoAtivo.toUpperCase();
        const tUpper = item.texto.replace(/\u00A0/g, " ").toUpperCase().trim();
        let entAuto = "0";

        const isRedOrGray = (displayColor.toUpperCase() === "#FF0000") || (isGray(displayColor));
        
        // --- NOVA REGRA: ELO FUSÍVEL (EF) ---
        const elosFusivel = ["0,5H", "1H", "2H", "3H", "5H", "6K", "8K", "10K", "12K", "15K", "25K", "30K", "40K"];
        if (elosFusivel.includes(tUpper) && isRedOrGray) {
            entAuto = "CHAVE";
            // Busca contexto nas proximidades (mesma página, +/- 10 linhas no cache global)
            const globalIdx = extractedDataCache.indexOf(item);
            if (globalIdx !== -1) {
                let qtyFound = "";
                const searchRange = 10;
                const start = Math.max(0, globalIdx - searchRange);
                const end = Math.min(extractedDataCache.length - 1, globalIdx + searchRange);
                
                for (let i = start; i <= end; i++) {
                    if (i === globalIdx) continue;
                    const neighbor = extractedDataCache[i];
                    if (neighbor.pagina !== item.pagina) continue;
                    
                    const nText = neighbor.texto.replace(/\u00A0/g, " ").toUpperCase();
                    const m = nText.match(/([13])\s*-\s*100A/);
                    if (m) {
                        qtyFound = m[1];
                        break;
                    }
                }
                
                if (qtyFound) {
                    const eloVal = tUpper.replace(",", "").replace(" ", ""); // 0,5H -> 05H
                    textoAtivo = `${qtyFound}-EF${eloVal}`;
                }
            }
        }

        entAuto = "0";
        const isRed = (displayColor.toUpperCase() === "#FF0000");

        if (!tUpper.includes("BLOCO")) {
            if (isRed && tUpper.includes("FLY")) {
                if (tUpper.includes("REF")) { textoAtivo = "1-RFLY"; opAuto = "I"; entAuto = "ESTRUTURA"; }
                else if (tUpper.includes("DESF")) { textoAtivo = "1-FLY"; opAuto = "R"; entAuto = "ESTRUTURA"; }
                else if (tUpper.includes("INST")) { textoAtivo = "1-FLY"; opAuto = "I"; entAuto = "ESTRUTURA"; }
            }
        } else { textoAtivo = ""; opAuto = "M"; entAuto = "0"; }
        
        const apoioMarkers = ["-ROCO", "-RECAL", "-BASE", "-CAVA", "PODA"];
        const foundApoioCount = apoioMarkers.filter(m => uAtivo.includes(m)).length;
        const hasApoioConflict = tUpper.includes("APOIOS") || tUpper.includes("LARGURA") || (tUpper.includes("BASE") && tUpper.includes("CALÇADA")) || foundApoioCount > 1;
        
        if (entAuto === "0") {
            if (/\sRS\s+[MT]\s/i.test(item.texto)) entAuto = "RAMAIS";
            else if (uAtivo.includes("-ROCO") && !hasApoioConflict) entAuto = "APOIO";
            else if (uAtivo.includes("-IP") || /\bIP\b/i.test(item.texto)) entAuto = "IP";
            else if (uAtivo === "1-AF") entAuto = "ESTRUTURA";
            else if ((uAtivo.includes("-RECAL")||uAtivo.includes("-BASE")||uAtivo.includes("-CAVA")) && !hasApoioConflict) entAuto = "APOIO";
            else if ((tUpper.includes("DT") || tUpper.includes("CV")) && tUpper.includes("/") && isRedOrGray) entAuto = "POSTE";
            else if (!tUpper.includes("DT") && !tUpper.includes("CV") && !tUpper.includes("AWG") && !tUpper.includes("#") && (() => {
                if (/\bM?\d+x\d+/.test(tUpper)) return true;
                if (tUpper.includes("ABC") && /\d+\s*M$/.test(tUpper)) return true;
                if (/^CU\s*\d/.test(tUpper) || /\bCU\s*\d/.test(tUpper.substring(0, 5))) return true;
                if (/^CA\s+\d/.test(tUpper) || /^CA\d/.test(tUpper)) return true;
                if (/(?:CAL|CAA|CAZ)(?:\s|\d)/.test(tUpper)) return true;
                if (/\bP\s*(16|25|35|50|70|95|120|150|185|240)\b/.test(tUpper)) return true;
                if (tUpper.includes("X1X") && /\d+\s*M$/.test(tUpper)) return true;
                return false;
            })()) {
                if (isRedOrGray) {
                    entAuto = "CABO";
                } else if (!isRedOrGray && (item.layer === "RETENS" || item.layer === "RETENS_LV")) {
                    entAuto = "CABO";
                    opAuto = item.layer === "RETENS_LV" ? "*M" : "M";
                }
            }
            else if (tUpper.includes("FIOS")) entAuto = "CERCA";
            else if ((uAtivo.includes("-CF") || uAtivo.includes("-EF")) && opAuto !== "M") entAuto = "CHAVE";
            else if (uAtivo.includes("-TR") && opAuto !== "M") {
                entAuto = "TRAFO";
                const match = textoAtivo.match(/(1-TR\d+)/i);
                if (match) textoAtivo = match[1].toUpperCase();
            }
            else if (uAtivo.includes("PODA") && opAuto !== "M" && !hasApoioConflict) entAuto = "APOIO";
        }
        if (entAuto === "IP" || entAuto === "APOIO" || entAuto === "CERCA" || entAuto === "RAMAIS") opAuto = "I";
        if (entAuto === "0" && !tUpper.includes("APOIOS")) {
            const hasEstruturaPattern = /\b\d+\s*-\s*(\d+[A-Z]{1,2}\d*|\d*[A-Z]{1,2}\d+|ISOL)\b/i.test(uAtivo) || /\b\d+\s*-\s*(SI|RA|BI|CE|N|U|T|R|B|S)\d+/i.test(uAtivo);
            const words = textoAtivo.trim().split(/\s+/);
            const allWordsValid = words.length > 0 && words.every(w => /^\d+-\S+$/i.test(w));
            if (hasEstruturaPattern && allWordsValid && (opAuto === "I" || opAuto === "R")) entAuto = "ESTRUTURA";
        }
        
        if (item.layer === "LV" && opAuto && !opAuto.startsWith("*")) {
            opAuto = "*" + opAuto;
        }

        item.entidade = entAuto; item.operacao = opAuto; item.ativo = textoAtivo;
        const entInput = tr.querySelector('[data-field="entidade"]');
        const opInput = tr.querySelector('[data-field="operacao"]');
        const atInput = tr.querySelector('[data-field="ativo"]');
        if(entInput) entInput.value = entAuto;
        if(opInput) opInput.value = opAuto;
        if(atInput) atInput.value = textoAtivo;
    }

    function renderTable(type) {
        const state = tableStates[type];
        const body = document.getElementById(state.bodyId);
        body.innerHTML = '';

        state.data.forEach((item, index) => {
            const tr = document.createElement('tr');
            tr.dataset.index = index;
            tr.dataset.type = type;
            const curColor = item.cor ? item.cor.toUpperCase() : "#000000";
            tr.innerHTML = `
                <td>${item.pagina || 1}</td>
                <td style="word-break: break-all; max-width: 400px;">
                    <div class="text-cell-container">
                        <span class="selectable-text editable-text-field" contenteditable="false">${escapeHtml(item.texto)}</span>
                    </div>
                </td>
                <td>
                    <select class="color-select" style="background: transparent; border: none; color: ${curColor}; font-weight: bold; cursor: pointer; outline: none;">
                        <option value="#000000" style="color: black;" ${curColor==='#000000'?'selected':''}>PRETO</option>
                        <option value="#FF0000" style="color: red;" ${curColor==='#FF0000'?'selected':''}>VERMELHO</option>
                        <option value="#808080" style="color: gray;" ${isGray(curColor)?'selected':''}>CINZA</option>
                    </select>
                </td>
                <td class="col-user"><input class="user-input" type="text" value="${item.entidade}" data-field="entidade"></td>
                <td class="col-user"><input class="user-input" type="text" value="${item.operacao}" data-field="operacao"></td>
                <td class="col-user"><input class="user-input" type="text" value="${item.ativo}" data-field="ativo"></td>
            `;
            body.appendChild(tr);

            const textField = tr.querySelector('.editable-text-field');
            textField.addEventListener('dblclick', function() { this.contentEditable = true; this.focus(); });
            textField.addEventListener('blur', function() {
                this.contentEditable = false;
                item.texto = this.innerText.trim();
                updateRowLogic(item, tr);
            });
            textField.addEventListener('keydown', function(e) { if(e.key==='Enter') { e.preventDefault(); this.blur(); } });

            const colorSelect = tr.querySelector('.color-select');
            colorSelect.addEventListener('change', (e) => {
                item.cor = e.target.value;
                colorSelect.style.color = item.cor;
                updateRowLogic(item, tr);
            });

            tr.querySelectorAll('.user-input').forEach(input => {
                input.addEventListener('input', (e) => {
                    item[e.target.dataset.field] = e.target.value;
                });
            });

            tr.addEventListener('contextmenu', (e) => {
                e.preventDefault();
                showContextMenu(e, tr);
            });
        });
        applyFilters(type);
    }

    function updateCounters() {
        document.getElementById('count-cabos').textContent = tableStates.cabos.data.length;
        document.getElementById('count-outros').textContent = tableStates.outros.data.length;
    }

    function setupFilters(type) {
        const containers = document.querySelectorAll(`.filter-container[data-table="${type}"]`);
        const state = tableStates[type];
        const columns = ['pagina', 'texto', 'cor', 'entidade', 'operacao', 'ativo'];
        containers.forEach(container => {
            const colIdx = parseInt(container.dataset.col);
            const uniqueValues = [...new Set(state.data.map(item => String(item[columns[colIdx]] || "")))].sort();
            createFilterDropdown(container, uniqueValues, colIdx, type);
        });
    }

    function createFilterDropdown(container, values, colIdx, type) {
        const state = tableStates[type];
        const selections = state.filters[colIdx];
        container.innerHTML = `
            <div class="filter-trigger"><span>Todos</span><div class="filter-active-indicator"></div></div>
            <div class="filter-dropdown">
                <div class="filter-search-container"><input type="text" class="filter-search" placeholder="Pesquisar..."></div>
                <div class="filter-options-list">
                    <label class="filter-option select-all-option"><input type="checkbox" checked><span>(Selecionar Tudo)</span></label>
                    <div class="options-container">
                        ${values.map(val => `<label class="filter-option" data-value="${val}"><input type="checkbox" checked><span>${val || "(Vazio)"}</span></label>`).join('')}
                    </div>
                </div>
            </div>
        `;
        const trigger = container.querySelector('.filter-trigger'), dropdown = container.querySelector('.filter-dropdown');
        const mainSelectAll = container.querySelector('.select-all-option input'), optionsContainer = container.querySelector('.options-container');
        trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            document.querySelectorAll('.filter-dropdown.active').forEach(d => {
                if (d !== dropdown) {
                    d.classList.remove('active');
                    // Aqui aplicamos o filtro da tabela correspondente ao fechar outros dropdowns
                    const otherTable = d.closest('.filter-container').dataset.table;
                    applyFilters(otherTable);
                }
            });
            if (dropdown.classList.contains('active')) {
                dropdown.classList.remove('active');
                applyFilters(type);
            } else {
                dropdown.classList.add('active');
            }
        });
        dropdown.addEventListener('click', e => e.stopPropagation());
        const onSelectionChange = () => {
            const checks = optionsContainer.querySelectorAll('input'), checked = optionsContainer.querySelectorAll('input:checked');
            mainSelectAll.checked = checked.length === checks.length;
            mainSelectAll.indeterminate = checked.length > 0 && checked.length < checks.length;
            selections.clear();
            if (checked.length < checks.length) { checked.forEach(c => selections.add(c.parentElement.dataset.value)); }
            // applyFilters(type); // REMOVIDO: Agora aplica apenas ao fechar
            updateFilterTrigger(container, selections);
        };
        optionsContainer.querySelectorAll('input').forEach(cb => cb.addEventListener('change', onSelectionChange));
        mainSelectAll.addEventListener('change', () => {
            optionsContainer.querySelectorAll('input').forEach(cb => { cb.checked = mainSelectAll.checked; });
            onSelectionChange();
        });
    }

    function updateFilterTrigger(container, selections) {
        const span = container.querySelector('.filter-trigger span'), ind = container.querySelector('.filter-active-indicator');
        span.textContent = selections.size > 0 ? selections.size + ' sel' : 'Todos';
        ind.classList.toggle('active', selections.size > 0);
    }

    function applyFilters(type) {
        const state = tableStates[type], rows = document.getElementById(state.bodyId).children;
        const columns = ['pagina', 'texto', 'cor', 'entidade', 'operacao', 'ativo'];
        for (let tr of rows) {
            const item = state.data[tr.dataset.index];
            let isVisible = true;
            for (let i = 0; i < 6; i++) {
                const sel = state.filters[i];
                if (sel.size > 0) { const val = String(item[columns[i]] || ""); if (!sel.has(val)) { isVisible = false; break; } }
            }
            tr.style.display = isVisible ? '' : 'none';
        }
    }

    const contextMenu = document.getElementById('context-menu');
    let contextTargetRow = null;
    function showContextMenu(e, tr) {
        contextTargetRow = tr;
        const x = Math.min(e.clientX, window.innerWidth - 180), y = Math.min(e.clientY, window.innerHeight - 180);
        contextMenu.style.left = x + 'px'; contextMenu.style.top = y + 'px';
        contextMenu.classList.add('active');
    }
    document.addEventListener('click', () => contextMenu.classList.remove('active'));

    document.getElementById('ctx-delete').addEventListener('click', () => {
        if (!contextTargetRow) return;
        const type = contextTargetRow.dataset.type, index = parseInt(contextTargetRow.dataset.index);
        tableStates[type].data.splice(index, 1);
        renderTable(type); updateCounters();
    });

    document.getElementById('ctx-edit').addEventListener('click', () => {
        if (!contextTargetRow) return;
        const span = contextTargetRow.querySelector('.editable-text-field');
        span.contentEditable = true; span.focus();
    });

    function addRow(type, index = -1) {
        const newItem = {
            pagina: (contextTargetRow ? contextTargetRow.children[0].innerText : "1"),
            texto: "",
            cor: "#FF0000",
            entidade: "0",
            operacao: "I",
            ativo: ""
        };
        if (index === -1) tableStates[type].data.push(newItem);
        else tableStates[type].data.splice(index + 1, 0, newItem);
        
        renderTable(type);
        updateCounters();

        // Foca na célula de texto da nova linha
        const body = document.getElementById(tableStates[type].bodyId);
        const targetRow = (index === -1) ? body.lastElementChild : body.children[index + 1];
        if (targetRow) {
            const span = targetRow.querySelector('.editable-text-field');
            span.contentEditable = true;
            span.focus();
        }
    }

    document.getElementById('ctx-add').addEventListener('click', () => {
        if (!contextTargetRow) return;
        addRow(contextTargetRow.dataset.type, parseInt(contextTargetRow.dataset.index));
    });

    document.querySelectorAll('.btn-add').forEach(btn => {
        btn.addEventListener('click', () => {
            const type = btn.dataset.target.includes('cabos') ? 'cabos' : 'outros';
            addRow(type);
        });
    });

    let isSelecting = false, selectionStart = null, selectionEnd = null, activeType = null;
    document.querySelectorAll('tbody').forEach(body => {
        const type = body.id.includes('cabos') ? 'cabos' : 'outros';
        body.addEventListener('mousedown', (e) => {
            if (e.button !== 0 || e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT') return;
            isSelecting = true; activeType = type;
            selectionStart = getCoords(e.target.closest('td'), body);
            selectionEnd = selectionStart;
            clearSelection(); updateHighlight();
        });
        body.addEventListener('mouseover', (e) => {
            if (!isSelecting || activeType !== type) return;
            const td = e.target.closest('td');
            if (td) { selectionEnd = getCoords(td, body); updateHighlight(); }
        });
    });
    document.addEventListener('mouseup', () => isSelecting = false);
    function getCoords(td, body) { const tr = td.parentElement; return { row: Array.from(body.children).indexOf(tr), col: Array.from(tr.children).indexOf(td) }; }
    function clearSelection() { document.querySelectorAll('.cell-selected').forEach(c => c.classList.remove('cell-selected')); }
    function updateHighlight() {
        if (!selectionStart || !selectionEnd) return;
        clearSelection();
        const body = document.getElementById(tableStates[activeType].bodyId);
        const rMin = Math.min(selectionStart.row, selectionEnd.row), rMax = Math.max(selectionStart.row, selectionEnd.row);
        const cMin = Math.min(selectionStart.col, selectionEnd.col), cMax = Math.max(selectionStart.col, selectionEnd.col);
        for (let r = rMin; r <= rMax; r++) {
            const tr = body.children[r];
            if (tr && tr.style.display !== 'none') {
                for (let c = cMin; c <= cMax; c++) { if (tr.children[c]) tr.children[c].classList.add('cell-selected'); }
            }
        }
    }

    function getVisibleRows(type) {
        if (!type) return [];
        const body = document.getElementById(tableStates[type].bodyId);
        return Array.from(body.children).filter(tr => tr.style.display !== 'none');
    }
    function getVisibleRowIndex(domIdx, type) {
        const body = document.getElementById(tableStates[type].bodyId);
        const vis = getVisibleRows(type);
        return vis.indexOf(body.children[domIdx]);
    }
    function getDomRowIndex(visIdx, type) {
        const vis = getVisibleRows(type);
        const body = document.getElementById(tableStates[type].bodyId);
        return Array.from(body.children).indexOf(vis[visIdx]);
    }

    document.addEventListener('copy', (e) => {
        const selected = document.querySelectorAll('.cell-selected');
        if (selected.length > 0 && document.activeElement.tagName !== 'INPUT') {
            copySelectionToClipboard();
            e.preventDefault();
        }
    });

    function copyToClipboard(text, callback) {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(text).then(callback).catch(err => {
                console.error("Erro ao copiar via API:", err);
                fallbackCopy(text, callback);
            });
        } else {
            fallbackCopy(text, callback);
        }
    }

    function fallbackCopy(text, callback) {
        const textArea = document.createElement("textarea");
        textArea.value = text;
        textArea.style.position = "fixed";
        textArea.style.left = "-9999px";
        textArea.style.top = "0";
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        try {
            const successful = document.execCommand('copy');
            if (successful && callback) callback();
        } catch (err) {
            console.error('Erro no fallback de cópia:', err);
        }
        document.body.removeChild(textArea);
    }

    function copySelectionToClipboard() {
        const selected = document.querySelectorAll('.cell-selected');
        if (selected.length === 0) return;
        
        let text = "", lastRow = -1;
        const sorted = Array.from(selected).sort((a, b) => {
            const rowA = a.parentElement, rowB = b.parentElement;
            if (rowA === rowB) return a.cellIndex - b.cellIndex;
            return rowA.rowIndex - rowB.rowIndex;
        });

        sorted.forEach(cell => {
            const tr = cell.parentElement;
            const row = tr.rowIndex;
            if (lastRow !== -1 && row !== lastRow) text += "\n";
            else if (lastRow === row) text += "\t";
            
            const val = cell.querySelector('input')?.value || cell.querySelector('select')?.value || cell.innerText;
            text += String(val || "").replace(/\n/g, ' ');
            lastRow = row;
        });
        copyToClipboard(text.trim());
    }

    document.addEventListener('keydown', (e) => {
        // Ctrl+C / Cmd+C: copy selection
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
            const selected = document.querySelectorAll('.cell-selected');
            if (selected.length > 0 && e.target.tagName !== 'INPUT') {
                e.preventDefault();
                copySelectionToClipboard();
            }
            return;
        }

        // Shift+Arrow Navigation
        const arrowKeys = ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'];
        if (e.shiftKey && arrowKeys.includes(e.key) && selectionStart !== null && activeType) {
            // Evita conflito se estiver editando texto
            if (e.target.contentEditable === 'true' || e.target.tagName === 'INPUT') return;
            
            e.preventDefault();
            const vis = getVisibleRows(activeType);
            if (vis.length === 0) return;

            const body = document.getElementById(tableStates[activeType].bodyId);
            const totalCols = body.children[0] ? body.children[0].children.length - 1 : 5;
            const curVisIdx = getVisibleRowIndex(selectionEnd.row, activeType);
            
            let newVisIdx = curVisIdx;
            let newCol = selectionEnd.col;

            if (e.key === 'ArrowDown') {
                newVisIdx = e.ctrlKey ? vis.length - 1 : Math.min(curVisIdx + 1, vis.length - 1);
            } else if (e.key === 'ArrowUp') {
                newVisIdx = e.ctrlKey ? 0 : Math.max(curVisIdx - 1, 0);
            } else if (e.key === 'ArrowRight') {
                newCol = e.ctrlKey ? totalCols : Math.min(selectionEnd.col + 1, totalCols);
            } else if (e.key === 'ArrowLeft') {
                newCol = e.ctrlKey ? 0 : Math.max(selectionEnd.col - 1, 0);
            }

            if (newVisIdx !== -1) {
                selectionEnd = { row: getDomRowIndex(newVisIdx, activeType), col: newCol };
                updateHighlight();

                // Scroll focus
                const lastTr = body.children[selectionEnd.row];
                if (lastTr) lastTr.scrollIntoView({ block: 'nearest' });
            }
            return;
        }
    });

    document.querySelectorAll('.btn-copy-report').forEach(btn => {
        btn.addEventListener('click', () => {
            const type = btn.dataset.target.includes('cabos') ? 'cabos' : 'outros';
            const table = document.getElementById(`table-${type}`), body = document.getElementById(tableStates[type].bodyId);
            const selectedCols = Array.from(table.querySelectorAll('.col-checkbox:checked')).map(cb => parseInt(cb.dataset.col));
            if (selectedCols.length === 0) { alert("Selecione colunas."); return; }
            let text = "";
            Array.from(body.children).forEach(tr => {
                if (tr.style.display !== 'none') {
                    const cells = Array.from(tr.children);
                    const rowData = selectedCols.map(colI => {
                        if (colI === 2) return cells[colI].querySelector('select').value;
                        if (colI < 3) return cells[colI].innerText.replace(/\n/g, ' ');
                        return cells[colI].querySelector('input').value.replace(/\n/g, ' ');
                    });
                    text += rowData.join('\t') + "\n";
                }
            });
            copyToClipboard(text, () => {
                const orig = btn.innerHTML; btn.innerHTML = "Copiado!"; setTimeout(() => btn.innerHTML = orig, 1000);
            });
        });
    });

    function escapeHtml(u) { return String(u || "").replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":"&#039;"}[m])); }
});
