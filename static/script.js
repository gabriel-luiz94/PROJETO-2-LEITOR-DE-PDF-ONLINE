document.addEventListener('DOMContentLoaded', () => {
    // WebSocket for remote file triggers
    let socket;
    try {
        const wsProtocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
        socket = new WebSocket(`${wsProtocol}//${window.location.host}/ws`);
        socket.onmessage = (event) => {
            const data = JSON.parse(event.data);
            if (data.type === 'load_file' && data.path) {
                autoExtractLocal(data.path);
            }
        };
        socket.onerror = (err) => console.warn("WebSocket status: Offline ou bloqueado (uso local apenas)");
    } catch (e) {
        console.warn("WebSocket não pôde ser inicializado:", e);
    }

    const uploadZone = document.getElementById('upload-zone');
    const fileInput = document.getElementById('pdf-file');
    const fileInfo = document.getElementById('file-info');
    const fileName = document.getElementById('file-name');
    const removeBtn = document.getElementById('remove-file');
    const extractBtn = document.getElementById('extract-btn');
    const spinner = document.getElementById('loading-spinner');
    const resultsSection = document.getElementById('results-section');
    const tableBody = document.getElementById('table-body');
    const exportCsvBtn = document.getElementById('export-csv');
    const copySelectedBtn = document.getElementById('copy-selected');
    const filterContainers = document.querySelectorAll('.filter-container');
    const colCheckboxes = document.querySelectorAll('.col-checkbox');

    const columnFilterSelections = { 
        0: new Set(), 1: new Set(), 2: new Set(), 
        3: new Set(), 4: new Set(), 5: new Set() 
    };

    // Close dropdowns when clicking outside — apply filter on close
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.filter-container')) {
            const hadActive = document.querySelectorAll('.filter-dropdown.active').length > 0;
            document.querySelectorAll('.filter-dropdown.active').forEach(d => d.classList.remove('active'));
            if (hadActive) applyFilters();
        }
    });

    let currentFile = null;
    let extractedDataCache = [];
    let userFields = {}; // stores { rowIndex: { entidade, operacao, ativo } }
    let deletedRows = new Set(); // tracks deleted row indices

    // Context menu
    const contextMenu = document.createElement('div');
    contextMenu.id = 'context-menu';
    contextMenu.innerHTML = `
        <div class="ctx-item" id="ctx-copy">
            <svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>
            Copiar seleção
        </div>
        <div class="ctx-item" id="ctx-edit">
            <svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>
            Editar texto
        </div>
        <div class="ctx-separator"></div>
        <div class="ctx-item danger" id="ctx-delete">
            <svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"></path><path d="M10 11v6"></path><path d="M14 11v6"></path></svg>
            Excluir linha
        </div>
    `;
    document.body.appendChild(contextMenu);

    let contextTargetRow = null;

    function hideContextMenu() {
        contextMenu.classList.remove('visible');
        contextTargetRow = null;
    }

    document.addEventListener('mousedown', (e) => {
        if (!contextMenu.contains(e.target)) hideContextMenu();
    });

    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') hideContextMenu(); });

    document.getElementById('ctx-copy').addEventListener('click', (e) => {
        e.stopPropagation();
        copySelectionToClipboard();
        hideContextMenu();
        // Visual feedback
        if (copySelectedBtn) {
            const origHTML = copySelectedBtn.innerHTML;
            copySelectedBtn.innerHTML = "Copiado!";
            copySelectedBtn.style.color = '#3fb950';
            setTimeout(() => { copySelectedBtn.innerHTML = origHTML; copySelectedBtn.style.color = ''; }, 1000);
        }
    });

    document.getElementById('ctx-edit').addEventListener('click', (e) => {
        e.stopPropagation();
        if (contextTargetRow) {
            const textSpan = contextTargetRow.querySelector('.editable-text-field');
            if (textSpan) {
                textSpan.contentEditable = true;
                textSpan.focus();
                const range = document.createRange();
                range.selectNodeContents(textSpan);
                const sel = window.getSelection();
                sel.removeAllRanges();
                sel.addRange(range);
            }
        }
        hideContextMenu();
    });

    document.getElementById('ctx-delete').addEventListener('click', (e) => {
        e.stopPropagation();
        if (contextTargetRow) deleteRow(contextTargetRow);
        hideContextMenu();
    });

    function deleteRow(tr) {
        const idx = tr.dataset.index;
        if (idx !== undefined) deletedRows.add(String(idx));
        tr.classList.add('row-deleted');
        tr.style.display = 'none';
        clearSelection();
        applyFilters();
    }

    // Drag and Drop
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        uploadZone.addEventListener(eventName, (e) => { e.preventDefault(); e.stopPropagation(); }, false);
    });

    uploadZone.addEventListener('dragenter', () => uploadZone.classList.add('dragover'));
    uploadZone.addEventListener('dragover', () => uploadZone.classList.add('dragover'));
    uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('dragover'));
    uploadZone.addEventListener('drop', (e) => {
        uploadZone.classList.remove('dragover');
        handleFiles(e.dataTransfer.files);
    });

    fileInput.addEventListener('change', function() { handleFiles(this.files); });

    function handleFiles(files) {
        if (files.length > 0) {
            const file = files[0];
            const isPDF = file.type === "application/pdf" || file.name.toLowerCase().endsWith('.pdf');
            const isDXF = file.name.toLowerCase().endsWith('.dxf');
            if (isPDF || isDXF) {
                currentFile = file;
                fileName.textContent = file.name;
                uploadZone.classList.add('hidden');
                fileInfo.classList.remove('hidden');
                extractBtn.classList.remove('hidden');
                resultsSection.classList.add('hidden');
            } else {
                alert("Selecione um arquivo PDF ou DXF.");
            }
        }
    }

    removeBtn.addEventListener('click', () => {
        currentFile = null;
        fileInput.value = '';
        uploadZone.classList.remove('hidden');
        fileInfo.classList.add('hidden');
        extractBtn.classList.add('hidden');
        resultsSection.classList.add('hidden');
        extractedDataCache = [];
        userFields = {};
        deletedRows = new Set();
    });

    extractBtn.addEventListener('click', async () => {
        if (!currentFile) return;
        extractBtn.disabled = true;
        extractBtn.querySelector('span').textContent = 'Processando...';
        spinner.classList.remove('hidden');
        try {
            console.log("Iniciando upload de:", currentFile.name);
            const formData = new FormData();
            formData.append('file', currentFile);
            const response = await fetch('/upload', { method: 'POST', body: formData });
            
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Erro no servidor: ${response.status} - ${errorText}`);
            }
            
            const result = await response.json();
            if (result.error) throw new Error(result.error);
            
            console.log("Dados extraídos com sucesso:", result.data.length, "itens");
            extractedDataCache = result.data;
            renderTable(extractedDataCache);
            resultsSection.classList.remove('hidden');
            resultsSection.scrollIntoView({ behavior: 'smooth' });
        } catch (e) { 
            console.error("Erro completo:", e);
            alert("Erro na extração: " + e.message); 
        } finally {
            extractBtn.disabled = false;
            extractBtn.querySelector('span').textContent = 'Extrair Dados';
            spinner.classList.add('hidden');
        }
    });

    function isGray(hex) {
        if (!hex || hex.length < 7) return false;
        const r = parseInt(hex.substring(1, 3), 16);
        const g = parseInt(hex.substring(3, 5), 16);
        const b = parseInt(hex.substring(5, 7), 16);
        return Math.abs(r-g)<5 && Math.abs(g-b)<5 && r>20 && r<230;
    }

    function processAtivoFormula(col2) {
        if (!col2) return "";
        // Transforma todos os tipos de hífens PDF em hífen comum e limpa espaços
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
        }
        else if (/CAL[CÇ]ADA|RECAL|REC\.\s*CAL/i.test(tUpperMatched)) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            const qty = matchQty ? matchQty[1] : "1";
            text1 = qty + "-RECAL";
        }
        else if (/\bIP\b/i.test(tUpperMatched)) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            const qty = matchQty ? matchQty[1] : "1";
            text1 = qty + "-IP";
        }
        else if (tUpperMatched.includes("CONC") && tUpperMatched.includes("BASE")) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            const qty = matchQty ? matchQty[1] : "1";
            text1 = qty + "-BASE";
        }
        else if (/\bCOMPRESSOR\b/i.test(tUpperMatched) || /\bCAVA\b/i.test(tUpperMatched)) {
            const matchQty = tUpperMatched.match(/(\d+)\s*X/i) || tUpperMatched.match(/\(\s*(\d+)/) || tUpperMatched.match(/X\s*(\d+)/i);
            const qty = matchQty ? matchQty[1] : "1";
            text1 = qty + "-CAVA";
        }
        else {
            if (!tUpperMatched.includes("FIOS")) text1 = tArrume;
            else {
                const arrumeRaw = col2.trim().replace(/\s+/g, ' ');
                text1 = "1-" + arrumeRaw.substring(0, Math.max(0, arrumeRaw.length - 5)) + "F ";
            }
        }

        let text2 = text1.replace(/TR\s*-\s*3\s*-\s*/g, "1-TR3")
                         .replace(/kVA/gi, "")
                         .replace(/TR\s*-\s*1\s*-\s*/g, "1-TR1")
                         .replace(/TR\s*-\s*2\s*-\s*/g, "1-TR2");
        if (text2.includes("-100A-")) text2 = text2.replace("-100A-", `-CFU ${text2.charAt(0)}-EF`);
        return text2.replace(/3CF-400A/g, "3-CFA").replace(/112,5/g, "112").replace(/0,5H/g, "05H").replace(/PODAS/g, "PODA").replace(/PODA M/g, "PODA").replace(/PODA G/g, "PODA").replace(/PODA P/g, "PODA").replace(/ PODA/g, "-PODA").replace(/3CL-300A/g, "3-CFU 3-CL").replace(/TR15/g, "TR105").trim();
    }

    function updateRowLogic(index, tr) {
        const item = extractedDataCache[index];
        const displayColor = item.cor || "#000000";
        let textoAtivo = processAtivoFormula(item.texto);
        let opAuto = "M";
        if (displayColor.toLowerCase() === "#ff0000") opAuto = "I";
        else if (isGray(displayColor)) opAuto = "R";
        
        const uAtivo = textoAtivo.toUpperCase();
        const tUpper = item.texto.replace(/\u00A0/g, " ").toUpperCase();

        // REGRAS ESPECÍFICAS PARA FLY
        let entAuto = "0";
        const isRed = (displayColor.toLowerCase() === "#ff0000");

        if (!tUpper.includes("BLOCO")) {
            if (isRed && tUpper.includes("FLY")) {
                if (tUpper.includes("REF")) {
                    textoAtivo = "1-RFLY";
                    opAuto = "I";
                    entAuto = "ESTRUTURA";
                } else if (tUpper.includes("DESF")) {
                    textoAtivo = "1-FLY";
                    opAuto = "R";
                    entAuto = "ESTRUTURA";
                } else if (tUpper.includes("INST")) {
                    textoAtivo = "1-FLY";
                    opAuto = "I";
                    entAuto = "ESTRUTURA";
                }
            }
        } else {
            // Se for BLOCO, resetamos as automações para manual/vazio
            textoAtivo = "";
            opAuto = "M";
            entAuto = "0";
        }
        
        const isRedOrGray = (displayColor.toLowerCase() === "#ff0000") || (isGray(displayColor));
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
            else if (isRedOrGray && !tUpper.includes("DT") && !tUpper.includes("CV") && !tUpper.includes("AWG") && !tUpper.includes("#") && (() => {
                // Multiplexados (M3x1..., 3x1x..., etc)
                if (/\bM?\d+x\d+/.test(tUpper)) return true;
                if (tUpper.includes("ABC") && /\d+\s*M$/.test(tUpper)) return true;
                
                // CU (Cobre)
                if (/^CU\s*\d/.test(tUpper) || /\bCU\s*\d/.test(tUpper.substring(0, 5))) return true;
                // CA / CAL / CAA (Alumínio)
                if (/^CA\s+\d/.test(tUpper) || /^CA\d/.test(tUpper)) return true;
                if (tUpper.includes("CAL") || tUpper.includes("CAA") || tUpper.includes("CAZ")) return true;
                
                // Bitolas P (protoduto) - Somente se seguido de bitola válida
                if (/\bP\s*(16|25|35|50|70|95|120|150|185|240)\b/.test(tUpper)) return true;
                
                // Multiplex padrão antigo
                if (tUpper.includes("X1X") && /\d+\s*M$/.test(tUpper)) return true;
                
                return false;
            })()) entAuto = "CABO";
            else if (tUpper.includes("FIOS")) entAuto = "CERCA";
            else if (uAtivo.includes("-CF") && opAuto !== "M") entAuto = "CHAVE";
            else if (uAtivo.includes("-TR") && opAuto !== "M") {
                entAuto = "TRAFO";
                const match = textoAtivo.match(/(1-TR\d+)/i);
                if (match) textoAtivo = match[1].toUpperCase();
            }
            else if (uAtivo.includes("PODA") && opAuto !== "M" && !hasApoioConflict) entAuto = "APOIO";
        }
        
        // Força "I" para exceções que não dependem da cor (Apoios, Cercas, IP, Ramais)
        if (entAuto === "IP" || entAuto === "APOIO" || entAuto === "CERCA" || entAuto === "RAMAIS") opAuto = "I";

        if (entAuto === "0" && !tUpper.includes("APOIOS")) {
            // Refined regex: prefixes like R, B, S, T, N, U, SI, RA must be followed by at least one digit if they are alone 
            // OR are part of the standard list. We use \d+ suffix for short prefixes to avoid matching RECAL or BASE.
            const hasEstruturaPattern = /\b\d+\s*-\s*(\d+[A-Z]{1,2}\d*|\d*[A-Z]{1,2}\d+|ISOL)\b/i.test(uAtivo) || 
                                       /\b\d+\s*-\s*(SI|RA|BI|CE|N|U|T|R|B|S)\d+/i.test(uAtivo);
            
            // Nova validação: TODAS as palavras em textoAtivo devem ser do tipo 1-SI3 (QUANTIDADE-ATIVO)
            const words = textoAtivo.trim().split(/\s+/);
            const allWordsValid = words.length > 0 && words.every(w => /^\d+-\S+$/i.test(w));
            
            if (hasEstruturaPattern && allWordsValid && (opAuto === "I" || opAuto === "R")) entAuto = "ESTRUTURA";
        }
        userFields[index] = { entidade: entAuto, operacao: opAuto, ativo: textoAtivo };
        const entInput = tr.querySelector('[data-field="entidade"]');
        const opInput = tr.querySelector('[data-field="operacao"]');
        const atInput = tr.querySelector('[data-field="ativo"]');
        if(entInput) entInput.value = entAuto;
        if(opInput) opInput.value = opAuto;
        if(atInput) atInput.value = textoAtivo;
    }

    function renderTable(data) {
        tableBody.innerHTML = '';
        if (!data || data.length === 0) { tableBody.innerHTML = '<tr><td colspan="6" style="text-align:center;">Vazio.</td></tr>'; return; }
        data.forEach((item, index) => {
            const tr = document.createElement('tr');
            tr.dataset.index = index;
            const displayColor = item.cor || "#000000";
            updateRowLogic(index, tr);
            const uf = userFields[index];
            tr.innerHTML = `
                <td>${item.pagina}</td>
                <td style="word-break: break-all; max-width: 300px;">
                    <div class="text-cell-container">
                        <span class="selectable-text editable-text-field">${escapeHtml(item.texto)}</span>
                        <button class="btn-copy-icon" data-clipboard="${escapeHtml(item.texto)}"><svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg></button>
                    </div>
                </td>
                <td><div class="color-badge"><div class="color-swatch" style="background-color: ${displayColor};"></div><span>${displayColor.toUpperCase()}</span></div></td>
                <td class="col-user"><input class="user-input" type="text" value="${uf.entidade}" data-field="entidade" data-row="${index}"></td>
                <td class="col-user"><input class="user-input" type="text" value="${uf.operacao}" data-field="operacao" data-row="${index}"></td>
                <td class="col-user"><input class="user-input" type="text" value="${uf.ativo}" data-field="ativo" data-row="${index}"></td>
            `;
            tableBody.appendChild(tr);

            const textField = tr.querySelector('.editable-text-field');
            textField.addEventListener('blur', function() {
                this.contentEditable = false;
                const newText = this.innerText.trim();
                if (newText !== item.texto) { item.texto = newText; updateRowLogic(index, tr); }
            });
            textField.addEventListener('keydown', function(e) { if (e.key === 'Enter') { e.preventDefault(); this.blur(); } });

            tr.querySelector('.btn-copy-icon').addEventListener('click', function() {
                navigator.clipboard.writeText(item.texto).then(() => {
                    const orig = this.style.color; this.style.color='#3fb950'; setTimeout(()=>this.style.color=orig, 1000);
                });
            });

            tr.querySelectorAll('.user-input').forEach(input => {
                input.addEventListener('input', function() {
                    const rowIdx = this.dataset.row; const field = this.dataset.field;
                    if (!userFields[rowIdx]) userFields[rowIdx] = { entidade:'', operacao:'', ativo:'' };
                    userFields[rowIdx][field] = this.value;
                    refreshUserFilter(field === 'entidade' ? 3 : field === 'operacao' ? 4 : 5);
                });
            });

            tr.addEventListener('contextmenu', (e) => {
                e.preventDefault(); e.stopPropagation(); hideContextMenu();
                contextTargetRow = tr;
                const x = Math.min(e.clientX, window.innerWidth - 200);
                const y = Math.min(e.clientY, window.innerHeight - 150);
                contextMenu.style.left = x + 'px'; contextMenu.style.top = y + 'px';
                setTimeout(() => contextMenu.classList.add('visible'), 10);
            });
        });
        populateFilters(data);
        refreshUserFilter(3); refreshUserFilter(4); refreshUserFilter(5);
    }

    function createFilterDropdown(container, values, colIdx) {
        const colName = ['Página', 'Texto', 'Cor', 'Entidade', 'Operação', 'Ativo'][colIdx];
        const selections = columnFilterSelections[colIdx];
        const isAllChecked = selections.size === 0 || selections.size === values.length;

        container.innerHTML = `
            <div class="filter-trigger" title="Filtrar ${colName}">
                <span>${selections.size > 0 ? selections.size + ' sel' : 'Todos'}</span>
                <div class="filter-active-indicator ${selections.size > 0 ? 'active' : ''}"></div>
            </div>
            <div class="filter-dropdown">
                <div class="filter-search-container">
                    <input type="text" class="filter-search" placeholder="Pesquisar...">
                </div>
                <div class="filter-actions">
                    <button class="btn-filter-action select-all">Limpar</button>
                    <button class="btn-filter-action clear-all">Todos</button>
                </div>
                <div class="filter-options-list">
                    <label class="filter-option select-all-option">
                        <input type="checkbox" ${isAllChecked ? 'checked' : ''}>
                        <span>(Selecionar Tudo)</span>
                    </label>
                    <div class="options-container">
                        ${values.map(val => `
                            <label class="filter-option" data-value="${val}">
                                <input type="checkbox" ${selections.has(String(val)) || selections.size === 0 ? 'checked' : ''}>
                                <span title="${val}">${val === "" ? "(Vazio)" : val}</span>
                            </label>
                        `).join('')}
                    </div>
                </div>
            </div>
        `;

        const trigger = container.querySelector('.filter-trigger');
        const dropdown = container.querySelector('.filter-dropdown');
        const searchInput = container.querySelector('.filter-search');
        const optionsContainer = container.querySelector('.options-container');
        const selectAllBtn = container.querySelector('.btn-filter-action.select-all'); // This actually clears (Select all effectively means clearing the specific filter set)
        const clearBtn = container.querySelector('.btn-filter-action.clear-all');
        const mainSelectAll = container.querySelector('.select-all-option input');

        trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            document.querySelectorAll('.filter-dropdown.active').forEach(d => {
                if (d !== dropdown) {
                    d.classList.remove('active');
                    applyFilters();
                }
            });
            dropdown.classList.toggle('active');
        });

        // Prevention of closing when clicking inside
        dropdown.addEventListener('click', (e) => e.stopPropagation());

        const onSelectionChange = () => {
            const checks = optionsContainer.querySelectorAll('input');
            const checkedCount = optionsContainer.querySelectorAll('input:checked').length;
            mainSelectAll.checked = checkedCount === checks.length;
            
            selections.clear();
            if (checkedCount < checks.length) {
                optionsContainer.querySelectorAll('input:checked').forEach(cb => {
                    selections.add(String(cb.parentElement.dataset.value));
                });
            }
        };

        optionsContainer.querySelectorAll('input').forEach(cb => {
            cb.addEventListener('change', onSelectionChange);
        });

        mainSelectAll.addEventListener('change', (e) => {
            const isChecked = e.target.checked;
            optionsContainer.querySelectorAll('.filter-option').forEach(opt => {
                if(opt.style.display !== 'none') {
                    opt.querySelector('input').checked = isChecked;
                }
            });
            onSelectionChange();
        });

        selectAllBtn.addEventListener('click', () => {
            optionsContainer.querySelectorAll('input').forEach(cb => cb.checked = false);
            mainSelectAll.checked = false;
            onSelectionChange();
        });

        clearBtn.addEventListener('click', () => {
            optionsContainer.querySelectorAll('input').forEach(cb => cb.checked = true);
            mainSelectAll.checked = true;
            onSelectionChange();
        });

        searchInput.addEventListener('input', (e) => {
            const term = e.target.value.toLowerCase();
            optionsContainer.querySelectorAll('.filter-option').forEach(opt => {
                const val = opt.dataset.value.toLowerCase();
                opt.style.display = val.includes(term) ? '' : 'none';
            });
        });
    }

    function populateFilters(data) {
        ['pagina', 'texto', 'cor'].forEach((key, i) => {
            const unique = [...new Set(data.map(item => String(item[key] || "")))].sort();
            createFilterDropdown(filterContainers[i], unique, i);
        });
    }

    function refreshUserFilter(colIdx) {
        const fieldMap = { 3: 'entidade', 4: 'operacao', 5: 'ativo' };
        const field = fieldMap[colIdx];
        const vals = [...new Set(Object.values(userFields).map(f => f[field] || ""))].sort();
        createFilterDropdown(document.querySelector(`.filter-container[data-col="${colIdx}"]`), vals, colIdx);
    }

    function applyFilters() {
        const trs = tableBody.querySelectorAll('tr');
        const container = document.querySelector('.table-container');
        const currentScroll = container.scrollTop;
        const columns = ['pagina', 'texto', 'cor', 'entidade', 'operacao', 'ativo'];

        trs.forEach(tr => {
            const idx = tr.dataset.index;
            if (idx === undefined) return;
            const item = extractedDataCache[idx];
            if (!item) return;

            const uf = userFields[idx] || {};
            let isVisible = !deletedRows.has(String(idx));
            
            if (isVisible) {
                for (let i = 0; i < 6; i++) {
                    const sel = columnFilterSelections[i];
                    if (sel.size > 0) {
                        let val;
                        if (i < 3) {
                            val = item[columns[i]];
                        } else {
                            val = uf[columns[i]] || "";
                        }
                        if (i === 2) val = val === '#0' ? '#000000' : val;
                        
                        if (!sel.has(String(val || ""))) {
                            isVisible = false;
                            break;
                        }
                    }
                }
            }
            tr.style.display = isVisible ? '' : 'none';
        });

        // Update trigger text / indicator states
        filterContainers.forEach((cont, i) => {
            const sel = columnFilterSelections[i];
            const span = cont.querySelector('.filter-trigger span');
            const ind = cont.querySelector('.filter-active-indicator');
            if (span) span.textContent = sel.size > 0 ? sel.size + ' sel' : 'Todos';
            if (ind) {
                if (sel.size > 0) ind.classList.add('active');
                else ind.classList.remove('active');
            }
        });

        requestAnimationFrame(() => {
            if (container) container.scrollTop = currentScroll;
        });
    }

    exportCsvBtn.addEventListener('click', () => {
        let csv = "Página\tTexto\tCor\tEntidade\tOperação\tAtivo\n";
        extractedDataCache.forEach((item, index) => {
            if (deletedRows.has(String(index))) return;
            const uf = userFields[index] || {};
            csv += `${item.pagina}\t"${(item.texto||'').replace(/\n/g,' ')}"\t"${item.cor}"\t"${uf.entidade||''}"\t"${uf.operacao||''}"\t"${uf.ativo||''}"\n`;
        });
        const blob = new Blob(["\uFEFF"+csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url; link.download = "extracao.csv"; link.click();
    });

    // Excel-style Cell Selection Logic
    let isSelecting = false, selectionStart = null, selectionEnd = null;

    tableBody.addEventListener('mousedown', (e) => {
        if (e.button !== 0) return;
        const td = e.target.closest('td');
        if (!td || e.target.tagName === 'INPUT' || e.target.contentEditable === 'true') return;
        isSelecting = true;
        tableBody.classList.add('table-selecting-active');
        const coords = getCellCoords(td);
        if (!e.shiftKey) {
            selectionStart = coords;
            selectionEnd = coords;
            clearSelection();
        } else {
            selectionEnd = coords;
        }
        updateSelectionHighlight();
    });

    tableBody.addEventListener('selectstart', (e) => {
        if (isSelecting) e.preventDefault();
    });

    tableBody.addEventListener('mouseover', (e) => {
        if (!isSelecting) return;
        const td = e.target.closest('td');
        if (!td) return;
        selectionEnd = getCellCoords(td);
        updateSelectionHighlight();
    });

    document.addEventListener('mouseup', () => { 
        isSelecting = false; 
        tableBody.classList.remove('table-selecting-active');
    });

    function getCellCoords(td) {
        const tr = td.parentElement;
        return { 
            row: Array.from(tableBody.children).indexOf(tr), 
            col: Array.from(tr.children).indexOf(td) 
        };
    }

    function clearSelection() {
        tableBody.querySelectorAll('.cell-selected').forEach(c => c.classList.remove('cell-selected'));
    }

    function updateSelectionHighlight() {
        if (!selectionStart || !selectionEnd) return;
        clearSelection();
        const rMin = Math.min(selectionStart.row, selectionEnd.row);
        const rMax = Math.max(selectionStart.row, selectionEnd.row);
        const cMin = Math.min(selectionStart.col, selectionEnd.col);
        const cMax = Math.max(selectionStart.col, selectionEnd.col);
        const rows = tableBody.children;
        for (let r = rMin; r <= rMax; r++) {
            if (rows[r] && rows[r].style.display !== 'none' && !rows[r].classList.contains('row-deleted')) {
                const cells = rows[r].children;
                for (let c = cMin; c <= cMax; c++) {
                    if (cells[c]) cells[c].classList.add('cell-selected');
                }
            }
        }
    }

    // Key Navigation Helpers (Visible rows focus)
    function getVisibleRows() {
        return Array.from(tableBody.children).filter(
            tr => tr.style.display !== 'none' && !tr.classList.contains('row-deleted')
        );
    }

    function getVisibleRowIndex(domRowIdx) {
        const vis = getVisibleRows();
        const tr = tableBody.children[domRowIdx];
        return vis.indexOf(tr);
    }

    function getDomRowIndex(visibleIdx) {
        const vis = getVisibleRows();
        if (vis.length === 0) return -1;
        if (visibleIdx < 0) visibleIdx = 0;
        if (visibleIdx >= vis.length) visibleIdx = vis.length - 1;
        return Array.from(tableBody.children).indexOf(vis[visibleIdx]);
    }

    document.addEventListener('keydown', (e) => {
        // Ctrl+C / Cmd+C: copy selection
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
            const sel = tableBody.querySelectorAll('.cell-selected');
            if (sel.length > 0 && e.target.tagName !== 'INPUT') { e.preventDefault(); copySelectionToClipboard(); }
            return;
        }

        // Delete: remove all rows with selected cells
        if ((e.key === 'Delete' || e.key === 'Del') && selectionStart !== null && e.target.tagName !== 'INPUT') {
            e.preventDefault();
            const toDel = new Set();
            tableBody.querySelectorAll('.cell-selected').forEach(c => toDel.add(c.closest('tr')));
            if (toDel.size > 0) toDel.forEach(tr => deleteRow(tr));
            else {
                // Fallback range if highlight empty
                const rMin = Math.min(selectionStart.row, selectionEnd.row);
                const rMax = Math.max(selectionStart.row, selectionEnd.row);
                const rows = tableBody.children;
                for(let r=rMin; r<=rMax; r++) {
                    const tr = rows[r];
                    if(tr&&tr.style.display!=='none'&&!tr.classList.contains('row-deleted')) deleteRow(tr);
                }
            }
            selectionStart = null; selectionEnd = null; clearSelection();
            return;
        }

        // Shift+Arrow Navigation
        const arrowKeys = ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'];
        if (e.shiftKey && arrowKeys.includes(e.key) && selectionStart !== null) {
            // Avoid conflict if editing text
            if(e.target.contentEditable === 'true') return;
            
            e.preventDefault();
            const vis = getVisibleRows();
            if(vis.length === 0) return;

            const totalCols = tableBody.children[0] ? tableBody.children[0].children.length - 1 : 5;
            const curVisIdx = getVisibleRowIndex(selectionEnd.row);
            
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

            selectionEnd = { row: getDomRowIndex(newVisIdx), col: newCol };
            updateSelectionHighlight();

            // Scroll focus
            const lastTr = tableBody.children[selectionEnd.row];
            if (lastTr) lastTr.scrollIntoView({ block: 'nearest' });
        }
    });

    // Copiar Colunas Selecionadas Acima (SEM cabeçalhos)
    copySelectedBtn.addEventListener('click', () => {
        const selectedColsIndex = [];
        colCheckboxes.forEach((cb, idx) => {
            if (cb.checked) selectedColsIndex.push(idx);
        });
        if (selectedColsIndex.length === 0) return;
        let textToCopy = "";
        const keys = ['pagina', 'texto', 'cor', 'entidade', 'operacao', 'ativo'];
        const trs = tableBody.querySelectorAll('tr');
        trs.forEach(tr => {
            if (tr.style.display !== 'none' && !tr.classList.contains('row-deleted')) {
                const idx = tr.dataset.index;
                if (idx !== undefined) {
                    const item = extractedDataCache[idx], uf = userFields[idx] || {};
                    let rowData = [];
                    selectedColsIndex.forEach(colI => {
                        let val;
                        if (colI < 3) {
                            val = item[keys[colI]] || "";
                            if (colI === 2) val = (val === '#0' || val === '0') ? '#000000' : val;
                        } else { val = uf[keys[colI]] || ""; }
                        rowData.push(String(val).replace(/\n/g, ' '));
                    });
                    textToCopy += rowData.join('\t') + "\n";
                }
            }
        });
        if (textToCopy) {
            navigator.clipboard.writeText(textToCopy).then(() => {
                const orig = copySelectedBtn.innerHTML;
                copySelectedBtn.innerHTML = "Copiado!"; copySelectedBtn.style.color = '#3fb950';
                setTimeout(() => { copySelectedBtn.innerHTML = orig; copySelectedBtn.style.color = ''; }, 1000);
            });
        }
    });

    document.addEventListener('copy', (e) => {
        if (tableBody.querySelectorAll('.cell-selected').length > 0 && document.activeElement.tagName !== 'INPUT') {
            copySelectionToClipboard(); e.preventDefault();
        }
    });

    function copySelectionToClipboard() {
        if (!selectionStart || !selectionEnd) return;
        const rMin = Math.min(selectionStart.row, selectionEnd.row), rMax = Math.max(selectionStart.row, selectionEnd.row);
        const cMin = Math.min(selectionStart.col, selectionEnd.col), cMax = Math.max(selectionStart.col, selectionEnd.col);
        let txt = "";
        for (let r = rMin; r <= rMax; r++) {
            const tr = tableBody.children[r];
            if (tr && tr.style.display !== 'none' && !tr.classList.contains('row-deleted')) {
                const idx = tr.dataset.index;
                if (idx !== undefined) {
                    const item = extractedDataCache[idx], uf = userFields[idx] || {};
                    let row = [];
                    for (let c = cMin; c <= cMax; c++) {
                        let v = [item.pagina, item.texto, item.cor, uf.entidade, uf.operacao, uf.ativo][c];
                        row.push(String(v || "").replace(/\n/g, ' '));
                    }
                    txt += row.join('\t') + "\n";
                }
            }
        }
        navigator.clipboard.writeText(txt);
    }

    function escapeHtml(u) { return String(u || "").replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":"&#039;"}[m])); }

    async function autoExtractLocal(path) {
        spinner.classList.remove('hidden'); resultsSection.classList.add('hidden');
        try {
            const res = await fetch(`/extract-local?path=${encodeURIComponent(path)}`);
            const json = await res.json();
            if (json.error) return alert(json.error);
            extractedDataCache = json.data; renderTable(extractedDataCache);
            resultsSection.classList.remove('hidden');
        } catch (e) { alert("Erro."); } finally { spinner.classList.add('hidden'); }
    }

    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('localFile')) autoExtractLocal(urlParams.get('localFile'));
});
