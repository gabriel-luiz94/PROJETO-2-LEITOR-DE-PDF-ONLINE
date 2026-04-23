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

    // Inicialização
    initTable('cabos');
    initTable('outros');

    function initTable(type) {
        renderTable(type);
        updateCounters();
        setupFilters(type);
    }

    function renderTable(type) {
        const state = tableStates[type];
        const body = document.getElementById(state.bodyId);
        body.innerHTML = '';

        state.data.forEach((item, index) => {
            const tr = document.createElement('tr');
            tr.dataset.index = index;
            tr.dataset.type = type;
            const displayColor = item.cor || "#000000";
            tr.innerHTML = `
                <td>${item.pagina || 1}</td>
                <td style="word-break: break-all; max-width: 400px;">
                    <div class="text-cell-container">
                        <span class="selectable-text editable-text-field">${escapeHtml(item.texto)}</span>
                    </div>
                </td>
                <td><div class="color-badge"><div class="color-swatch" style="background-color: ${displayColor};"></div><span>${displayColor.toUpperCase()}</span></div></td>
                <td class="col-user"><input class="user-input" type="text" value="${item.entidade}" data-field="entidade"></td>
                <td class="col-user"><input class="user-input" type="text" value="${item.operacao}" data-field="operacao"></td>
                <td class="col-user"><input class="user-input" type="text" value="${item.ativo}" data-field="ativo"></td>
            `;
            body.appendChild(tr);

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

    // --- FILTROS ---
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
                    <label class="filter-option select-all-option">
                        <input type="checkbox" checked>
                        <span>(Selecionar Tudo)</span>
                    </label>
                    <div class="options-container">
                        ${values.map(val => `<label class="filter-option" data-value="${val}"><input type="checkbox" checked><span>${val || "(Vazio)"}</span></label>`).join('')}
                    </div>
                </div>
            </div>
        `;

        const trigger = container.querySelector('.filter-trigger');
        const dropdown = container.querySelector('.filter-dropdown');
        const mainSelectAll = container.querySelector('.select-all-option input');
        const optionsContainer = container.querySelector('.options-container');
        
        trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            document.querySelectorAll('.filter-dropdown.active').forEach(d => d !== dropdown && d.classList.remove('active'));
            dropdown.classList.toggle('active');
        });

        dropdown.addEventListener('click', e => e.stopPropagation());

        const onSelectionChange = () => {
            const checks = optionsContainer.querySelectorAll('input');
            const checked = optionsContainer.querySelectorAll('input:checked');
            mainSelectAll.checked = checked.length === checks.length;
            mainSelectAll.indeterminate = checked.length > 0 && checked.length < checks.length;

            selections.clear();
            if (checked.length < checks.length) {
                checked.forEach(c => selections.add(c.parentElement.dataset.value));
            }
            applyFilters(type);
            updateFilterTrigger(container, selections);
        };

        optionsContainer.querySelectorAll('input').forEach(cb => {
            cb.addEventListener('change', onSelectionChange);
        });

        mainSelectAll.addEventListener('change', () => {
            optionsContainer.querySelectorAll('input').forEach(cb => {
                cb.checked = mainSelectAll.checked;
            });
            onSelectionChange();
        });
    }

    function updateFilterTrigger(container, selections) {
        const span = container.querySelector('.filter-trigger span');
        const ind = container.querySelector('.filter-active-indicator');
        span.textContent = selections.size > 0 ? selections.size + ' sel' : 'Todos';
        ind.classList.toggle('active', selections.size > 0);
    }

    function applyFilters(type) {
        const state = tableStates[type];
        const rows = document.getElementById(state.bodyId).children;
        const columns = ['pagina', 'texto', 'cor', 'entidade', 'operacao', 'ativo'];

        for (let tr of rows) {
            const item = state.data[tr.dataset.index];
            let isVisible = true;

            for (let i = 0; i < 6; i++) {
                const sel = state.filters[i];
                if (sel.size > 0) {
                    const val = String(item[columns[i]] || "");
                    if (!sel.has(val)) { isVisible = false; break; }
                }
            }
            tr.style.display = isVisible ? '' : 'none';
        }
    }

    // --- CONTEXT MENU ---
    const contextMenu = document.getElementById('context-menu');
    let contextTargetRow = null;

    function showContextMenu(e, tr) {
        contextTargetRow = tr;
        const x = Math.min(e.clientX, window.innerWidth - 180);
        const y = Math.min(e.clientY, window.innerHeight - 150);
        contextMenu.style.left = x + 'px';
        contextMenu.style.top = y + 'px';
        contextMenu.classList.add('active');
    }

    document.addEventListener('click', () => contextMenu.classList.remove('active'));

    document.getElementById('ctx-delete').addEventListener('click', () => {
        if (!contextTargetRow) return;
        const type = contextTargetRow.dataset.type;
        const index = parseInt(contextTargetRow.dataset.index);
        tableStates[type].data.splice(index, 1);
        renderTable(type);
        updateCounters();
    });

    document.getElementById('ctx-edit').addEventListener('click', () => {
        if (!contextTargetRow) return;
        const span = contextTargetRow.querySelector('.editable-text-field');
        span.contentEditable = true;
        span.focus();
    });

    // --- ADICIONAR LINHA ---
    document.querySelectorAll('.btn-add').forEach(btn => {
        btn.addEventListener('click', () => {
            const type = btn.dataset.target.includes('cabos') ? 'cabos' : 'outros';
            tableStates[type].data.push({ pagina: 1, texto: "NOVA LINHA", cor: "#000000", entidade: type === 'cabos' ? "CABO" : "", operacao: "M", ativo: "" });
            renderTable(type);
            updateCounters();
        });
    });

    // --- SELEÇÃO ESTILO EXCEL & CTRL+C / CTRL+V ---
    let isSelecting = false, selectionStart = null, selectionEnd = null, activeType = null;

    document.querySelectorAll('tbody').forEach(body => {
        const type = body.id.includes('cabos') ? 'cabos' : 'outros';
        body.addEventListener('mousedown', (e) => {
            if (e.button !== 0 || e.target.tagName === 'INPUT') return;
            isSelecting = true;
            activeType = type;
            selectionStart = getCoords(e.target.closest('td'), body);
            selectionEnd = selectionStart;
            clearSelection();
            updateHighlight();
        });

        body.addEventListener('mouseover', (e) => {
            if (!isSelecting || activeType !== type) return;
            const td = e.target.closest('td');
            if (td) {
                selectionEnd = getCoords(td, body);
                updateHighlight();
            }
        });
    });

    document.addEventListener('mouseup', () => isSelecting = false);

    function getCoords(td, body) {
        const tr = td.parentElement;
        return { row: Array.from(body.children).indexOf(tr), col: Array.from(tr.children).indexOf(td) };
    }

    function clearSelection() {
        document.querySelectorAll('.cell-selected').forEach(c => c.classList.remove('cell-selected'));
    }

    function updateHighlight() {
        if (!selectionStart || !selectionEnd) return;
        clearSelection();
        const body = document.getElementById(tableStates[activeType].bodyId);
        const rMin = Math.min(selectionStart.row, selectionEnd.row), rMax = Math.max(selectionStart.row, selectionEnd.row);
        const cMin = Math.min(selectionStart.col, selectionEnd.col), cMax = Math.max(selectionStart.col, selectionEnd.col);
        
        for (let r = rMin; r <= rMax; r++) {
            const tr = body.children[r];
            if (tr && tr.style.display !== 'none') {
                for (let c = cMin; c <= cMax; c++) {
                    if (tr.children[c]) tr.children[c].classList.add('cell-selected');
                }
            }
        }
    }

    // CTRL+C / CTRL+V
    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.key === 'c' && !window.getSelection().toString()) {
            const selected = document.querySelectorAll('.cell-selected');
            if (selected.length === 0) return;
            let text = "", lastRow = -1;
            selected.forEach(cell => {
                const row = Array.from(cell.parentElement.parentElement.children).indexOf(cell.parentElement);
                if (lastRow !== -1 && row !== lastRow) text += "\n";
                text += (cell.querySelector('input')?.value || cell.innerText) + "\t";
                lastRow = row;
            });
            navigator.clipboard.writeText(text.trim());
        }

        if (e.ctrlKey && e.key === 'v') {
            const selected = document.querySelectorAll('.cell-selected');
            if (selected.length === 0) return;
            navigator.clipboard.readText().then(content => {
                if (!content) return;
                const lines = content.split(/\r?\n/).filter(l => l.trim() !== "");
                if (lines.length === 1) {
                    const val = lines[0].trim();
                    selected.forEach(cell => {
                        const input = cell.querySelector('input');
                        if (input) {
                            input.value = val;
                            input.dispatchEvent(new Event('input'));
                        } else if (cell.querySelector('.editable-text-field')) {
                            cell.querySelector('.editable-text-field').innerText = val;
                        }
                    });
                }
            });
        }
    });

    // Copiar Botão (Seção)
    document.querySelectorAll('.btn-copy-report').forEach(btn => {
        btn.addEventListener('click', () => {
            const type = btn.dataset.target.includes('cabos') ? 'cabos' : 'outros';
            const table = document.getElementById(`table-${type}`);
            const body = document.getElementById(tableStates[type].bodyId);
            
            // Verifica colunas selecionadas
            const selectedCols = Array.from(table.querySelectorAll('.col-checkbox:checked')).map(cb => parseInt(cb.dataset.col));
            if (selectedCols.length === 0) {
                alert("Selecione pelo menos uma coluna para copiar.");
                return;
            }

            const headers = ["Pág", "Texto", "Cor", "Entidade", "Operação", "Ativo"];
            let text = selectedCols.map(i => headers[i]).join('\t') + "\n";

            Array.from(body.children).forEach(tr => {
                if (tr.style.display !== 'none') {
                    const cells = Array.from(tr.children);
                    const rowData = selectedCols.map(colI => {
                        if (colI < 3) return cells[colI].innerText.replace(/\n/g, ' ');
                        return cells[colI].querySelector('input').value.replace(/\n/g, ' ');
                    });
                    text += rowData.join('\t') + "\n";
                }
            });

            navigator.clipboard.writeText(text);
            const orig = btn.innerHTML;
            btn.innerHTML = "Copiado!";
            setTimeout(() => btn.innerHTML = orig, 1000);
        });
    });

    function escapeHtml(u) { return String(u || "").replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":"&#039;"}[m])); }
});
