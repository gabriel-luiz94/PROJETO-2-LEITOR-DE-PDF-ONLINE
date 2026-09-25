document.addEventListener('DOMContentLoaded', () => {
    // ====== AUTO-UPDATER ======
    fetch('/api/update/check')
        .then(r => r.json())
        .then(d => {
            if (d.has_update) {
                const notif = document.getElementById('update-notification');
                if (notif) {
                    document.getElementById('update-version').textContent = d.latest_version;
                    document.getElementById('update-notes').textContent = d.release_notes || '';
                    notif.classList.remove('hidden');
                    
                    document.getElementById('btn-apply-update').onclick = () => {
                        document.getElementById('btn-apply-update').textContent = 'Baixando...';
                        fetch('/api/update/apply', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ download_url: d.download_url })
                        }).then(async r => {
                            const res = await r.json();
                            if (r.ok) {
                                document.getElementById('btn-apply-update').textContent = 'Reiniciando...';
                                setTimeout(() => window.location.reload(), 3000); // Tenta recarregar após um tempo
                            } else {
                                alert('Erro: ' + (res.detail || 'Desconhecido'));
                                document.getElementById('btn-apply-update').textContent = 'Instalar e Reiniciar';
                            }
                        }).catch(e => {
                            alert('Erro de conexão ao atualizar.');
                            document.getElementById('btn-apply-update').textContent = 'Instalar e Reiniciar';
                        });
                    };
                }
            }
        })
        .catch(e => console.log('Auto-update verificação falhou:', e));

    // ====== AUTENTICAÇÃO E SYNC ======
    const token = localStorage.getItem('auth_token');
    if (!token) {
        window.location.href = '/login';
        return;
    } else {
        fetch('/api/health/sync-master', {method: 'POST'})
            .then(r => r.json())
            .then(d => console.log('Sync Mestre:', d.status))
            .catch(e => console.error('Erro no sync background:', e));
    }

    // ====== LOGOUT ======
    const btnLogout = document.getElementById('btn-logout');
    if (btnLogout) {
        btnLogout.addEventListener('click', () => {
            localStorage.removeItem('auth_token');
            localStorage.removeItem('user_id');
            localStorage.removeItem('user_email');
            localStorage.removeItem('is_admin');
            window.location.href = '/login';
        });
    }
    
    const is_admin = localStorage.getItem('is_admin') === 'true';
    if (is_admin) {
        const btnAdmin = document.getElementById('btn-admin');
        if (btnAdmin) btnAdmin.style.display = 'block';
    }

    // ====== SELETOR DE PROJETO (aba Leitor) + REGRAS DO LEITOR (TASK-006) ======
    window.__regrasLeitorProcessamento = [];
    window.__regrasLeitorClassificacao = [];
    window.__regrasLeitorProjetoCarregado = null; // projeto_codigo cujas regras estão carregadas
    window.__listaProjetosCache = [];

    async function carregarProjetosLeitor() {
        try {
            const res = await fetch('/api/projetos');
            const data = await res.json();
            window.__listaProjetosCache = data.projetos || [];
            const sel = document.getElementById('select-projeto-leitor');
            if (!sel) return;
            sel.innerHTML = '';
            if (data.projetos && data.projetos.length > 0) {
                data.projetos.forEach(p => {
                    const opt = document.createElement('option');
                    opt.value = p.nome;
                    opt.dataset.codigo = p.codigo;
                    opt.textContent = `${p.nome} (${p.codigo})`;
                    sel.appendChild(opt);
                });
                const saved = localStorage.getItem('projeto_selecionado');
                if (saved) sel.value = saved;
                // Garante que projeto_selecionado_codigo reflita a opção efetivamente selecionada
                const selected = sel.options[sel.selectedIndex];
                const codigo = selected ? (selected.dataset.codigo || '') : '';
                if (selected) localStorage.setItem('projeto_selecionado_codigo', codigo);
                await carregarRegrasLeitor(codigo);
            }
        } catch (e) {
            console.error('Erro ao carregar projetos:', e);
        }
    }
    carregarProjetosLeitor();

    // Busca as regras de Processamento + Classificação do projeto informado e guarda em
    // window.__regrasLeitor*. Se o projeto não tiver nenhuma regra, oferece usar as regras de
    // outro projeto SÓ para esta sessão de leitura (não associa permanentemente) — TASK-006.
    let _carregandoRegrasLeitor = false;
    async function carregarRegrasLeitor(projetoCodigo) {
        if (!projetoCodigo || _carregandoRegrasLeitor) return;
        _carregandoRegrasLeitor = true;
        try {
            const [resProc, resCls] = await Promise.all([
                fetch(`/api/regras-leitor/processamento?projeto_codigo=${encodeURIComponent(projetoCodigo)}`),
                fetch(`/api/regras-leitor/classificacao?projeto_codigo=${encodeURIComponent(projetoCodigo)}`)
            ]);
            const dataProc = resProc.ok ? await resProc.json() : { regras: [] };
            const dataCls = resCls.ok ? await resCls.json() : { regras: [] };

            if ((!dataProc.regras || dataProc.regras.length === 0) && (!dataCls.regras || dataCls.regras.length === 0)) {
                window.__regrasLeitorProcessamento = [];
                window.__regrasLeitorClassificacao = [];
                window.__regrasLeitorProjetoCarregado = projetoCodigo;
                await ofereceRegrasDeOutroProjeto(projetoCodigo);
                return;
            }

            window.__regrasLeitorProcessamento = dataProc.regras || [];
            window.__regrasLeitorClassificacao = dataCls.regras || [];
            window.__regrasLeitorProjetoCarregado = projetoCodigo;
        } catch (e) {
            console.error('Erro ao carregar regras do leitor:', e);
        } finally {
            _carregandoRegrasLeitor = false;
        }
    }

    async function ofereceRegrasDeOutroProjeto(projetoSemRegras) {
        const outros = (window.__listaProjetosCache || []).filter(p => p.codigo !== projetoSemRegras);
        if (outros.length === 0) return;
        const usar = confirm(
            `O projeto selecionado não tem regras do leitor cadastradas.\n\n` +
            `Deseja usar as regras de outro projeto só para esta leitura? (elas não ficam associadas a este projeto)`
        );
        if (!usar) return;

        const nomes = outros.map((p, i) => `${i + 1}. ${p.nome} (${p.codigo})`).join('\n');
        const escolha = prompt(`Escolha o número do projeto de origem das regras:\n\n${nomes}`);
        const idx = parseInt(escolha, 10) - 1;
        if (isNaN(idx) || idx < 0 || idx >= outros.length) return;

        const origem = outros[idx];
        try {
            const [resProc, resCls] = await Promise.all([
                fetch(`/api/regras-leitor/processamento?projeto_codigo=${encodeURIComponent(origem.codigo)}`),
                fetch(`/api/regras-leitor/classificacao?projeto_codigo=${encodeURIComponent(origem.codigo)}`)
            ]);
            const dataProc = resProc.ok ? await resProc.json() : { regras: [] };
            const dataCls = resCls.ok ? await resCls.json() : { regras: [] };
            window.__regrasLeitorProcessamento = dataProc.regras || [];
            window.__regrasLeitorClassificacao = dataCls.regras || [];
            alert(`Usando as regras do projeto ${origem.nome} (${origem.codigo}) só para esta leitura.`);
        } catch (e) {
            console.error('Erro ao carregar regras do projeto de origem:', e);
        }
    }

    // Reage à troca de projeto no seletor da aba Leitor (o select já grava o localStorage sozinho
    // via onchange inline no HTML — aqui só recarregamos as regras do novo projeto).
    document.addEventListener('change', (e) => {
        if (e.target && e.target.id === 'select-projeto-leitor') {
            const opt = e.target.options[e.target.selectedIndex];
            const codigo = opt ? (opt.dataset.codigo || '') : '';
            carregarRegrasLeitor(codigo);
        }
    });

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
            
            // Força a visibilidade dos botões de ação
            const procBtn = document.getElementById('processarResumo');
            if (procBtn) procBtn.style.display = 'flex';
            
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

    // isGray/isBlue/processAtivoFormula foram substituidas pelo motor de regras do leitor
    // (TASK-006) -- ver static/regras_leitor_engine.js e .ai/tasks/TASK-006-25-09-2026.md.
    function updateRowLogic(index, tr) {
        const item = extractedDataCache[index];
        const engineItem = {
            texto: item.texto,
            cor: item.cor || "#000000",
            layer: item.layer || "",
            pagina: item.pagina,
            index: index,
            allItems: extractedDataCache
        };
        const resultado = RegrasLeitorEngine.processarEClassificar(
            engineItem,
            window.__regrasLeitorProcessamento || [],
            window.__regrasLeitorClassificacao || []
        );
        userFields[index] = { entidade: resultado.entidade, operacao: resultado.operacao, ativo: resultado.ativo };
        const entInput = tr.querySelector('[data-field="entidade"]');
        const opInput = tr.querySelector('[data-field="operacao"]');
        const atInput = tr.querySelector('[data-field="ativo"]');
        if (entInput) entInput.value = resultado.entidade;
        if (opInput) opInput.value = resultado.operacao;
        if (atInput) atInput.value = resultado.ativo;
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
        
        // Garante que o botão apareça após extrair
        const procBtn = document.getElementById('processarResumo');
        if (procBtn) procBtn.style.display = 'flex';
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
        const currentScroll = container ? container.scrollTop : 0;
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

        const csvBtn = document.getElementById('downloadCsv');
        const procBtn = document.getElementById('processarResumo');
        if (csvBtn) csvBtn.style.display = extractedDataCache.length ? 'flex' : 'none';
        if (procBtn) procBtn.style.display = extractedDataCache.length ? 'flex' : 'none';

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

        if (container) {
            requestAnimationFrame(() => {
                container.scrollTop = currentScroll;
            });
        }
    }

    const processarBtn = document.getElementById('processarResumo');
    if (processarBtn) {
        processarBtn.addEventListener('click', () => {
            const rows = tableBody.querySelectorAll('tr');
            const exportData = [];

            rows.forEach(tr => {
                const idx = tr.dataset.index;
                if (idx === undefined || tr.style.display === 'none' || tr.classList.contains('row-deleted')) return;

                const entInput = tr.querySelector('input[data-field="entidade"]');
                const opInput = tr.querySelector('input[data-field="operacao"]');
                const atInput = tr.querySelector('input[data-field="ativo"]');
                const item = extractedDataCache[idx];

                exportData.push({
                    pagina: item.pagina,
                    texto: tr.querySelector('.editable-text-field')?.innerText || item.texto,
                    cor: item.cor,
                    layer: item.layer || "",
                    entidade: entInput ? entInput.value : "0",
                    operacao: opInput ? opInput.value : "",
                    ativo: atInput ? atInput.value : ""
                });
            });

            localStorage.setItem('processar_dados', JSON.stringify(exportData));
            window.open('/resumo', '_blank');
        });
    }

    // Lógica de Colar (CTRL+V)
    document.addEventListener('paste', (e) => {
        const selected = tableBody.querySelectorAll('.cell-selected');
        if (selected.length === 0) return;

        const pasteData = e.clipboardData.getData('text/plain');
        if (!pasteData) return;

        const lines = pasteData.split(/\r?\n/).filter(l => l.trim() !== "");
        if (lines.length === 0) return;

        // Se for uma única linha colada, aplica em todas as células selecionadas
        if (lines.length === 1) {
            const val = lines[0].trim();
            selected.forEach(cell => {
                const input = cell.querySelector('input');
                if (input) {
                    input.value = val;
                    // Trigger input event to update filters/cache
                    input.dispatchEvent(new Event('input'));
                } else if (cell.querySelector('.editable-text-field')) {
                    cell.querySelector('.editable-text-field').innerText = val;
                }
            });
        } 
        // Se for múltiplas linhas, cola sequencialmente (opcional, por agora fazemos o simples)
        e.preventDefault();
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
        if (selectedColsIndex.length === 0) {
            alert("Selecione ao menos uma coluna para copiar.");
            return;
        }
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
            copyToClipboard(textToCopy, () => {
                const orig = copySelectedBtn.innerHTML;
                copySelectedBtn.innerHTML = "Copiado!"; copySelectedBtn.style.color = '#3fb950';
                setTimeout(() => { copySelectedBtn.innerHTML = orig; copySelectedBtn.style.color = ''; }, 1000);
            });
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
        copyToClipboard(txt);
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

/* ═══════════════════════════════════════
   REGRAS DO LEITOR (TASK-006)
═══════════════════════════════════════ */
// escapeHtml já existe dentro do DOMContentLoaded (usado pela extração), mas essas funções
// rodam em escopo top-level (como abrirModalRecs) — cópia local para não depender do fechamento.
function _escapeHtmlRegrasLeitor(u) {
    return String(u || '').replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[m]));
}

function _descreverRegraProcessamento(r) {
    const partes = [];
    if (r.cor_em && r.cor_em.length) partes.push(`cor: ${r.cor_em.join('/')}`);
    if (r.layer_em && r.layer_em.length) partes.push(`layer: ${r.layer_em.join('/')}`);
    if (r.texto_regex) partes.push(`texto ~ /${r.texto_regex}/`);
    if (r.ativo_regex) partes.push(`ativo ~ /${r.ativo_regex}/`);
    if (r.vizinhanca) partes.push(`vizinhança (±${r.vizinhanca.janela || 10} linhas)`);
    const condicao = partes.length ? partes.join(', ') : 'qualquer';
    const acoes = [];
    if (r.operacao) acoes.push(`operação=${r.operacao}`);
    if (r.ativo_template !== undefined && r.ativo_template !== null) acoes.push(`ativo="${r.ativo_template}"`);
    return `<div style="padding:6px 8px; border-bottom:1px solid #21262d;">
        <span style="color:#8b949e;">[fase ${r.fase || 3} · ordem ${r.ordem}] ${r.modo || 'DEFINIR'}</span> —
        <b>se</b> ${_escapeHtmlRegrasLeitor(condicao)} <b>então</b> ${_escapeHtmlRegrasLeitor(acoes.join(', ') || '(sem ação)')}
        ${r.parar ? '<span style="color:#f0883e;">· para</span>' : ''}
    </div>`;
}

function _descreverRegraClassificacao(r) {
    const partes = [];
    if (r.operacao_em && r.operacao_em.length) partes.push(`operação: ${r.operacao_em.join('/')}`);
    if (r.cor_em && r.cor_em.length) partes.push(`cor: ${r.cor_em.join('/')}`);
    if (r.cor_nao_em && r.cor_nao_em.length) partes.push(`cor ≠ ${r.cor_nao_em.join('/')}`);
    if (r.layer_em && r.layer_em.length) partes.push(`layer: ${r.layer_em.join('/')}`);
    if (r.ativo_regex) partes.push(`ativo ~ /${r.ativo_regex}/`);
    if (r.texto_regex) partes.push(`texto ~ /${r.texto_regex}/`);
    const condicao = partes.length ? partes.join(', ') : 'qualquer';
    const acoes = [];
    if (r.entidade) acoes.push(`entidade=${r.entidade}`);
    if (r.operacao_ajustada) acoes.push(`operação=${r.operacao_ajustada}`);
    return `<div style="padding:6px 8px; border-bottom:1px solid #21262d;">
        <span style="color:#8b949e;">[ordem ${r.ordem}]</span> —
        <b>se</b> ${_escapeHtmlRegrasLeitor(condicao)} <b>então</b> ${_escapeHtmlRegrasLeitor(acoes.join(', ') || '(bloqueia sem classificar)')}
        ${r.parar ? '<span style="color:#f0883e;">· para</span>' : ''}
    </div>`;
}

function abrirModalRegrasLeitor() {
    const modal = document.getElementById('modal-regras-leitor');
    if (!modal) return;
    modal.style.display = 'flex';

    const sel = document.getElementById('select-projeto-leitor');
    const opt = sel ? sel.options[sel.selectedIndex] : null;
    const nomeProjeto = opt ? opt.value : '(nenhum projeto selecionado)';
    const infoEl = document.getElementById('regras-leitor-projeto-atual');
    if (infoEl) {
        infoEl.textContent = `Projeto selecionado: ${nomeProjeto}` +
            (window.__regrasLeitorProjetoCarregado ? ` — regras carregadas de: ${window.__regrasLeitorProjetoCarregado}` : '');
    }

    const proc = window.__regrasLeitorProcessamento || [];
    const cls = window.__regrasLeitorClassificacao || [];

    const procEl = document.getElementById('regras-leitor-processamento-lista');
    const clsEl = document.getElementById('regras-leitor-classificacao-lista');
    if (procEl) {
        procEl.innerHTML = proc.length
            ? proc.slice().sort((a, b) => (a.fase || 3) - (b.fase || 3) || (a.ordem || 0) - (b.ordem || 0)).map(_descreverRegraProcessamento).join('')
            : '<div style="color:#8b949e; padding:8px;">Nenhuma regra de processamento carregada para este projeto.</div>';
    }
    if (clsEl) {
        clsEl.innerHTML = cls.length
            ? cls.slice().sort((a, b) => (a.ordem || 0) - (b.ordem || 0)).map(_descreverRegraClassificacao).join('')
            : '<div style="color:#8b949e; padding:8px;">Nenhuma regra de classificação carregada para este projeto.</div>';
    }
}

/* ═══════════════════════════════════════
   HISTÓRICO DE RECS (Página Inicial)
═══════════════════════════════════════ */
async function abrirModalRecs() {
    const modal = document.getElementById('modal-recs');
    if (!modal) return;
    modal.style.display = 'flex';
    const tbody = document.getElementById('tbody-recs-salvos-index');
    if (!tbody) return;
    tbody.innerHTML = '<tr><td colspan="3" style="text-align:center; padding: 20px; color:#8b949e;">Carregando RECs...</td></tr>';
    
    try {
        const projCodeLeitor = localStorage.getItem('projeto_selecionado_codigo') || '229';
        const res = await fetch(`/api/recs?projeto=${encodeURIComponent(projCodeLeitor)}`);
        if (!res.ok) throw new Error('Erro ao listar RECs');
        const recs = await res.json();
        tbody.innerHTML = '';
        
        if (!recs || recs.length === 0) {
            tbody.innerHTML = '<tr><td colspan="3" style="text-align:center; padding: 20px; color:#8b949e;">Nenhum REC salvo no banco ainda.</td></tr>';
            return;
        }

        recs.forEach(r => {
            const tr = document.createElement('tr');
            tr.style.borderBottom = '1px solid rgba(255,255,255,0.08)';
            tr.innerHTML = `
                <td style="padding: 10px; font-weight: 600; color: #fff;">${r.numero_obra}</td>
                <td style="padding: 10px; color: #8b949e;">${r.data_criacao || '-'}</td>
                <td style="padding: 10px; text-align: right; white-space: nowrap;">
                    <button onclick="abrirRecNoOrcamento('${encodeURIComponent(r.numero_obra)}')" style="background:#238636; color:white; border:none; padding:4px 9px; border-radius:4px; cursor:pointer; font-size:0.8rem; margin-right:4px;" title="Abrir esta obra na tela de Orçamento">Abrir Orçamento</button>
                    <button onclick="baixarRecDireto('${encodeURIComponent(r.numero_obra)}')" style="background:#1f6feb; color:white; border:none; padding:4px 9px; border-radius:4px; cursor:pointer; font-size:0.8rem; margin-right:4px;" title="Exportar arquivo .rec">Baixar .rec</button>
                    <button onclick="excluirRecIndex('${encodeURIComponent(r.numero_obra)}')" style="background:#da3633; color:white; border:none; padding:4px 8px; border-radius:4px; cursor:pointer; font-size:0.8rem;" title="Excluir do banco">✖</button>
                </td>
            `;
            tbody.appendChild(tr);
        });
    } catch(e) {
        console.error("Erro ao carregar RECs:", e);
        tbody.innerHTML = '<tr><td colspan="3" style="text-align:center; padding: 20px; color:#f85149;">Erro ao carregar histórico de RECs.</td></tr>';
    }
}

function filtrarModalRecsIndex() {
    const input = document.getElementById("search-recs-index");
    if (!input) return;
    const filter = input.value.toUpperCase();
    const tbody = document.getElementById("tbody-recs-salvos-index");
    if (!tbody) return;
    const trs = tbody.getElementsByTagName("tr");
    for (let i = 0; i < trs.length; i++) {
        const td = trs[i].getElementsByTagName("td")[0];
        if (td) {
            const txt = td.textContent || td.innerText;
            trs[i].style.display = txt.toUpperCase().indexOf(filter) > -1 ? "" : "none";
        }
    }
}

function abrirRecNoOrcamento(numObra) {
    const decoded = decodeURIComponent(numObra);
    window.location.href = `/resultado_orcamento?rec=${encodeURIComponent(decoded)}`;
}

async function baixarRecDireto(numObra) {
    const decoded = decodeURIComponent(numObra);
    try {
        const res = await fetch(`/api/recs/${encodeURIComponent(decoded)}`);
        if (!res.ok) throw new Error('Obra não encontrada');
        const data = await res.json();
        const conteudo = data.dados_json;
        if (!conteudo) { alert("Esta obra não possui conteúdo REC."); return; }

        const filename = `${decoded}_ODI.rec`;
        const blob = new Blob([conteudo], { type: 'text/plain;charset=utf-8' });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    } catch(e) {
        alert("Erro ao baixar arquivo .rec: " + e.message);
    }
}

async function excluirRecIndex(numObra) {
    const decoded = decodeURIComponent(numObra);
    if (!confirm(`Deseja realmente excluir a obra ${decoded} do banco?`)) return;
    try {
        const res = await fetch(`/api/recs/${encodeURIComponent(decoded)}`, { method: 'DELETE' });
        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Erro ao excluir');
        }
        abrirModalRecs(); // Recarrega lista
    } catch(e) {
        alert("Erro ao excluir: " + e.message);
    }
}
