
/* ═══════════════════════════════════════
   GERADOR DE CÓDIGOS - CABOS
═══════════════════════════════════════ */

let linhasModalCabos = [];
let debounceTimeoutCabos = null;

function abrirModalCabos() {
    linhasModalCabos = [{ op: 'I', ativo: '', fase: 'ABC', comp: 80, qtd: 1, desc: '' }];
    renderizarTabelaModalCabos();
    document.getElementById('modal-cabos-gerador').classList.remove('hidden');
}

function adicionarLinhaCabos() {
    linhasModalCabos.push({ op: 'I', ativo: '', fase: 'ABC', comp: 80, qtd: 1, desc: '' });
    renderizarTabelaModalCabos();
}

function removerLinhaCabos(index) {
    linhasModalCabos.splice(index, 1);
    renderizarTabelaModalCabos();
}

async function buscarDescricaoAtivo(index) {
    const linha = linhasModalCabos[index];
    const ativoStr = (linha.ativo || '').trim();
    if (!ativoStr) {
        linha.desc = '';
        renderizarTabelaModalCabos();
        return;
    }

    // Auto-detect length based on input
    const upperStr = ativoStr.toUpperCase();
    if (upperStr.includes('MULT') || upperStr.includes('MTX')) {
        linha.comp = 40;
    } else {
        linha.comp = 80;
    }

    try {
        const res = await fetch('/api/orcamento/search?q=' + encodeURIComponent(ativoStr));
        const data = await res.json();
        if (data.resultados && data.resultados.length > 0) {
            linha.desc = data.resultados[0].desc_ativo || data.resultados[0].desc_codigo || 'Ativo encontrado';
        } else {
            linha.desc = 'Ativo não encontrado';
        }
    } catch (e) {
        linha.desc = 'Erro na busca';
    }
    renderizarTabelaModalCabos();
}

function atualizarLinhaCabos(index, field, value) {
    linhasModalCabos[index][field] = value;
    
    if (field === 'ativo') {
        clearTimeout(debounceTimeoutCabos);
        debounceTimeoutCabos = setTimeout(() => {
            buscarDescricaoAtivo(index);
        }, 500);
    }
}

function renderizarTabelaModalCabos() {
    const tbody = document.getElementById('body-modal-cabos');
    tbody.innerHTML = '';

    linhasModalCabos.forEach((linha, index) => {
        const tr = document.createElement('tr');
        
        // OP
        const tdOp = document.createElement('td');
        tdOp.innerHTML = `<select class="row-select" onchange="atualizarLinhaCabos(${index}, 'op', this.value)">
            ${OPERACOES.map(op => `<option value="${op}" ${linha.op === op ? 'selected' : ''}>${op}</option>`).join('')}
        </select>`;
        
        // ATIVO
        const tdAtivo = document.createElement('td');
        tdAtivo.innerHTML = `<input type="text" class="row-input" value="${linha.ativo}" oninput="atualizarLinhaCabos(${index}, 'ativo', this.value)" placeholder="Digite o ativo...">`;
        
        // FASE
        const tdFase = document.createElement('td');
        const fases = ['ABC', 'A', 'B', 'C', 'AC'];
        tdFase.innerHTML = `<select class="row-select" onchange="atualizarLinhaCabos(${index}, 'fase', this.value)">
            ${fases.map(f => `<option value="${f}" ${linha.fase === f ? 'selected' : ''}>${f}</option>`).join('')}
        </select>`;
        
        // COMP
        const tdComp = document.createElement('td');
        tdComp.innerHTML = `<input type="number" class="row-input" value="${linha.comp}" oninput="atualizarLinhaCabos(${index}, 'comp', parseFloat(this.value) || 0)">`;
        
        // QTD
        const tdQtd = document.createElement('td');
        tdQtd.innerHTML = `<input type="number" class="row-input" value="${linha.qtd}" min="1" oninput="atualizarLinhaCabos(${index}, 'qtd', parseInt(this.value) || 1)">`;
        
        // DESC
        const tdDesc = document.createElement('td');
        tdDesc.style.fontSize = '0.75rem';
        tdDesc.style.color = '#8b949e';
        tdDesc.textContent = linha.desc;
        
        // DELETE
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

function inserirGeradorNaTabelaCabos() {
    if (!linhasModalCabos.length) return;
    
    pushHistory(); // Salva estado antes de inserir
    
    linhasModalCabos.forEach(linha => {
        const ativoBase = (linha.ativo || '').trim();
        if (!ativoBase) return;
        
        let faseFormatada = linha.fase;
        if (faseFormatada === 'ABC') faseFormatada = '3';
        else if (faseFormatada === 'AC') faseFormatada = '2';
        else faseFormatada = '1';
        
        // Formato final: "ATIVO FASE COMP" -> Ex: "MULTIPLEX 70 3 40"
        const ativoFinal = `${ativoBase} ${faseFormatada} ${linha.comp}`;
        
        for (let i = 0; i < linha.qtd; i++) {
            tableStates.cabos.data.push({
                entidade: 'CABO',
                operacao: linha.op,
                ativo: ativoFinal
            });
        }
    });
    
    renderTable('cabos');
    document.getElementById('modal-cabos-gerador').classList.add('hidden');
    buildAtivoSets();
    buildDataLists();
}
