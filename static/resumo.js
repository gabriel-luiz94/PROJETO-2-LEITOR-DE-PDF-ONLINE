document.addEventListener('DOMContentLoaded', () => {
    const dataRaw = localStorage.getItem('processar_dados');
    if (!dataRaw) {
        alert("Nenhum dado encontrado para processar.");
        return;
    }

    const allData = JSON.parse(dataRaw);
    const cabosData = allData.filter(d => d.entidade === "CABO");
    const outrosData = allData.filter(d => d.entidade !== "CABO" && d.entidade !== "0");

    renderTable('body-cabos', cabosData);
    renderTable('body-outros', outrosData);

    document.getElementById('count-cabos').textContent = cabosData.length;
    document.getElementById('count-outros').textContent = outrosData.length;

    function renderTable(bodyId, data) {
        const body = document.getElementById(bodyId);
        body.innerHTML = '';
        if (data.length === 0) {
            body.innerHTML = '<tr><td colspan="6" style="text-align:center; opacity: 0.5;">Nenhum item encontrado.</td></tr>';
            return;
        }

        data.forEach((item, index) => {
            const tr = document.createElement('tr');
            const displayColor = item.cor || "#000000";
            tr.innerHTML = `
                <td>${item.pagina || 1}</td>
                <td style="word-break: break-all; max-width: 400px;">
                    <div class="text-cell-container">
                        <span class="selectable-text editable-text-field">${escapeHtml(item.texto)}</span>
                    </div>
                </td>
                <td><div class="color-badge"><div class="color-swatch" style="background-color: ${displayColor};"></div><span>${displayColor.toUpperCase()}</span></div></td>
                <td class="col-user"><input class="user-input" type="text" value="${item.entidade}"></td>
                <td class="col-user"><input class="user-input" type="text" value="${item.operacao}"></td>
                <td class="col-user"><input class="user-input" type="text" value="${item.ativo}"></td>
            `;
            body.appendChild(tr);

            // Funções de Edição
            const textField = tr.querySelector('.editable-text-field');
            textField.addEventListener('blur', function() { this.contentEditable = false; });
            textField.addEventListener('dblclick', function() { this.contentEditable = true; this.focus(); });
        });
    }

    function escapeHtml(u) { return String(u || "").replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":"&#039;"}[m])); }

    // Implementação de Seleção Estilo Excel (Simplificada para o resumo)
    let isSelecting = false, selectionStart = null, selectionEnd = null, activeBody = null;

    document.querySelectorAll('tbody').forEach(body => {
        body.addEventListener('mousedown', (e) => {
            if (e.button !== 0) return;
            const td = e.target.closest('td');
            if (!td || e.target.tagName === 'INPUT') return;
            isSelecting = true;
            activeBody = body;
            clearAllSelections();
            selectionStart = getCoords(td, body);
            selectionEnd = selectionStart;
            updateHighlight(body);
        });

        body.addEventListener('mouseover', (e) => {
            if (!isSelecting || activeBody !== body) return;
            const td = e.target.closest('td');
            if (!td) return;
            selectionEnd = getCoords(td, body);
            updateHighlight(body);
        });
    });

    document.addEventListener('mouseup', () => { isSelecting = false; });

    function getCoords(td, body) {
        const tr = td.parentElement;
        return { row: Array.from(body.children).indexOf(tr), col: Array.from(tr.children).indexOf(td) };
    }

    function clearAllSelections() {
        document.querySelectorAll('.cell-selected').forEach(c => c.classList.remove('cell-selected'));
    }

    function updateHighlight(body) {
        if (!selectionStart || !selectionEnd) return;
        document.querySelectorAll('.cell-selected').forEach(c => c.classList.remove('cell-selected'));
        const rMin = Math.min(selectionStart.row, selectionEnd.row), rMax = Math.max(selectionStart.row, selectionEnd.row);
        const cMin = Math.min(selectionStart.col, selectionEnd.col), cMax = Math.max(selectionStart.col, selectionEnd.col);
        const rows = body.children;
        for (let r = rMin; r <= rMax; r++) {
            if (rows[r]) {
                const cells = rows[r].children;
                for (let c = cMin; c <= cMax; c++) {
                    if (cells[c]) cells[c].classList.add('cell-selected');
                }
            }
        }
    }

    // Atalho de Cópia (CTRL+C)
    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.key === 'c' && !window.getSelection().toString()) {
            const selected = document.querySelectorAll('.cell-selected');
            if (selected.length === 0) return;
            
            let text = "";
            let lastRow = -1;
            selected.forEach(cell => {
                const row = Array.from(cell.parentElement.parentElement.children).indexOf(cell.parentElement);
                if (lastRow !== -1 && row !== lastRow) text += "\n";
                text += (cell.querySelector('input')?.value || cell.innerText) + "\t";
                lastRow = row;
            });
            navigator.clipboard.writeText(text.trim());
        }
    });
});
