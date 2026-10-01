/* ═══════════════════════════════════════════════════════════════════════════
   Prompts de validação (TASK-012/019) — aba do painel de regras do Resumo; só admin (o backend valida).
   Conteúdo sempre inserido via textContent/value (nunca innerHTML).
═══════════════════════════════════════════════════════════════════════════ */
const PV = { prompts: [] };

function pvEl(id) { return document.getElementById(id); }

function pvProjeto() { return pvEl('rdProjeto').value || 'DEFAULT'; }  // seletor de projeto do painel, compartilhado

function pvErros(lista) {
    const box = pvEl('pvErros');
    box.replaceChildren();
    if (!lista || !lista.length) { box.classList.add('hidden'); return; }
    lista.forEach(e => { const p = document.createElement('div'); p.textContent = '• ' + e; box.appendChild(p); });
    box.classList.remove('hidden');
}

async function pvLerErro(res) {
    let corpo = {};
    try { corpo = await res.json(); } catch (e) { /* sem corpo */ }
    const d = corpo.detail;
    if (d && d.erros) { pvErros(d.erros); return 'Prompt inválido — veja os erros acima.'; }
    return (typeof d === 'string' && d) || 'Erro ao processar a requisição.';
}

async function iniciarPromptsValidacao() {
    // O painel (painel_regras.js) monta o DOM e preenche o seletor de projeto compartilhado (#rdProjeto).
    pvEl('pvPrompt').addEventListener('change', pvMostrar);
    pvEl('pvSalvar').addEventListener('click', pvSalvar);
    pvEl('pvSemente').addEventListener('click', pvRestaurarSemente);
    pvEl('pvHistorico').addEventListener('click', pvAlternarHistorico);
}

async function pvCarregarLista(manter) {
    const selPrompt = pvEl('pvPrompt');
    const atual = manter === true ? selPrompt.value : '';
    const res = await fetch(`/api/validacao/prompts?projeto_codigo=${encodeURIComponent(pvProjeto())}`);
    if (!res.ok) { rpMensagem('error', 'Não foi possível carregar os prompts de validação.'); return; }
    PV.prompts = (await res.json()).prompts;
    selPrompt.replaceChildren();
    PV.prompts.forEach(p => selPrompt.appendChild(new Option(`${p.prompt_id} (${p.meta.modo || '?'})`, p.prompt_id)));
    if (atual) selPrompt.value = atual;
    pvEl('pvHistLista').classList.add('hidden');
    pvMostrar();
}

function pvMostrar() {
    const p = PV.prompts.find(x => x.prompt_id === pvEl('pvPrompt').value);
    pvErros(null);
    pvEl('pvTexto').value = p ? p.conteudo : '';
    pvEl('pvStatus').textContent = !p ? '' : (p.personalizado
        ? 'Versão própria deste projeto'
        : (pvProjeto() === 'DEFAULT' ? 'Padrão do sistema' : 'Herdando o padrão (salvar cria a versão do projeto)'));
}

async function pvSalvar() {
    pvErros(null);
    const res = await fetch(`/api/validacao/prompts/${encodeURIComponent(pvEl('pvPrompt').value)}`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ projeto_codigo: pvProjeto(), conteudo: pvEl('pvTexto').value })
    });
    if (!res.ok) { rpMensagem('error', await pvLerErro(res)); return; }
    rpMensagem('success', 'Prompt salvo.');
    await pvCarregarLista(true);
}

async function pvRestaurarSemente() {
    if (!confirm('Substituir o texto pelo da semente do sistema? A versão atual fica no histórico.')) return;
    pvErros(null);
    const res = await fetch(`/api/validacao/prompts/${encodeURIComponent(pvEl('pvPrompt').value)}/restaurar-semente`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ projeto_codigo: pvProjeto() })
    });
    if (!res.ok) { rpMensagem('error', await pvLerErro(res)); return; }
    rpMensagem('success', 'Semente restaurada.');
    await pvCarregarLista(true);
}

async function pvAlternarHistorico() {
    const box = pvEl('pvHistLista');
    if (!box.classList.contains('hidden')) { box.classList.add('hidden'); return; }
    const res = await fetch(`/api/validacao/prompts/${encodeURIComponent(pvEl('pvPrompt').value)}/historico?projeto_codigo=${encodeURIComponent(pvProjeto())}`);
    if (!res.ok) { rpMensagem('error', await pvLerErro(res)); return; }
    const historico = (await res.json()).historico;
    box.replaceChildren();
    if (!historico.length) {
        const vazio = document.createElement('div');
        vazio.className = 'p-3 text-[#8b949e]';
        vazio.textContent = 'Sem versões anteriores para este projeto.';
        box.appendChild(vazio);
    }
    historico.forEach(h => {
        const linha = document.createElement('div');
        linha.className = 'flex items-center justify-between gap-3 p-3';
        const info = document.createElement('span');
        info.textContent = `${h.criado_em} — ${h.criado_por || 'desconhecido'}`;
        const acoes = document.createElement('span');
        acoes.className = 'space-x-3';
        const ver = document.createElement('button');
        ver.className = 'text-blue-400 hover:text-blue-300';
        ver.textContent = 'Ver no editor';
        ver.onclick = () => { pvEl('pvTexto').value = h.conteudo; pvErros(null); };
        const rev = document.createElement('button');
        rev.className = 'text-[#d97706] hover:text-[#b45309]';
        rev.textContent = 'Reverter';
        rev.onclick = () => pvReverter(h.id);
        acoes.append(ver, rev);
        linha.append(info, acoes);
        box.appendChild(linha);
    });
    box.classList.remove('hidden');
}

async function pvReverter(historicoId) {
    if (!confirm('Reverter para esta versão? A versão atual fica no histórico.')) return;
    const res = await fetch(`/api/validacao/prompts/${encodeURIComponent(pvEl('pvPrompt').value)}/reverter`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ projeto_codigo: pvProjeto(), historico_id: historicoId })
    });
    if (!res.ok) { rpMensagem('error', await pvLerErro(res)); return; }
    rpMensagem('success', 'Versão revertida.');
    await pvCarregarLista(true);
}
