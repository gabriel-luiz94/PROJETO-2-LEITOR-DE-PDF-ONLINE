document.addEventListener('DOMContentLoaded', () => {
    const token = localStorage.getItem('auth_token');
    if (!token) {
        window.location.href = '/login';
        return;
    }

    const is_admin = localStorage.getItem('is_admin') === 'true';
    if (!is_admin) {
        alert("Acesso negado. Apenas administradores podem ver o painel.");
        window.location.href = '/';
        return;
    }

    loadUsers();
    iniciarPromptsValidacao();
    iniciarRegrasDominio();

    document.getElementById('addUserForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const email = document.getElementById('newEmail').value;
        const password = document.getElementById('newPassword').value;
        const role = document.getElementById('newRole').value;
        
        try {
            const res = await fetch(`/api/admin/users?token=${token}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email, password, role: role })
            });
            const data = await res.json();
            
            if (res.ok) {
                showMessage('success', data.msg);
                document.getElementById('addUserForm').reset();
                loadUsers();
            } else {
                showMessage('error', data.detail || 'Erro ao criar usuário');
            }
        } catch (err) {
            showMessage('error', 'Erro de conexão com o servidor.');
        }
    });
});

async function loadUsers() {
    const token = localStorage.getItem('auth_token');
    const tbody = document.getElementById('userTableBody');
    tbody.innerHTML = '<tr><td colspan="3" class="px-6 py-4 text-center">Carregando...</td></tr>';
    
    try {
        const res = await fetch(`/api/admin/users?token=${token}`);
        if (!res.ok) {
            if (res.status === 403) {
                window.location.href = '/';
                return;
            }
            throw new Error('Falha ao carregar');
        }
        const users = await res.json();
        
        tbody.innerHTML = '';
        users.forEach(user => {
            const tr = document.createElement('tr');
            
            const tdEmail = document.createElement('td');
            tdEmail.className = 'px-6 py-4';
            tdEmail.textContent = user.email;
            
            const tdType = document.createElement('td');
            tdType.className = 'px-6 py-4';
            const badge = document.createElement('span');
            badge.className = user.role === 'admin' ? 'px-2 py-1 bg-[#d97706] text-black text-xs font-bold rounded' : 'px-2 py-1 bg-[#30363d] text-gray-300 text-xs rounded';
            badge.textContent = user.role === 'admin' ? 'Admin' : 'Operador';
            tdType.appendChild(badge);
            
            const tdAction = document.createElement('td');
            tdAction.className = 'px-6 py-4 text-right space-x-2';
            
            const btnToggle = document.createElement('button');
            btnToggle.className = 'text-sm text-blue-400 hover:text-blue-300 transition-colors';
            btnToggle.textContent = user.role === 'admin' ? 'Rebaixar p/ Operador' : 'Promover p/ Admin';
            btnToggle.onclick = () => updateRole(user.id, user.role === 'admin' ? 'operador' : 'admin');
            
            const btnReset = document.createElement('button');
            btnReset.className = 'text-sm text-yellow-500 hover:text-yellow-400 transition-colors ml-4';
            btnReset.textContent = 'Mudar Senha';
            btnReset.onclick = () => resetPassword(user.id, user.email);
            
            const btnDelete = document.createElement('button');
            btnDelete.className = 'text-sm text-red-400 hover:text-red-300 transition-colors ml-4';
            btnDelete.textContent = 'Excluir';
            btnDelete.onclick = () => deleteUser(user.id, user.email);
            
            tdAction.appendChild(btnToggle);
            tdAction.appendChild(btnReset);
            tdAction.appendChild(btnDelete);
            
            tr.appendChild(tdEmail);
            tr.appendChild(tdType);
            tr.appendChild(tdAction);
            tbody.appendChild(tr);
        });
        
        if (users.length === 0) {
            tbody.innerHTML = '<tr><td colspan="3" class="px-6 py-4 text-center">Nenhum usuário local encontrado.</td></tr>';
        }
    } catch (err) {
        tbody.innerHTML = '<tr><td colspan="3" class="px-6 py-4 text-center text-red-500">Erro ao carregar usuários locais.</td></tr>';
    }
}

async function updateRole(userId, newRole) {
    const token = localStorage.getItem('auth_token');
    try {
        const res = await fetch(`/api/admin/users/${userId}/role?token=${token}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ role: newRole })
        });
        if (res.ok) {
            loadUsers();
        } else {
            const data = await res.json();
            showMessage('error', data.detail || 'Erro ao alterar privilégios');
        }
    } catch (err) {
        showMessage('error', 'Erro de conexão.');
    }
}

async function resetPassword(userId, email) {
    const newPassword = prompt(`Digite a nova senha para ${email}: (Mín. 6 caracteres)`);
    if (!newPassword) return; // Cancelou ou deixou em branco
    
    if (newPassword.length < 6) {
        showMessage('error', 'A senha precisa ter no mínimo 6 caracteres.');
        return;
    }
    
    const token = localStorage.getItem('auth_token');
    try {
        const res = await fetch(`/api/admin/users/${userId}/password?token=${token}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ password: newPassword })
        });
        
        if (res.ok) {
            showMessage('success', `Senha de ${email} redefinida com sucesso!`);
        } else {
            const data = await res.json();
            showMessage('error', data.detail || 'Erro ao redefinir senha');
        }
    } catch (err) {
        showMessage('error', 'Erro de conexão.');
    }
}

async function deleteUser(userId, email) {
    if (!confirm(`Tem certeza que deseja excluir o usuário ${email}?`)) return;
    
    const token = localStorage.getItem('auth_token');
    try {
        const res = await fetch(`/api/admin/users/${userId}?token=${token}`, {
            method: 'DELETE'
        });
        if (res.ok) {
            loadUsers();
        } else {
            const data = await res.json();
            showMessage('error', data.detail || 'Erro ao excluir usuário');
        }
    } catch (err) {
        showMessage('error', 'Erro de conexão.');
    }
}

function showMessage(type, msg) {
    const errorMsg = document.getElementById('errorMsg');
    const successMsg = document.getElementById('successMsg');
    
    errorMsg.classList.add('hidden');
    successMsg.classList.add('hidden');
    
    if (type === 'error') {
        errorMsg.textContent = msg;
        errorMsg.classList.remove('hidden');
    } else {
        successMsg.textContent = msg;
        successMsg.classList.remove('hidden');
    }
    
    setTimeout(() => {
        errorMsg.classList.add('hidden');
        successMsg.classList.add('hidden');
    }, 5000);
}


/* ═══════════════════════════════════════════════════════════════════════════
   Prompts de validação (TASK-012) — GET aberto a autenticados; escrita só admin (o backend valida).
   Conteúdo sempre inserido via textContent/value (nunca innerHTML).
═══════════════════════════════════════════════════════════════════════════ */
const PV = { prompts: [] };

function pvEl(id) { return document.getElementById(id); }

function pvProjeto() { return pvEl('pvProjeto').value || 'DEFAULT'; }

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
    const selProj = pvEl('pvProjeto');
    const opDefault = new Option('DEFAULT (padrão para todos)', 'DEFAULT');
    selProj.appendChild(opDefault);
    try {
        const res = await fetch('/api/projetos');
        if (res.ok) (await res.json()).forEach(p => selProj.appendChild(new Option(`${p.nome} (${p.codigo})`, p.codigo)));
    } catch (e) { /* segue só com DEFAULT */ }

    selProj.addEventListener('change', pvCarregarLista);
    pvEl('pvPrompt').addEventListener('change', pvMostrar);
    pvEl('pvSalvar').addEventListener('click', pvSalvar);
    pvEl('pvSemente').addEventListener('click', pvRestaurarSemente);
    pvEl('pvHistorico').addEventListener('click', pvAlternarHistorico);
    pvCarregarLista();
}

async function pvCarregarLista(manter) {
    const selPrompt = pvEl('pvPrompt');
    const atual = manter === true ? selPrompt.value : '';
    const res = await fetch(`/api/validacao/prompts?projeto_codigo=${encodeURIComponent(pvProjeto())}`);
    if (!res.ok) { showMessage('error', 'Não foi possível carregar os prompts de validação.'); return; }
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
    if (!res.ok) { showMessage('error', await pvLerErro(res)); return; }
    showMessage('success', 'Prompt salvo.');
    await pvCarregarLista(true);
}

async function pvRestaurarSemente() {
    if (!confirm('Substituir o texto pelo da semente do sistema? A versão atual fica no histórico.')) return;
    pvErros(null);
    const res = await fetch(`/api/validacao/prompts/${encodeURIComponent(pvEl('pvPrompt').value)}/restaurar-semente`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ projeto_codigo: pvProjeto() })
    });
    if (!res.ok) { showMessage('error', await pvLerErro(res)); return; }
    showMessage('success', 'Semente restaurada.');
    await pvCarregarLista(true);
}

async function pvAlternarHistorico() {
    const box = pvEl('pvHistLista');
    if (!box.classList.contains('hidden')) { box.classList.add('hidden'); return; }
    const res = await fetch(`/api/validacao/prompts/${encodeURIComponent(pvEl('pvPrompt').value)}/historico?projeto_codigo=${encodeURIComponent(pvProjeto())}`);
    if (!res.ok) { showMessage('error', await pvLerErro(res)); return; }
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
    if (!res.ok) { showMessage('error', await pvLerErro(res)); return; }
    showMessage('success', 'Versão revertida.');
    await pvCarregarLista(true);
}


/* ═══════════════════════════════════════════════════════════════════════════
   Regras de domínio da validação (TASK-013) — GET aberto a autenticados; escrita/teste só admin
   (o backend valida schema e permissão). Conteúdo sempre via textContent/value.
═══════════════════════════════════════════════════════════════════════════ */
const RD = { regras: [] };

const RD_MODELOS = {
    requer: { se_regex: '', exige_regex: '', qtd_min: 1 },
    proibe: { se_regex: '', com_regex: '' },
    nao_isolado: { se_regex: '', acompanhado_por_regex: '', min_outros: 1 },
    texto: { texto_regex: '' },
    minimo_total: { ativo_regex: '^P50$', contribuicoes: [{ se_regex: '', metros_por_unidade: 1 }] }
};

function rdEl(id) { return document.getElementById(id); }
function rdProjeto() { return rdEl('rdProjeto').value || 'DEFAULT'; }

function rdErros(lista) {
    const box = rdEl('rdErros');
    box.replaceChildren();
    if (!lista || !lista.length) { box.classList.add('hidden'); return; }
    lista.forEach(e => { const p = document.createElement('div'); p.textContent = '• ' + e; box.appendChild(p); });
    box.classList.remove('hidden');
}

async function rdLerErro(res) {
    let corpo = {};
    try { corpo = await res.json(); } catch (e) { /* sem corpo */ }
    const d = corpo.detail;
    if (d && d.erros) { rdErros(d.erros); return 'Regras inválidas — veja os erros acima.'; }
    return (typeof d === 'string' && d) || 'Erro ao processar a requisição.';
}

const RD_CLS_INPUT = 'px-2 py-1 border border-[#30363d] bg-[#0d1117] text-white rounded-md text-sm';

async function iniciarRegrasDominio() {
    const sel = rdEl('rdProjeto');
    sel.appendChild(new Option('DEFAULT (padrão para todos)', 'DEFAULT'));
    try {
        const res = await fetch('/api/projetos');
        if (res.ok) (await res.json()).forEach(p => sel.appendChild(new Option(`${p.nome} (${p.codigo})`, p.codigo)));
    } catch (e) { /* segue só com DEFAULT */ }
    sel.addEventListener('change', rdCarregar);
    rdEl('rdAdicionar').addEventListener('click', rdAdicionar);
    rdEl('rdSalvar').addEventListener('click', rdSalvar);
    rdEl('rdSemente').addEventListener('click', rdRestaurarSemente);
    rdEl('rdNovas').addEventListener('click', rdAdicionarNovas);
    rdEl('rdHistorico').addEventListener('click', rdAlternarHistorico);
    rdEl('rdTestar').addEventListener('click', rdTestar);
    rdCarregar();
}

async function rdCarregar() {
    const res = await fetch(`/api/validacao/regras?projeto_codigo=${encodeURIComponent(rdProjeto())}`);
    if (!res.ok) { showMessage('error', 'Não foi possível carregar as regras de domínio.'); return; }
    const dados = await res.json();
    RD.regras = dados.regras;
    rdEl('rdStatus').textContent = dados.personalizado
        ? 'Versão própria deste projeto'
        : (rdProjeto() === 'DEFAULT' ? 'Padrão do sistema' : 'Herdando o padrão (salvar cria a versão do projeto)');
    rdEl('rdHistLista').classList.add('hidden');
    rdErros(null);
    rdRenderizar();
}

function rdCampo(rotulo, elemento) {
    const wrap = document.createElement('label');
    wrap.className = 'block text-xs text-[#8b949e]';
    wrap.append(rotulo, elemento);
    elemento.classList.add('block', 'mt-1');
    return wrap;
}

function rdRenderizar() {
    const lista = rdEl('rdLista');
    lista.replaceChildren();
    RD.regras.forEach((regra, i) => lista.appendChild(rdCartao(regra, i)));
}

function rdCartao(regra, i) {
    const card = document.createElement('div');
    card.className = 'border border-[#30363d] rounded-md p-3';

    const topo = document.createElement('div');
    topo.className = 'flex flex-wrap items-center gap-3';
    const ativa = document.createElement('input');
    ativa.type = 'checkbox'; ativa.checked = !!regra.ativa; ativa.title = 'Regra ligada';
    ativa.onchange = () => { regra.ativa = ativa.checked; };
    const rid = document.createElement('input');
    rid.className = RD_CLS_INPUT + ' w-48 font-mono'; rid.value = regra.id || '';
    rid.oninput = () => { regra.id = rid.value.trim(); };
    const tipo = document.createElement('select');
    tipo.className = RD_CLS_INPUT;
    Object.keys(RD_MODELOS).forEach(t => tipo.appendChild(new Option(t, t)));
    tipo.value = regra.tipo;
    tipo.onchange = () => {
        regra.tipo = tipo.value;
        regra.parametros = JSON.parse(JSON.stringify(RD_MODELOS[tipo.value]));
        regra.escopo = tipo.value === 'minimo_total' ? 'cabos+outros' : 'outros';
        rdRenderizar();
    };
    const sev = document.createElement('select');
    sev.className = RD_CLS_INPUT;
    ['erro', 'aviso', 'info'].forEach(s => sev.appendChild(new Option(s, s)));
    sev.value = regra.severidade;
    sev.onchange = () => { regra.severidade = sev.value; };
    const excluir = document.createElement('button');
    excluir.className = 'ml-auto text-red-400 hover:text-red-300 text-sm';
    excluir.textContent = 'Excluir';
    excluir.onclick = () => { if (confirm(`Excluir a regra ${regra.id}?`)) { RD.regras.splice(i, 1); rdRenderizar(); } };
    topo.append(ativa, rid, tipo, sev, excluir);

    const msg = document.createElement('input');
    msg.className = RD_CLS_INPUT + ' w-full mt-2'; msg.value = regra.mensagem || ''; msg.placeholder = 'Mensagem mostrada ao usuário';
    msg.oninput = () => { regra.mensagem = msg.value; };

    const linha2 = document.createElement('div');
    linha2.className = 'grid grid-cols-1 md:grid-cols-3 gap-3 mt-2';
    const ops = document.createElement('input');
    ops.className = RD_CLS_INPUT + ' w-full'; ops.value = (regra.operacoes || ['I', '*I']).join(', ');
    ops.oninput = () => { regra.operacoes = ops.value.split(',').map(s => s.trim()).filter(Boolean); };
    const params = document.createElement('textarea');
    params.rows = 5; params.spellcheck = false;
    params.className = RD_CLS_INPUT + ' w-full font-mono text-xs md:col-span-2';
    params.value = JSON.stringify(regra.parametros, null, 2);
    params.oninput = () => {
        try { regra.parametros = JSON.parse(params.value); params.classList.remove('border-red-500'); regra._jsonInvalido = false; }
        catch (e) { params.classList.add('border-red-500'); regra._jsonInvalido = true; }
    };
    linha2.append(rdCampo('Operações em que vale', ops), rdCampo('Parâmetros (JSON)', params));

    const desc = document.createElement('div');
    desc.className = 'text-xs text-[#8b949e] mt-2';
    desc.textContent = regra.descricao || '';
    card.append(topo, msg, linha2, desc);
    return card;
}

function rdAdicionar() {
    let n = RD.regras.length + 1;
    while (RD.regras.some(r => r.id === `C2-NOVA-${n}`)) n++;
    RD.regras.push({
        id: `C2-NOVA-${n}`, escopo: 'outros', tipo: 'requer', severidade: 'aviso', mensagem: '',
        ativa: false, operacoes: ['I', '*I'], parametros: JSON.parse(JSON.stringify(RD_MODELOS.requer)), descricao: ''
    });
    rdRenderizar();
}

function rdRegrasParaEnvio() {
    if (RD.regras.some(r => r._jsonInvalido)) { rdErros(['Há parâmetros com JSON inválido (campo em vermelho).']); return null; }
    return RD.regras.map(r => { const c = { ...r }; delete c._jsonInvalido; return c; });
}

async function rdSalvar() {
    rdErros(null);
    const regras = rdRegrasParaEnvio();
    if (!regras) return;
    const res = await fetch('/api/validacao/regras', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ projeto_codigo: rdProjeto(), regras })
    });
    if (!res.ok) { showMessage('error', await rdLerErro(res)); return; }
    showMessage('success', 'Regras salvas.');
    await rdCarregar();
}

async function rdRestaurarSemente() {
    if (!confirm('Substituir as regras pelas da semente do sistema (todas desligadas)? A versão atual fica no histórico.')) return;
    rdErros(null);
    const res = await fetch('/api/validacao/regras/restaurar-semente', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ projeto_codigo: rdProjeto() })
    });
    if (!res.ok) { showMessage('error', await rdLerErro(res)); return; }
    showMessage('success', 'Semente restaurada.');
    await rdCarregar();
}

async function rdAdicionarNovas() {
    rdErros(null);
    const res = await fetch('/api/validacao/regras/adicionar-novas', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ projeto_codigo: rdProjeto() })
    });
    if (!res.ok) { showMessage('error', await rdLerErro(res)); return; }
    const { adicionadas } = await res.json();
    showMessage('success', adicionadas.length
        ? `Adicionadas (desligadas): ${adicionadas.join(', ')}.`
        : 'Não há regras novas na semente: a lista já tem todas.');
    await rdCarregar();
}

async function rdAlternarHistorico() {
    const box = rdEl('rdHistLista');
    if (!box.classList.contains('hidden')) { box.classList.add('hidden'); return; }
    const res = await fetch(`/api/validacao/regras/historico?projeto_codigo=${encodeURIComponent(rdProjeto())}`);
    if (!res.ok) { showMessage('error', await rdLerErro(res)); return; }
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
        info.textContent = `${h.criado_em} — ${h.criado_por || 'desconhecido'} — ${h.regras.length} regras, ${h.regras.filter(r => r.ativa).length} ligadas`;
        const acoes = document.createElement('span');
        acoes.className = 'space-x-3';
        const ver = document.createElement('button');
        ver.className = 'text-blue-400 hover:text-blue-300'; ver.textContent = 'Ver no editor';
        ver.onclick = () => { RD.regras = JSON.parse(JSON.stringify(h.regras)); rdErros(null); rdRenderizar(); };
        const rev = document.createElement('button');
        rev.className = 'text-[#d97706] hover:text-[#b45309]'; rev.textContent = 'Reverter';
        rev.onclick = () => rdReverter(h.id);
        acoes.append(ver, rev);
        linha.append(info, acoes);
        box.appendChild(linha);
    });
    box.classList.remove('hidden');
}

async function rdReverter(historicoId) {
    if (!confirm('Reverter para esta versão? A versão atual fica no histórico.')) return;
    const res = await fetch('/api/validacao/regras/reverter', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ projeto_codigo: rdProjeto(), historico_id: historicoId })
    });
    if (!res.ok) { showMessage('error', await rdLerErro(res)); return; }
    showMessage('success', 'Versão revertida.');
    await rdCarregar();
}

function rdLinhasDeTeste(textoId) {
    return rdEl(textoId).value.split('\n').map(l => l.trim()).filter(Boolean).map(l => {
        const i = l.search(/\s/);
        return i < 0 ? { operacao: '', ativo: l } : { operacao: l.slice(0, i), ativo: l.slice(i + 1).trim() };
    });
}

async function rdTestar() {
    rdErros(null);
    const regras = rdRegrasParaEnvio();
    if (!regras) return;
    const res = await fetch('/api/validacao/regras/testar', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ regras, cabos: rdLinhasDeTeste('rdTesteCabos'), outros: rdLinhasDeTeste('rdTesteOutros') })
    });
    const alvo = rdEl('rdTesteResultado');
    alvo.replaceChildren();
    if (!res.ok) { showMessage('error', await rdLerErro(res)); return; }
    const { achados, resumo } = await res.json();
    const titulo = document.createElement('div');
    titulo.className = 'mb-1 font-medium';
    titulo.textContent = achados.length
        ? `${achados.length} achado(s): ${resumo.erro} erro, ${resumo.aviso} aviso, ${resumo.info} info`
        : 'Nenhum achado para estas linhas com as regras ligadas.';
    alvo.appendChild(titulo);
    achados.forEach(a => {
        const item = document.createElement('div');
        item.textContent = `[${a.severidade}] ${a.linha_id} · ${a.regra_id} — ${a.mensagem}`;
        alvo.appendChild(item);
    });
}
