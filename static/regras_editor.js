/* ═══════════════════════════════════════════════════════════════════════════
   Editor visual das regras de domínio da validação (TASK-013/017/018).
   - Cada regra é mostrada como frase em português (vinda do backend: `frase`, /descrever); nada de regra de negócio
     é reavaliado aqui. Conteúdo sempre via textContent/value (nunca innerHTML).
   - Camadas por projeto: padrão + ajustes do projeto; selos de origem, botões −/+ (ocultar/reexibir) e "voltar ao padrão".
   Depende de showMessage() (admin.js). Escrita/teste só admin (o backend valida schema e permissão).
═══════════════════════════════════════════════════════════════════════════ */
const RD = { regras: [], grupos: {}, grupos_origem: {}, grupos_usados: {}, padrao: null, ativos: [], timer: null };
const RD_META = ['origem', 'oculta', 'frase'];
const RD_CMPS = ['=', '!=', '>=', '<=', '>', '<', 'entre'];
const RD_CLS_INPUT = 'px-2 py-1 border border-[#30363d] bg-[#0d1117] text-white rounded-md text-sm';
const RD_CLS_BTN = 'px-3 py-1 rounded-md border border-[#30363d] text-[#c9d1d9] hover:bg-[#21262d] text-sm';
const RD_SELOS = {
    padrao: ['padrão', 'bg-[#21262d] text-[#8b949e]'],
    sobrescrita: ['ajustada neste projeto', 'bg-amber-900 text-amber-200'],
    projeto: ['só deste projeto', 'bg-blue-900 text-blue-200']
};

function rdEl(id) { return document.getElementById(id); }
function rdProjeto() { return rdEl('rdProjeto').value || 'DEFAULT'; }
function rdEhProjeto() { return rdProjeto() !== 'DEFAULT'; }
function rdClone(x) { return JSON.parse(JSON.stringify(x)); }
function rdNo(tag, cls, texto) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (texto !== undefined) e.textContent = texto;
    return e;
}
function rdBotao(texto, fn, cls, titulo) {
    const b = rdNo('button', cls || RD_CLS_BTN, texto);
    b.type = 'button';
    if (titulo) b.title = titulo;
    b.onclick = fn;
    return b;
}
function rdLimpa(regra) {
    const c = {};
    Object.keys(regra).forEach(k => { if (!RD_META.includes(k) && !k.startsWith('_')) c[k] = regra[k]; });
    return c;
}

function rdErros(lista) {
    const box = rdEl('rdErros');
    box.replaceChildren();
    if (!lista || !lista.length) { box.classList.add('hidden'); return; }
    lista.forEach(e => box.appendChild(rdNo('div', '', '• ' + e)));
    box.classList.remove('hidden');
}
function rdAvisos(lista) {
    const box = rdEl('rdAvisos');
    box.replaceChildren();
    if (!lista || !lista.length) { box.classList.add('hidden'); return; }
    lista.forEach(e => box.appendChild(rdNo('div', '', '• ' + e)));
    box.classList.remove('hidden');
}
async function rdLerErro(res) {
    let corpo = {};
    try { corpo = await res.json(); } catch (e) { /* sem corpo */ }
    const d = corpo.detail;
    if (d && d.erros) { rdErros(d.erros); return 'Regras inválidas — veja os erros acima.'; }
    return (typeof d === 'string' && d) || 'Erro ao processar a requisição.';
}
async function rdPost(url, corpo) {
    return fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(corpo) });
}

/* ── seletores e valores ↔ texto ─────────────────────────────────────────── */
function rdSelParaTexto(sel) {
    if (typeof sel === 'string') return sel;
    if (Array.isArray(sel) && sel.every(s => typeof s === 'string')) return sel.join(', ');
    if (sel && typeof sel === 'object' && !Array.isArray(sel) && typeof sel.regex === 'string' && Object.keys(sel).length === 1) return '/' + sel.regex + '/';
    return JSON.stringify(sel);
}
function rdTextoParaSel(t) {
    t = t.trim();
    if (t.startsWith('{') || t.startsWith('[')) { try { return JSON.parse(t); } catch (e) { return t; } }
    if (t.length > 2 && t.startsWith('/') && t.endsWith('/')) return { regex: t.slice(1, -1) };
    if (t.includes(',')) { const l = t.split(',').map(s => s.trim()).filter(Boolean); return l.length > 1 ? l : (l[0] || ''); }
    return t;
}
function rdValParaTexto(v) { return typeof v === 'number' ? String(v) : JSON.stringify(v); }
function rdTextoParaVal(t) {
    t = t.trim();
    if (/^-?\d+(\.\d+)?$/.test(t)) return Number(t);
    if (t.startsWith('{') || t.startsWith('[')) { try { return JSON.parse(t); } catch (e) { return t; } }
    return t;
}
function rdEntreParaTexto(v) { return Array.isArray(v) && v.every(x => typeof x === 'number') ? v.join(', ') : JSON.stringify(v); }
function rdTextoParaEntre(t) {
    t = t.trim();
    if (t.startsWith('[')) { try { return JSON.parse(t); } catch (e) { return t; } }
    return t.split(',').map(s => rdTextoParaVal(s));
}

/* ── atualização da frase (só o backend descreve a regra) ────────────────── */
function rdAgendarFrases(regra) {
    clearTimeout(RD.timer);
    RD.timer = setTimeout(() => (regra ? [regra] : RD.regras).forEach(rdAtualizarFrase), 350);
}
async function rdAtualizarFrase(regra) {
    if (regra._jsonInvalido || !regra._frase) return;
    try {
        const res = await rdPost('/api/validacao/regras/descrever', { regra: rdLimpa(regra), grupos: RD.grupos });
        if (!res.ok) return;
        const d = await res.json();
        regra._frase.textContent = d.erros.length ? '⚠ ' + d.erros.join(' · ') : d.frase;
        regra._frase.classList.toggle('text-red-400', d.erros.length > 0);
    } catch (e) { /* sem rede: mantém a frase anterior */ }
}

/* ── carga ───────────────────────────────────────────────────────────────── */
async function iniciarRegrasDominio() {
    const sel = rdEl('rdProjeto');
    sel.appendChild(new Option('DEFAULT (padrão para todos)', 'DEFAULT'));
    try {
        const res = await fetch('/api/projetos');
        if (res.ok) { const d = await res.json(); (d.projetos || d).forEach(p => sel.appendChild(new Option(`${p.nome} (${p.codigo})`, p.codigo))); }
    } catch (e) { /* segue só com DEFAULT */ }
    try {
        const res = await fetch('/api/validacao/regras/ativos');
        if (res.ok) RD.ativos = (await res.json()).ativos;
    } catch (e) { /* sem autocomplete de códigos */ }
    sel.addEventListener('change', rdCarregar);
    rdEl('rdAdicionar').addEventListener('click', rdAbrirAssistente);
    rdEl('rdSalvar').addEventListener('click', rdSalvar);
    rdEl('rdSemente').addEventListener('click', rdRestaurarSemente);
    rdEl('rdNovas').addEventListener('click', rdAdicionarNovas);
    rdEl('rdHistorico').addEventListener('click', rdAlternarHistorico);
    rdEl('rdTestar').addEventListener('click', rdTestar);
    rdEl('rdNovoGrupo').addEventListener('click', rdNovoGrupo);
    rdCarregar();
}

async function rdCarregar() {
    const res = await fetch(`/api/validacao/regras?projeto_codigo=${encodeURIComponent(rdProjeto())}`);
    if (!res.ok) { showMessage('error', 'Não foi possível carregar as regras de domínio.'); return; }
    const dados = await res.json();
    RD.regras = dados.regras;
    RD.grupos = dados.grupos;
    RD.grupos_origem = dados.grupos_origem || {};
    RD.grupos_usados = dados.grupos_usados || {};
    RD.padrao = null;
    if (rdEhProjeto()) {  // o padrão é necessário para "voltar ao padrão" e para o selo ao vivo
        try {
            const p = await fetch('/api/validacao/regras?projeto_codigo=DEFAULT');
            if (p.ok) { const d = await p.json(); RD.padrao = { regras: Object.fromEntries(d.regras.map(r => [r.id, r])), grupos: d.grupos }; }
        } catch (e) { /* sem padrão: desativa "voltar ao padrão" */ }
    }
    rdEl('rdStatus').textContent = !rdEhProjeto() ? 'Padrão do sistema (vale para todos os projetos)'
        : (dados.personalizado ? 'Este projeto tem ajustes sobre o padrão' : 'Este projeto usa o padrão sem ajustes');
    rdEl('rdNovas').classList.toggle('hidden', rdEhProjeto());
    rdEl('rdSemente').textContent = rdEhProjeto() ? 'Descartar ajustes do projeto' : 'Restaurar semente';
    rdEl('rdAdicionar').textContent = rdEhProjeto() ? 'Adicionar regra só neste projeto' : 'Adicionar regra';
    rdEl('rdHistLista').classList.add('hidden');
    rdEl('rdAssistente').classList.add('hidden');
    rdErros(null);
    rdAvisos(dados.avisos);
    rdRenderizar();
}

function rdRenderizar() {
    const lista = rdEl('rdLista');
    lista.replaceChildren();
    RD.regras.forEach((regra, i) => lista.appendChild(rdCartao(regra, i)));
    rdRenderizarGrupos();
    rdAtualizarDatalist();
}

function rdAtualizarDatalist() {
    const dl = rdEl('rdSugestoes');
    dl.replaceChildren();
    Object.keys(RD.grupos).forEach(g => dl.appendChild(new Option('@' + g)));
    RD.ativos.forEach(a => dl.appendChild(new Option(a)));
}

/* ── origem (selo) ao vivo ───────────────────────────────────────────────── */
function rdOrigemViva(regra) {
    if (!rdEhProjeto()) return null;
    const base = RD.padrao && RD.padrao.regras[regra.id];
    if (regra.origem === 'projeto') return 'projeto';
    if (!base) return regra.origem || 'padrao';
    return JSON.stringify(rdLimpa(base)) === JSON.stringify(rdLimpa(regra)) ? 'padrao' : 'sobrescrita';
}

/* ── editor de condições ─────────────────────────────────────────────────── */
function rdTipoDe(c) { return ['tem', 'soma', 'texto', 'todos', 'algum', 'nenhum', 'nao'].find(k => k in c) || 'tem'; }
const RD_TIPOS = { tem: 'tem o ativo', soma: 'soma dos ativos', texto: 'texto casa com', todos: 'TODAS (E)', algum: 'ALGUMA (OU)', nenhum: 'NENHUMA', nao: 'NÃO' };

function rdCondicaoPadrao(tipo, anterior) {
    if (tipo === 'tem') return { tem: '' };
    if (tipo === 'soma') return { soma: '', qtd: { '>=': 1 } };
    if (tipo === 'texto') return { texto: '' };
    if (tipo === 'nao') return { nao: anterior || { tem: '' } };
    return { [tipo]: anterior ? [anterior] : [{ tem: '' }] };
}

function rdEditorSel(valor, aoMudar) {
    const inp = rdNo('input', RD_CLS_INPUT + ' w-56 font-mono');
    inp.value = rdSelParaTexto(valor);
    inp.setAttribute('list', 'rdSugestoes');
    inp.placeholder = 'CFU · TR* · @GRUPO · A, B · /regex/';
    inp.title = 'Código do ativo, curinga (*, ?), @GRUPO, lista separada por vírgula, ou /regex/ (avançado)';
    inp.oninput = () => aoMudar(rdTextoParaSel(inp.value));
    return inp;
}

function rdEditorQtd(cond, regra, refazer) {
    const caixa = rdNo('span', 'inline-flex flex-wrap items-center gap-1');
    const qtd = cond.qtd;
    if (!qtd) {
        if (rdTipoDe(cond) === 'tem') caixa.appendChild(rdBotao('qualquer quantidade · restringir', () => { cond.qtd = { '>=': 1 }; refazer(); }, RD_CLS_BTN + ' opacity-80'));
        return caixa;
    }
    caixa.appendChild(rdNo('span', 'text-xs text-[#8b949e]', 'com quantidade'));
    Object.keys(qtd).forEach(op => {
        const linha = rdNo('span', 'inline-flex items-center gap-1');
        const s = rdNo('select', RD_CLS_INPUT);
        RD_CMPS.forEach(o => s.appendChild(new Option(o === 'entre' ? 'entre' : o.replace('>=', '≥').replace('<=', '≤').replace('!=', '≠'), o)));
        s.value = op;
        const v = rdNo('input', RD_CLS_INPUT + ' w-24 font-mono');
        v.value = op === 'entre' ? rdEntreParaTexto(qtd[op]) : rdValParaTexto(qtd[op]);
        v.title = op === 'entre' ? 'dois números: 1, 3' : 'número (ou avançado: {"base":1,"vezes":2,"mais":0,"se":{...}})';
        v.oninput = () => { qtd[s.value] = s.value === 'entre' ? rdTextoParaEntre(v.value) : rdTextoParaVal(v.value); rdAgendarFrases(regra); };
        s.onchange = () => {
            const novo = {};
            Object.keys(qtd).forEach(k => { novo[k === op ? s.value : k] = k === op ? (s.value === 'entre' ? [1, 2] : (op === 'entre' ? 1 : qtd[k])) : qtd[k]; });
            cond.qtd = novo; refazer(); rdAgendarFrases(regra);
        };
        linha.append(s, v);
        if (Object.keys(qtd).length > 1 || rdTipoDe(cond) === 'tem') {
            linha.appendChild(rdBotao('×', () => { delete qtd[op]; if (!Object.keys(qtd).length) delete cond.qtd; refazer(); rdAgendarFrases(regra); }, RD_CLS_BTN, 'remover comparador'));
        }
        caixa.appendChild(linha);
    });
    const livre = RD_CMPS.find(o => !(o in qtd));
    if (livre) caixa.appendChild(rdBotao('+ E', () => { qtd[livre] = livre === 'entre' ? [1, 2] : 1; refazer(); rdAgendarFrases(regra); }, RD_CLS_BTN, 'exigir também outro comparador'));
    return caixa;
}

/* `cond` é editado no lugar; `trocar(novo)` substitui o nó no pai; `remover` (opcional) tira o nó */
function rdEditorCondicao(cond, regra, trocar, remover, refazer) {
    const caixa = rdNo('div', 'border-l-2 border-[#30363d] pl-3 py-1 space-y-1');
    const tipo = rdTipoDe(cond);
    const topo = rdNo('div', 'flex flex-wrap items-center gap-2');
    const sel = rdNo('select', RD_CLS_INPUT);
    Object.entries(RD_TIPOS).forEach(([k, r]) => sel.appendChild(new Option(r, k)));
    sel.value = tipo;
    sel.onchange = () => {
        const novo = sel.value;
        const era_folha = ['tem', 'soma', 'texto'].includes(tipo);
        trocar(rdCondicaoPadrao(novo, era_folha || tipo === 'nao' ? cond : null));
        refazer(); rdAgendarFrases(regra);
    };
    topo.appendChild(sel);
    if (tipo === 'tem' || tipo === 'soma') {
        topo.appendChild(rdEditorSel(cond[tipo], v => { cond[tipo] = v; rdAgendarFrases(regra); }));
        topo.appendChild(rdEditorQtd(cond, regra, refazer));
    } else if (tipo === 'texto') {
        const inp = rdNo('input', RD_CLS_INPUT + ' w-64 font-mono');
        inp.value = cond.texto; inp.placeholder = 'regex sobre o texto da linha';
        inp.oninput = () => { cond.texto = inp.value; rdAgendarFrases(regra); };
        topo.appendChild(inp);
    }
    if (remover) topo.appendChild(rdBotao('remover', remover, RD_CLS_BTN + ' text-red-400'));
    caixa.appendChild(topo);
    if (['todos', 'algum', 'nenhum'].includes(tipo)) {
        cond[tipo].forEach((filho, i) => caixa.appendChild(rdEditorCondicao(
            filho, regra, novo => { cond[tipo][i] = novo; },
            cond[tipo].length > 1 ? () => { cond[tipo].splice(i, 1); refazer(); rdAgendarFrases(regra); } : null, refazer)));
        caixa.appendChild(rdBotao('+ condição', () => { cond[tipo].push({ tem: '' }); refazer(); }));
    } else if (tipo === 'nao') {
        caixa.appendChild(rdEditorCondicao(cond.nao, regra, novo => { cond.nao = novo; }, null, refazer));
    }
    return caixa;
}

function rdBloco(rotulo, dica, regra, campo, refazer) {
    const bloco = rdNo('div', 'mt-2');
    bloco.appendChild(rdNo('div', 'text-xs font-semibold text-[#d97706] uppercase', rotulo));
    if (dica) bloco.appendChild(rdNo('div', 'text-xs text-[#8b949e]', dica));
    if (regra[campo]) {
        bloco.appendChild(rdEditorCondicao(regra[campo], regra, novo => { regra[campo] = novo; },
            () => { delete regra[campo]; refazer(); rdAgendarFrases(regra); }, refazer));
    } else {
        bloco.appendChild(rdBotao('+ adicionar', () => { regra[campo] = { tem: '' }; refazer(); }));
    }
    return bloco;
}

/* ── cartão da regra ─────────────────────────────────────────────────────── */
function rdCartao(regra, i) {
    const herdada = rdEhProjeto() && regra.origem !== 'projeto';
    const card = rdNo('div', 'border border-[#30363d] rounded-md p-3');
    card.dataset.regra = regra.id;
    if (regra.oculta) card.classList.add('opacity-50');

    const topo = rdNo('div', 'flex flex-wrap items-center gap-2');
    const ativa = rdNo('input'); ativa.type = 'checkbox'; ativa.checked = !!regra.ativa; ativa.title = 'Regra ligada';
    ativa.onchange = () => { regra.ativa = ativa.checked; atualizarSelo(); };
    const selo = rdNo('span', 'text-xs px-2 py-0.5 rounded-full');
    const rid = rdNo('input', RD_CLS_INPUT + ' w-44 font-mono'); rid.value = regra.id || '';
    rid.readOnly = herdada; if (herdada) rid.title = 'Regra herdada do padrão: o código não muda';
    rid.oninput = () => { regra.id = rid.value.trim(); };
    const sev = rdNo('select', RD_CLS_INPUT);
    ['erro', 'aviso', 'info'].forEach(s => sev.appendChild(new Option(s, s)));
    sev.value = regra.severidade;
    sev.onchange = () => { regra.severidade = sev.value; atualizarSelo(); rdAgendarFrases(regra); };
    topo.append(ativa, selo, rid, rdNo('span', 'text-xs text-[#8b949e]', 'gravidade'), sev);

    const acoes = rdNo('span', 'ml-auto flex flex-wrap gap-2');
    if (herdada) {
        acoes.appendChild(rdBotao(regra.oculta ? '+' : '−', () => { regra.oculta = !regra.oculta; rdRenderizar(); },
            RD_CLS_BTN + ' font-bold w-9', regra.oculta ? 'Reexibir/religar esta regra neste projeto' : 'Ocultar: desliga esta regra só neste projeto (continua listada)'));
        if (RD.padrao && RD.padrao.regras[regra.id]) {
            acoes.appendChild(rdBotao('voltar ao padrão', () => {
                const base = rdClone(RD.padrao.regras[regra.id]);
                Object.keys(regra).forEach(k => { if (!['oculta', 'origem'].includes(k)) delete regra[k]; });
                Object.assign(regra, base, { oculta: regra.oculta, origem: 'padrao' });
                rdRenderizar();
            }, RD_CLS_BTN, 'Desfaz os ajustes desta regra neste projeto'));
        }
    }
    acoes.appendChild(rdBotao(regra._aberta ? 'Fechar' : 'Editar', () => { regra._aberta = !regra._aberta; rdRenderizar(); }));
    if (!herdada) {
        acoes.appendChild(rdBotao('Excluir', () => {
            if (confirm(`Excluir a regra ${regra.id}?`)) { RD.regras.splice(i, 1); rdRenderizar(); }
        }, RD_CLS_BTN + ' text-red-400'));
    }
    topo.appendChild(acoes);
    card.appendChild(topo);

    const frase = rdNo('div', 'text-sm mt-2', regra.frase || '');
    regra._frase = frase;
    card.appendChild(frase);
    if (regra.oculta) card.appendChild(rdNo('div', 'text-xs text-[#8b949e] mt-1', 'Oculta neste projeto: não é avaliada. Use + para reexibir.'));

    function atualizarSelo() {
        const o = rdEhProjeto() ? rdOrigemViva(regra) : null;
        selo.classList.toggle('hidden', !o);
        if (o) { selo.className = 'text-xs px-2 py-0.5 rounded-full ' + RD_SELOS[o][1]; selo.textContent = RD_SELOS[o][0]; }
    }
    atualizarSelo();
    if (regra._aberta) card.appendChild(rdCorpoEditor(regra, atualizarSelo));
    return card;
}

function rdCorpoEditor(regra, atualizarSelo) {
    const corpo = rdNo('div', 'mt-3 pt-3 border-t border-[#30363d] space-y-2');
    const refazer = () => { rdRenderizar(); atualizarSelo(); };
    const tocou = () => { atualizarSelo(); rdAgendarFrases(regra); };

    const msg = rdNo('input', RD_CLS_INPUT + ' w-full'); msg.value = regra.mensagem || ''; msg.placeholder = 'Mensagem mostrada ao usuário';
    msg.oninput = () => { regra.mensagem = msg.value; atualizarSelo(); };
    const ops = rdNo('input', RD_CLS_INPUT + ' w-40'); ops.value = (regra.operacoes || ['I', '*I']).join(', ');
    ops.oninput = () => { regra.operacoes = ops.value.split(',').map(s => s.trim()).filter(Boolean); tocou(); };
    const linha = rdNo('div', 'grid grid-cols-1 md:grid-cols-3 gap-3');
    const c1 = rdNo('label', 'block text-xs text-[#8b949e] md:col-span-2', 'Mensagem'); c1.appendChild(msg); msg.classList.add('block', 'mt-1');
    const c2 = rdNo('label', 'block text-xs text-[#8b949e]', 'Operações em que vale (I, *I, R…)'); c2.appendChild(ops); ops.classList.add('block', 'mt-1');
    linha.append(c1, c2);
    corpo.appendChild(linha);

    const avancado = !!regra._avancado || regra.escopo === 'planilha';
    if (regra.escopo === 'planilha') {
        corpo.appendChild(rdNo('div', 'text-xs text-[#8b949e]', 'Regra de total da planilha: edite pelo JSON abaixo (esq, cmp, dir).'));
    } else {
        corpo.appendChild(rdBotao(avancado ? '← Voltar ao editor visual' : 'Modo avançado (JSON)', () => {
            if (regra._jsonInvalido) { alert('Corrija o JSON antes de voltar ao editor visual.'); return; }
            regra._avancado = !regra._avancado; rdRenderizar();
        }));
    }
    if (avancado) {
        const ta = rdNo('textarea', RD_CLS_INPUT + ' w-full font-mono text-xs'); ta.rows = 12; ta.spellcheck = false;
        ta.value = JSON.stringify(rdLimpa(regra), null, 2);
        ta.oninput = () => {
            try {
                const novo = JSON.parse(ta.value);
                if (!novo || typeof novo !== 'object' || Array.isArray(novo)) throw new Error('objeto esperado');
                Object.keys(regra).forEach(k => { if (!RD_META.includes(k) && !k.startsWith('_')) delete regra[k]; });
                Object.assign(regra, novo);
                ta.classList.remove('border-red-500'); regra._jsonInvalido = false; tocou();
            } catch (e) { ta.classList.add('border-red-500'); regra._jsonInvalido = true; }
        };
        corpo.appendChild(ta);
    } else {
        corpo.appendChild(rdBloco('SE (gatilho)', 'A regra só vale nas linhas em que isto acontece. Vazio = todas as linhas.', regra, 'quando', refazer));
        corpo.appendChild(rdBloco('ENTÃO deve valer', 'O que precisa estar na linha. Vazio = a linha que cair no gatilho é proibida.', regra, 'entao', refazer));
        corpo.appendChild(rdBloco('EXCETO SE', 'Se isto valer, a regra não se aplica à linha.', regra, 'excecoes', refazer));
        const desc = rdNo('input', RD_CLS_INPUT + ' w-full mt-2'); desc.value = regra.descricao || ''; desc.placeholder = 'Anotação interna (opcional)';
        desc.oninput = () => { regra.descricao = desc.value; };
        corpo.appendChild(desc);
    }
    return corpo;
}

/* ── assistente de nova regra ────────────────────────────────────────────── */
const RD_MODELOS = {
    exigir: 'Exigir X quando houver Y',
    proibir: 'Proibir X junto de Y',
    minimo: 'Mínimo de X quando houver Y',
    total: 'Total da planilha'
};
function rdAbrirAssistente() {
    const box = rdEl('rdAssistente');
    box.replaceChildren();
    box.classList.remove('hidden');
    box.appendChild(rdNo('div', 'font-medium mb-2', 'Nova regra guiada'));
    const campos = {};
    const nova = (chave, rotulo, valor, extra) => {
        const l = rdNo('label', 'block text-xs text-[#8b949e]', rotulo);
        const i = rdNo('input', RD_CLS_INPUT + ' block mt-1 w-full'); i.value = valor || '';
        if (extra) i.setAttribute('list', 'rdSugestoes');
        l.appendChild(i); campos[chave] = i; return l;
    };
    const modelo = rdNo('select', RD_CLS_INPUT + ' block mt-1');
    Object.entries(RD_MODELOS).forEach(([k, r]) => modelo.appendChild(new Option(r, k)));
    const lm = rdNo('label', 'block text-xs text-[#8b949e]', 'Modelo'); lm.appendChild(modelo);
    const grade = rdNo('div', 'grid grid-cols-1 md:grid-cols-3 gap-3');
    const sev = rdNo('select', RD_CLS_INPUT + ' block mt-1');
    ['aviso', 'erro', 'info'].forEach(s => sev.appendChild(new Option(s, s)));
    const ls = rdNo('label', 'block text-xs text-[#8b949e]', 'Gravidade'); ls.appendChild(sev);
    const ajuda = rdNo('div', 'text-xs text-[#8b949e] mt-2');
    const tipoTotal = rdNo('select', RD_CLS_INPUT + ' block mt-1');
    tipoTotal.appendChild(new Option('metros nos Cabos', 'soma_metros'));
    tipoTotal.appendChild(new Option('quantidade nos Outros', 'soma_qtd'));
    const lt = rdNo('label', 'block text-xs text-[#8b949e]', 'Somar'); lt.appendChild(tipoTotal);
    const cmp = rdNo('select', RD_CLS_INPUT + ' block mt-1');
    ['>=', '<=', '=', '>', '<', '!='].forEach(o => cmp.appendChild(new Option(o, o)));
    const lc = rdNo('label', 'block text-xs text-[#8b949e]', 'Deve ser'); lc.appendChild(cmp);
    const partes = [lm, ls, nova('id', 'Código da regra', ''), nova('x', 'X (ativo, curinga ou @GRUPO)', '', true), nova('y', 'Y (ativo, curinga ou @GRUPO)', '', true),
        nova('n', 'Quantidade / valor N', '1'), lt, lc, nova('msg', 'Mensagem ao usuário', '')];
    grade.append(...partes);
    box.append(grade, ajuda);
    const dicas = {
        exigir: 'Em cada linha em que houver Y, precisa haver X (quantidade ≥ N).',
        proibir: 'Em cada linha em que houver Y, X não pode aparecer.',
        minimo: 'Em cada linha em que houver Y, a soma de X precisa ser ≥ N.',
        total: 'Compara a soma de X na planilha inteira com N (ex.: metros de P50 nos Cabos ≥ 100). Y não é usado.'
    };
    const mostrar = () => {
        ajuda.textContent = dicas[modelo.value];
        const total = modelo.value === 'total';
        campos.y.parentElement.classList.toggle('hidden', total);
        lt.classList.toggle('hidden', !total); lc.classList.toggle('hidden', !total);
        campos.n.parentElement.classList.toggle('hidden', modelo.value === 'proibir');
    };
    modelo.onchange = mostrar; mostrar();
    const criar = rdBotao('Criar regra (desligada)', () => {
        let id = campos.id.value.trim();
        if (!id) { let n = RD.regras.length + 1; while (RD.regras.some(r => r.id === `C2-NOVA-${n}`)) n++; id = (rdEhProjeto() ? 'PROJ-' : 'C2-') + 'NOVA-' + n; }
        const x = rdTextoParaSel(campos.x.value), y = rdTextoParaSel(campos.y.value), n = rdTextoParaVal(campos.n.value || '1');
        if (!campos.x.value.trim() || (modelo.value !== 'total' && !campos.y.value.trim())) { rdErros(['Preencha X' + (modelo.value === 'total' ? '' : ' e Y') + '.']); return; }
        const base = { id, versao: 2, escopo: 'linha', operacoes: ['I', '*I'], ativa: false, severidade: sev.value,
            mensagem: campos.msg.value.trim() || 'Regra nova: ajuste a mensagem.', descricao: '' };
        let regra;
        if (modelo.value === 'exigir') regra = { ...base, quando: { tem: y }, entao: { tem: x, qtd: { '>=': n } } };
        else if (modelo.value === 'proibir') regra = { ...base, quando: { tem: y }, entao: { nao: { tem: x } } };
        else if (modelo.value === 'minimo') regra = { ...base, quando: { tem: y }, entao: { soma: x, qtd: { '>=': n } } };
        else regra = { ...base, escopo: 'planilha', operacoes: undefined, deve_ser: { esq: { [tipoTotal.value]: x }, cmp: cmp.value, dir: n } };
        if (regra.operacoes === undefined) delete regra.operacoes;
        if (rdEhProjeto()) regra.origem = 'projeto';
        regra._aberta = true;
        RD.regras.push(regra);
        rdErros(null); box.classList.add('hidden');
        rdRenderizar(); rdAtualizarFrase(regra); rdAgendarFrases(regra);
    }, 'btn-primary font-medium py-1 px-4 rounded-md');
    box.appendChild(rdNo('div', 'mt-3 flex gap-2')).append(criar, rdBotao('Cancelar', () => box.classList.add('hidden')));
}

/* ── grupos ──────────────────────────────────────────────────────────────── */
function rdGrupoParaTexto(v) { return Array.isArray(v) && v.every(x => typeof x === 'string') ? v.join(', ') : JSON.stringify(v); }
function rdTextoParaGrupo(t) {
    t = t.trim();
    if (t.startsWith('[') || t.startsWith('{')) { try { return JSON.parse(t); } catch (e) { return t; } }
    return t.split(',').map(s => s.trim()).filter(Boolean);
}
function rdGrupoOrigemViva(nome) {
    if (!rdEhProjeto()) return null;
    if (!RD.padrao || !(nome in RD.padrao.grupos)) return 'projeto';
    return JSON.stringify(RD.padrao.grupos[nome]) === JSON.stringify(RD.grupos[nome]) ? 'padrao' : 'sobrescrita';
}
function rdRenderizarGrupos() {
    const box = rdEl('rdGrupos');
    box.replaceChildren();
    Object.keys(RD.grupos).sort().forEach(nome => {
        const linha = rdNo('div', 'border border-[#30363d] rounded-md p-2');
        const topo = rdNo('div', 'flex flex-wrap items-center gap-2');
        topo.appendChild(rdNo('span', 'font-mono text-sm', '@' + nome));
        const selo = rdNo('span', 'text-xs px-2 py-0.5 rounded-full');
        topo.appendChild(selo);
        const usados = (RD.grupos_usados[nome] || []);
        topo.appendChild(rdNo('span', 'text-xs text-[#8b949e]', usados.length
            ? `usado por ${usados.length} regra(s): ${usados.join(', ')} — mudar o grupo muda todas`
            : 'não usado por nenhuma regra'));
        const ta = rdNo('input', RD_CLS_INPUT + ' w-full font-mono text-xs mt-2');
        ta.value = rdGrupoParaTexto(RD.grupos[nome]);
        ta.title = 'Ativos do grupo separados por vírgula (aceita curinga *, ? e outros @GRUPO)';
        const volta = rdBotao('voltar ao padrão', () => {
            RD.grupos[nome] = rdClone(RD.padrao.grupos[nome]); ta.value = rdGrupoParaTexto(RD.grupos[nome]);
            atualizarSelo(); rdAgendarFrases();
        });
        function atualizarSelo() {   // atualiza no lugar: recriar a lista aqui engoliria o clique no campo seguinte
            const o = rdGrupoOrigemViva(nome);
            selo.classList.toggle('hidden', !o);
            if (o) { selo.className = 'text-xs px-2 py-0.5 rounded-full ' + RD_SELOS[o][1]; selo.textContent = RD_SELOS[o][0]; }
            volta.classList.toggle('hidden', o !== 'sobrescrita' || !RD.padrao);
        }
        ta.oninput = () => { RD.grupos[nome] = rdTextoParaGrupo(ta.value); atualizarSelo(); rdAgendarFrases(); };
        const acoes = rdNo('span', 'ml-auto flex gap-2');
        acoes.appendChild(volta);
        if (!rdEhProjeto() || rdGrupoOrigemViva(nome) === 'projeto') {
            acoes.appendChild(rdBotao('Excluir', () => {
                if (confirm(`Excluir o grupo ${nome}?` + (usados.length ? ` Ele é usado por: ${usados.join(', ')}.` : ''))) { delete RD.grupos[nome]; rdRenderizar(); }
            }, RD_CLS_BTN + ' text-red-400'));
        }
        topo.appendChild(acoes);
        atualizarSelo();
        linha.append(topo, ta);
        box.appendChild(linha);
    });
    if (!Object.keys(RD.grupos).length) box.appendChild(rdNo('div', 'text-sm text-[#8b949e]', 'Nenhum grupo definido.'));
}
function rdNovoGrupo() {
    const nome = (prompt('Nome do novo grupo (letras, números e _):') || '').trim().toUpperCase();
    if (!nome) return;
    if (!/^[A-Z0-9_]+$/.test(nome)) { showMessage('error', 'Nome inválido: use letras, números e _.'); return; }
    if (RD.grupos[nome]) { showMessage('error', 'Já existe um grupo com esse nome.'); return; }
    RD.grupos[nome] = [];
    RD.grupos_origem[nome] = 'projeto';
    rdRenderizar();
}

/* ── salvar / semente / histórico / teste ────────────────────────────────── */
function rdRegrasParaEnvio() {
    if (RD.regras.some(r => r._jsonInvalido)) { rdErros(['Há regras com JSON inválido (campo em vermelho no modo avançado).']); return null; }
    return RD.regras.map(r => { const c = rdLimpa(r); if (r.oculta) c.oculta = true; if (r.origem) c.origem = r.origem; return c; });
}
async function rdSalvar() {
    rdErros(null);
    const regras = rdRegrasParaEnvio();
    if (!regras) return;
    const res = await rdPost('/api/validacao/regras', { projeto_codigo: rdProjeto(), regras, grupos: RD.grupos });
    if (!res.ok) { showMessage('error', await rdLerErro(res)); return; }
    showMessage('success', 'Regras salvas.');
    await rdCarregar();
}
async function rdRestaurarSemente() {
    const msg = rdEhProjeto()
        ? 'Descartar todos os ajustes deste projeto e voltar ao padrão? A versão atual fica no histórico.'
        : 'Substituir as regras pelas da semente do sistema (todas desligadas)? A versão atual fica no histórico.';
    if (!confirm(msg)) return;
    rdErros(null);
    const res = await rdPost('/api/validacao/regras/restaurar-semente', { projeto_codigo: rdProjeto() });
    if (!res.ok) { showMessage('error', await rdLerErro(res)); return; }
    showMessage('success', rdEhProjeto() ? 'Ajustes do projeto descartados.' : 'Semente restaurada.');
    await rdCarregar();
}
async function rdAdicionarNovas() {
    rdErros(null);
    const res = await rdPost('/api/validacao/regras/adicionar-novas', { projeto_codigo: rdProjeto() });
    if (!res.ok) { showMessage('error', await rdLerErro(res)); return; }
    const { adicionadas } = await res.json();
    showMessage('success', adicionadas.length
        ? `Adicionadas (desligadas): ${adicionadas.join(', ')}. Os projetos passam a herdá-las.`
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
    if (!historico.length) box.appendChild(rdNo('div', 'p-3 text-[#8b949e]', 'Sem versões anteriores para este projeto.'));
    historico.forEach(h => {
        const linha = rdNo('div', 'flex items-center justify-between gap-3 p-3');
        linha.appendChild(rdNo('span', '', `${h.criado_em} — ${h.criado_por || 'desconhecido'} — ${h.regras.length} regras, ${h.regras.filter(r => r.ativa).length} ligadas`));
        const acoes = rdNo('span', 'space-x-3');
        const ver = rdNo('button', 'text-blue-400 hover:text-blue-300', 'Ver no editor');
        ver.onclick = () => { RD.regras = rdClone(h.regras); RD.grupos = rdClone(h.grupos); rdErros(null); rdRenderizar(); rdAgendarFrases(); };
        const rev = rdNo('button', 'text-[#d97706] hover:text-[#b45309]', 'Reverter');
        rev.onclick = () => rdReverter(h.id);
        acoes.append(ver, rev);
        linha.appendChild(acoes);
        box.appendChild(linha);
    });
    box.classList.remove('hidden');
}
async function rdReverter(historicoId) {
    if (!confirm('Reverter para esta versão? A versão atual fica no histórico.')) return;
    const res = await rdPost('/api/validacao/regras/reverter', { projeto_codigo: rdProjeto(), historico_id: historicoId });
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
    const res = await rdPost('/api/validacao/regras/testar', {
        regras, grupos: RD.grupos, explicar: true, cabos: rdLinhasDeTeste('rdTesteCabos'), outros: rdLinhasDeTeste('rdTesteOutros')
    });
    const alvo = rdEl('rdTesteResultado');
    alvo.replaceChildren();
    if (!res.ok) { showMessage('error', await rdLerErro(res)); return; }
    const { achados, resumo, explicacoes } = await res.json();
    alvo.appendChild(rdNo('div', 'mb-1 font-medium', achados.length
        ? `${achados.length} achado(s): ${resumo.erro} erro, ${resumo.aviso} aviso, ${resumo.info} info`
        : 'Nenhum achado para estas linhas com as regras ligadas.'));
    const mostrarTodas = rdEl('rdTesteTodas').checked;
    const rotulo = { alerta: 'DISPAROU', ok: 'não disparou', nao_aplica: 'não se aplica' };
    (explicacoes || []).filter(x => mostrarTodas || x.situacao === 'alerta').forEach(x => {
        const item = rdNo('div', 'py-0.5' + (x.situacao === 'alerta' ? '' : ' text-[#8b949e]'));
        item.textContent = `[${x.situacao === 'alerta' ? x.severidade : rotulo[x.situacao]}] ${x.linha_id} · ${x.regra_id} — ` +
            (x.situacao === 'alerta' ? `${x.mensagem} (${x.explicacao})` : x.explicacao);
        alvo.appendChild(item);
    });
}
