/* ═══════════════════════════════════════════════════════════════════════════
   Aba "Ajustes" da gaveta de regras do Resumo (TASK-023, ADR-006).
   Receitas de ajuste recorrentes (substituir texto, normalizar, ordenar, excluir/adicionar linhas…), por projeto em camadas
   (padrão + ajustes do projeto: selos, −/+, "voltar ao padrão"), com pré-visualização do diff nas tabelas atuais.
   A edição fica em RASCUNHO até o admin clicar em Salvar (grava local e no Supabase). Aplicar nas tabelas é a TASK-024.
   Nenhuma lógica de ajuste aqui: frases, validação e diff vêm do backend. Conteúdo sempre via textContent/value.
   Reaproveita helpers de regras_editor.js (rdNo, rdBotao, rdEditorCondicao, seletores…), carregado antes.
═══════════════════════════════════════════════════════════════════════════ */
const AJ = { itens: [], padrao: null, sujo: false, timer: null };
const AJ_TABELAS = { cabos: 'Cabos', outros: 'Outros', ambos: 'Cabos e Outros' };
const AJ_OPERACOES = ['I', '*I', 'R', '*R', 'M', '*M'];
const AJ_ACOES = {
    substituir: { rot: 'Substituir texto', tabelas: ['outros', 'cabos', 'ambos'], filtros: true,
        novo: () => ({ acao: 'substituir', tabela: 'outros', de: '', para: '' }) },
    normalizar: { rot: 'Normalizar', tabelas: ['outros', 'cabos', 'ambos'], filtros: true,
        novo: () => ({ acao: 'normalizar', tabela: 'outros', regras: ['espacos'] }) },
    ordenar: { rot: 'Ordenar', tabelas: ['outros', 'cabos', 'ambos'], filtros: false,
        novo: () => ({ acao: 'ordenar', tabela: 'outros', por: [{ coluna: 'ativo' }] }) },
    excluir_linhas: { rot: 'Excluir linhas', tabelas: ['outros', 'cabos', 'ambos'], filtros: true,
        novo: () => ({ acao: 'excluir_linhas', tabela: 'outros', onde: { vazias: true } }) },
    adicionar_linha: { rot: 'Adicionar linha', tabelas: ['outros', 'cabos'], filtros: false,
        novo: () => ({ acao: 'adicionar_linha', tabela: 'outros', valores: { operacao: 'I', ativo: '' }, posicao: 'fim' }) },
    adicionar_ativo: { rot: 'Adicionar ativo à linha', tabelas: ['outros'], filtros: true,
        novo: () => ({ acao: 'adicionar_ativo', tabela: 'outros', ativo: '', qtd: 1 }) },
    remover_ativo: { rot: 'Remover ativo da linha', tabelas: ['outros'], filtros: true,
        novo: () => ({ acao: 'remover_ativo', tabela: 'outros', ativo: '' }) },
    mesclar_duplicadas: { rot: 'Mesclar ativos repetidos na linha', tabelas: ['outros'], filtros: true,
        novo: () => ({ acao: 'mesclar_duplicadas', tabela: 'outros' }) }
};

function ajEl(id) { return document.getElementById(id); }
function ajProjeto() { return rdEl('rdProjeto').value || 'DEFAULT'; }
function ajEhProjeto() { return ajProjeto() !== 'DEFAULT'; }
/** Contexto ativo (TASK-044) — só faz sentido dentro de um projeto (nunca no DEFAULT); lido do
 * seletor da linha de totais do Resumo (resumo.js:contextoSelecionado), nunca editável aqui. */
function ajContexto() { return (ajEhProjeto() && typeof window.contextoSelecionado === 'function') ? window.contextoSelecionado() : null; }
function ajCsv(txt) { return txt.split(',').map(s => s.trim()).filter(Boolean); }
function ajLimpa(r) { const c = {}; Object.keys(r).forEach(k => { if (!['origem', 'oculta', 'frases'].includes(k) && !k.startsWith('_')) c[k] = r[k]; }); return c; }

function ajMarcarSujo() {
    AJ.sujo = true;
    const s = ajEl('ajStatus');
    if (s && !s.dataset.sujo) { s.dataset.base = s.textContent; s.dataset.sujo = '1'; }
    if (s) s.textContent = (s.dataset.base || '') + ' · Rascunho não salvo — clique em Salvar para gravar';
}

function ajErros(lista) {
    const box = ajEl('ajErros');
    box.replaceChildren();
    if (!lista || !lista.length) { box.classList.add('hidden'); return; }
    lista.forEach(e => box.appendChild(rdNo('div', '', '• ' + e)));
    box.classList.remove('hidden');
}
async function ajLerErro(res) {
    let corpo = {};
    try { corpo = await res.json(); } catch (e) { /* sem corpo */ }
    const d = corpo.detail;
    if (d && d.erros) { ajErros(d.erros); return 'Ajustes inválidos — veja os erros acima.'; }
    return (typeof d === 'string' && d) || 'Erro ao processar a requisição.';
}

async function iniciarAjustes() {
    const lig = (id, fn) => { const e = ajEl(id); if (e) e.addEventListener('click', fn); };
    lig('ajAdicionar', ajNovo);
    lig('ajSalvar', ajSalvar);
    lig('ajSemente', ajRestaurarSemente);
    lig('ajHistorico', ajAlternarHistorico);
}

async function ajCarregar() {
    const contexto = ajContexto();
    const qs = `projeto_codigo=${encodeURIComponent(ajProjeto())}` + (contexto ? `&contexto=${encodeURIComponent(contexto)}` : '');
    const res = await fetch(`/api/validacao/ajustes?${qs}`);
    if (!res.ok) { rpMensagem('error', 'Não foi possível carregar os ajustes.'); return; }
    const d = await res.json();
    AJ.itens = d.ajustes;
    AJ.contexto = contexto;
    AJ.padrao = null;
    // Base para o selo "sobrescrita"/"padrão" ao vivo: COM contexto, a base é o efetivo do PRÓPRIO
    // projeto (sem contexto) — é contra ela que o overlay do contexto é calculado. SEM contexto,
    // a base continua sendo o DEFAULT, como antes desta tarefa.
    if (contexto) {
        try {
            const p = await fetch(`/api/validacao/ajustes?projeto_codigo=${encodeURIComponent(ajProjeto())}`);
            if (p.ok) AJ.padrao = Object.fromEntries((await p.json()).ajustes.map(r => [r.id, r]));
        } catch (e) { /* sem base: desativa "voltar ao padrão" */ }
    } else if (ajEhProjeto()) {
        try {
            const p = await fetch('/api/validacao/ajustes?projeto_codigo=DEFAULT');
            if (p.ok) AJ.padrao = Object.fromEntries((await p.json()).ajustes.map(r => [r.id, r]));
        } catch (e) { /* sem padrão: desativa "voltar ao padrão" */ }
    }
    AJ.sujo = false;
    const st = ajEl('ajStatus');
    delete st.dataset.sujo;
    st.textContent = contexto
        ? `Editando o contexto "${contexto}" — histórico e "restaurar" valem só para o projeto, não para o contexto`
        : !ajEhProjeto() ? 'Padrão do sistema (vale para todos os projetos)'
        : (d.personalizado ? 'Este projeto tem ajustes sobre o padrão' : 'Este projeto usa o padrão sem ajustes');
    ajEl('ajSemente').textContent = ajEhProjeto() ? 'Descartar ajustes do projeto' : 'Restaurar semente';
    ajEl('ajSemente').disabled = !!contexto;
    ajEl('ajHistorico').disabled = !!contexto;
    ajEl('ajAdicionar').textContent = contexto ? `Novo ajuste só no contexto "${contexto}"` : (ajEhProjeto() ? 'Novo ajuste só neste projeto' : 'Novo ajuste');
    ajEl('ajHistLista').classList.add('hidden');
    ajErros(null);
    const av = ajEl('ajAvisos');
    av.replaceChildren();
    (d.avisos || []).forEach(a => av.appendChild(rdNo('div', '', '• ' + a)));
    av.classList.toggle('hidden', !(d.avisos || []).length);
    ajRenderizar();
    // sugestões de ids de regras para o vínculo
    const dl = ajEl('ajRegrasSugestoes');
    dl.replaceChildren();
    (typeof RD !== 'undefined' ? RD.regras : []).forEach(r => dl.appendChild(new Option(r.id)));
}

/* ── origem (selo) ao vivo ────────────────────────────────────────────────── */
function ajOrigemViva(item) {
    if (!ajEhProjeto()) return null;
    if (item.origem === 'projeto') return 'projeto';
    const base = AJ.padrao && AJ.padrao[item.id];
    if (!base) return item.origem || 'padrao';
    return JSON.stringify(ajLimpa(base)) === JSON.stringify(ajLimpa(item)) ? 'padrao' : 'sobrescrita';
}

/* ── frases (só o backend descreve) ───────────────────────────────────────── */
function ajAgendarFrases(item) {
    clearTimeout(AJ.timer);
    AJ.timer = setTimeout(() => ajAtualizarFrases(item), 350);
}
async function ajAtualizarFrases(item) {
    if (item._jsonInvalido || !item._frases) return;
    try {
        const res = await rdPost('/api/validacao/ajustes/descrever', { acoes: item.acoes, projeto_codigo: ajEhProjeto() ? ajProjeto() : null });
        if (!res.ok) return;
        const d = await res.json();
        item._frases.replaceChildren();
        (d.erros.length ? d.erros.map(e => '⚠ ' + e) : d.frases).forEach(f => item._frases.appendChild(rdNo('div', d.erros.length ? 'text-red-400' : '', f)));
    } catch (e) { /* mantém a frase anterior */ }
}

/* ── campos de uma ação ───────────────────────────────────────────────────── */
function ajInput(valor, aoMudar, cls, dica, lista) {
    const i = rdNo('input', RD_CLS_INPUT + ' ' + (cls || 'w-44'));
    i.value = valor == null ? '' : valor;
    if (dica) i.title = dica;
    if (lista) i.setAttribute('list', lista);
    i.oninput = () => aoMudar(i.value);
    return i;
}
function ajSelect(opcoes, valor, aoMudar) {
    const s = rdNo('select', RD_CLS_INPUT);
    opcoes.forEach(([v, r]) => s.appendChild(new Option(r, v)));
    s.value = valor;
    s.onchange = () => aoMudar(s.value);
    return s;
}
function ajRotulo(texto, el) {
    const l = rdNo('label', 'inline-flex items-center gap-1 text-xs text-[#8b949e]', texto);
    l.appendChild(el);
    return l;
}
function ajCondicao(acao, campo, rotulo, refazer, container) {
    const bloco = rdNo('div', 'mt-2');
    bloco.appendChild(rdNo('div', 'text-xs font-semibold text-[#d97706] uppercase', rotulo));
    if (container[campo]) {
        bloco.appendChild(rdEditorCondicao(container[campo], acao, novo => { container[campo] = novo; },
            () => { delete container[campo]; refazer(); }, refazer));
    } else {
        bloco.appendChild(rdBotao('+ adicionar condição', () => { container[campo] = { tem: '' }; refazer(); }));
    }
    return bloco;
}

function ajCamposDaAcao(a, refazer) {
    const f = rdNo('div', 'flex flex-wrap items-center gap-2 mt-2');
    const nome = a.acao;
    if (nome === 'substituir') {
        const modo = a.modo || 'texto';
        f.appendChild(ajRotulo('modo', ajSelect([['texto', 'texto'], ['item', 'item (código do ativo)']], modo, v => {
            if (v === 'texto') delete a.modo; else { a.modo = v; a.tabela = 'outros'; }
            refazer();
        })));
        const de = modo === 'item' ? rdSelParaTexto(a.de) : (a.de && typeof a.de === 'object' ? '/' + a.de.regex + '/' : a.de);
        f.appendChild(ajRotulo('trocar', ajInput(de, v => {
            a.de = modo === 'item' ? rdTextoParaSel(v) : (v.length > 2 && v.startsWith('/') && v.endsWith('/') ? { regex: v.slice(1, -1) } : v);
        }, 'w-44 font-mono', modo === 'item' ? 'código, curinga, @GRUPO, lista' : 'texto (ex.: SUP-L) ou /regex/', 'rdSugestoes')));
        f.appendChild(ajRotulo('por', ajInput(a.para, v => { a.para = v; }, 'w-44 font-mono')));
        if (modo === 'texto') {
            const c = document.createElement('input'); c.type = 'checkbox'; c.checked = a.palavra_inteira !== false;
            c.onchange = () => { if (c.checked) delete a.palavra_inteira; else a.palavra_inteira = false; };
            f.appendChild(ajRotulo('só palavra inteira', c));
        }
    } else if (nome === 'normalizar') {
        [['espacos', 'espaços'], ['maiusculas', 'maiúsculas'], ['poste', "poste sem '1-' no início"]].forEach(([k, r]) => {
            const c = document.createElement('input'); c.type = 'checkbox'; c.checked = (a.regras || []).includes(k);
            c.onchange = () => { a.regras = ['espacos', 'maiusculas', 'poste'].filter(x => x === k ? c.checked : (a.regras || []).includes(x)); };
            f.appendChild(ajRotulo(r, c));
        });
    } else if (nome === 'ordenar') {
        const col = rdNo('div', 'space-y-1 w-full');
        a.por.forEach((c, i) => {
            const l = rdNo('div', 'flex flex-wrap items-center gap-2');
            l.append(
                ajSelect([['operacao', 'operação'], ['ativo', 'ativo'], ['entidade', 'entidade']], c.coluna, v => { c.coluna = v; }),
                ajSelect([['asc', 'crescente'], ['desc', 'decrescente']], c.ordem || 'asc', v => { c.ordem = v; }),
                ajRotulo('ordem própria', ajInput((c.valores || []).join(', '), v => { const l2 = ajCsv(v); if (l2.length) c.valores = l2; else delete c.valores; }, 'w-44', 'valores separados por vírgula, ex.: I, *I, R')));
            if (a.por.length > 1) l.appendChild(rdBotao('×', () => { a.por.splice(i, 1); refazer(); }));
            col.appendChild(l);
        });
        col.appendChild(rdBotao('+ critério', () => { a.por.push({ coluna: 'ativo' }); refazer(); }));
        f.appendChild(col);
    } else if (nome === 'excluir_linhas') {
        const tipo = Object.keys(a.onde)[0];
        f.appendChild(ajRotulo('excluir', ajSelect([['vazias', 'linhas vazias'], ['duplicadas', 'linhas duplicadas'], ['condicao', 'linhas que satisfazem...'], ['texto', 'linhas cujo texto casa com...']], tipo, v => {
            a.onde = v === 'condicao' ? { condicao: { tem: '' } } : v === 'texto' ? { texto: '' } : { [v]: true };
            refazer();
        })));
        if (tipo === 'texto') f.appendChild(ajInput(a.onde.texto, v => { a.onde.texto = v; }, 'w-56 font-mono', 'regex'));
        if (tipo === 'condicao') f.appendChild(ajCondicao(a, 'condicao', 'Condição', refazer, a.onde));
    } else if (nome === 'adicionar_linha') {
        f.appendChild(ajRotulo('operação', ajSelect(AJ_OPERACOES.map(o => [o, o]), a.valores.operacao, v => { a.valores.operacao = v; })));
        f.appendChild(ajRotulo('ativo', ajInput(a.valores.ativo, v => { a.valores.ativo = v; }, 'w-56 font-mono', 'ex.: 1-RA2')));
        const pos = typeof a.posicao === 'string' ? a.posicao : Object.keys(a.posicao)[0];
        f.appendChild(ajRotulo('posição', ajSelect([['fim', 'no fim'], ['inicio', 'no início'], ['depois_de', 'depois de cada linha que...'], ['antes_de', 'antes de cada linha que...']], pos, v => {
            a.posicao = v === 'fim' || v === 'inicio' ? v : { [v]: { tem: '' } };
            refazer();
        })));
        const c = document.createElement('input'); c.type = 'checkbox'; c.checked = a.apenas_se_nao_existir !== false;
        c.onchange = () => { if (c.checked) delete a.apenas_se_nao_existir; else a.apenas_se_nao_existir = false; };
        f.appendChild(ajRotulo('só se ainda não existir', c));
        if (typeof a.posicao === 'object') f.appendChild(ajCondicao(a, Object.keys(a.posicao)[0], 'Linha de referência', refazer, a.posicao));
    } else if (nome === 'adicionar_ativo') {
        const ativoDinamico = a.ativo && typeof a.ativo === 'object';
        // TASK-046: ativo fixo (código literal) ou dinâmico (nome = prefixo + código encontrado por um
        // seletor na própria linha + sufixo — ex.: "se tem TR1*, adicionar o próprio código + 'VP'").
        f.appendChild(ajRotulo('ativo', ajSelect([['fixo', 'fixo'], ['dinamico', 'igual ao ativo encontrado']], ativoDinamico ? 'dinamico' : 'fixo', v => {
            a.ativo = v === 'dinamico' ? { igual_a: '' } : '';
            refazer();
        })));
        if (ativoDinamico) {
            f.appendChild(ajRotulo('igual a', ajInput(rdSelParaTexto(a.ativo.igual_a), v => { a.ativo.igual_a = rdTextoParaSel(v); }, 'w-44 font-mono', 'código, curinga, @GRUPO, lista', 'rdSugestoes')));
            f.appendChild(ajRotulo('prefixo', ajInput(a.ativo.prefixo || '', v => { if (v) a.ativo.prefixo = v; else delete a.ativo.prefixo; }, 'w-20 font-mono')));
            f.appendChild(ajRotulo('sufixo', ajInput(a.ativo.sufixo || '', v => { if (v) a.ativo.sufixo = v; else delete a.ativo.sufixo; }, 'w-20 font-mono')));
        } else {
            f.appendChild(ajRotulo('ativo', ajInput(a.ativo, v => { a.ativo = v.trim(); }, 'w-44 font-mono', 'código, ex.: SUPL', 'rdSugestoes')));
        }
        if (ativoDinamico) {
            // Com ativo dinâmico, `qtd` é só o FATOR (número) aplicado à quantidade de cada item
            // encontrado — não aceita o modo "mesma quantidade de outro ativo" (soma) aqui.
            f.appendChild(ajRotulo('fator', ajInput(a.qtd != null && typeof a.qtd !== 'object' ? a.qtd : 1, v => { const n = Number(v.replace(',', '.')); a.qtd = isNaN(n) ? v : n; }, 'w-16', '-1 = negativo (retirada)')));
        } else {
            const qtdDinamica = a.qtd && typeof a.qtd === 'object';
            // TASK-045: qtd fixa (número) ou dinâmica (mesma quantidade de outro ativo da própria linha,
            // opcionalmente multiplicada por um fator — ex.: fator -1 = negativo/retirada).
            f.appendChild(ajRotulo('quantidade', ajSelect([['fixa', 'fixa'], ['dinamica', 'mesma quantidade de outro ativo']], qtdDinamica ? 'dinamica' : 'fixa', v => {
                a.qtd = v === 'dinamica' ? { soma: '', fator: 1 } : 1;
                refazer();
            })));
            if (qtdDinamica) {
                f.appendChild(ajRotulo('soma de', ajInput(rdSelParaTexto(a.qtd.soma), v => { a.qtd.soma = rdTextoParaSel(v); }, 'w-44 font-mono', 'código, curinga, @GRUPO, lista', 'rdSugestoes')));
                f.appendChild(ajRotulo('fator', ajInput(a.qtd.fator != null ? a.qtd.fator : 1, v => { const n = Number(v.replace(',', '.')); a.qtd.fator = isNaN(n) ? v : n; }, 'w-16', '-1 = negativo (retirada)')));
            } else {
                f.appendChild(ajRotulo('qtd', ajInput(a.qtd, v => { const n = Number(v.replace(',', '.')); a.qtd = isNaN(n) ? v : n; }, 'w-20')));
            }
        }
        f.appendChild(ajRotulo('se já existir', ajSelect([['ignorar', 'ignorar'], ['somar', 'somar a quantidade'], ['substituir', 'trocar a quantidade']], a.se_ja_existe || 'ignorar', v => { if (v === 'ignorar') delete a.se_ja_existe; else a.se_ja_existe = v; })));
    } else if (nome === 'remover_ativo') {
        f.appendChild(ajRotulo('ativo', ajInput(rdSelParaTexto(a.ativo), v => { a.ativo = rdTextoParaSel(v); }, 'w-56 font-mono', 'código, curinga, @GRUPO, lista', 'rdSugestoes')));
    }
    return f;
}

function ajAcaoEditor(item, a, i, refazer) {
    const spec = AJ_ACOES[a.acao] || AJ_ACOES.normalizar;
    const caixa = rdNo('div', 'border rounded-md p-2');
    const topo = rdNo('div', 'flex flex-wrap items-center gap-2');
    topo.append(
        rdNo('span', 'text-xs text-[#8b949e]', `Ação ${i + 1}`),
        ajSelect(Object.entries(AJ_ACOES).map(([k, v]) => [k, v.rot]), a.acao, v => {
            const n = AJ_ACOES[v].novo();
            if (AJ_ACOES[v].tabelas.includes(a.tabela)) n.tabela = a.tabela;
            item.acoes[i] = n; refazer();
        }),
        ajSelect(spec.tabelas.map(t => [t, AJ_TABELAS[t]]), spec.tabelas.includes(a.tabela) ? a.tabela : spec.tabelas[0], v => { a.tabela = v; refazer(); }));
    const acoes = rdNo('span', 'ml-auto flex gap-1');
    if (i > 0) acoes.appendChild(rdBotao('↑', () => { [item.acoes[i - 1], item.acoes[i]] = [item.acoes[i], item.acoes[i - 1]]; refazer(); }));
    if (i < item.acoes.length - 1) acoes.appendChild(rdBotao('↓', () => { [item.acoes[i + 1], item.acoes[i]] = [item.acoes[i], item.acoes[i + 1]]; refazer(); }));
    if (item.acoes.length > 1) acoes.appendChild(rdBotao('×', () => { item.acoes.splice(i, 1); refazer(); }));
    topo.appendChild(acoes);
    caixa.append(topo, ajCamposDaAcao(a, refazer));
    if (spec.filtros) {
        const fl = rdNo('div', 'flex flex-wrap items-center gap-2 mt-2');
        fl.appendChild(ajRotulo('só nas operações', ajInput((a.operacoes || []).join(', '), v => { const l = ajCsv(v); if (l.length) a.operacoes = l; else delete a.operacoes; }, 'w-32', 'ex.: I, *I (vazio = todas)')));
        if (['substituir', 'adicionar_ativo', 'remover_ativo', 'mesclar_duplicadas'].includes(a.acao)) {
            fl.appendChild(ajRotulo('mudar a operação para', ajSelect([['', '(manter)'], ...AJ_OPERACOES.map(o => [o, o])], a.operacao_nova || '', v => { if (v) a.operacao_nova = v; else delete a.operacao_nova; })));
        }
        caixa.appendChild(fl);
        caixa.appendChild(ajCondicao(a, 'quando', 'Só nas linhas em que (opcional)', refazer, a));
    }
    return caixa;
}

/* ── pré-visualização do diff ─────────────────────────────────────────────── */
async function ajPreVisualizar(item, alvo) {
    alvo.replaceChildren();
    if (typeof window.resumoLinhasParaAjuste !== 'function') { alvo.textContent = 'Tabelas do Resumo indisponíveis.'; return; }
    const { cabos, outros } = window.resumoLinhasParaAjuste();
    const res = await rdPost('/api/validacao/ajustes/preview', { acoes: item.acoes, cabos, outros, projeto_codigo: ajEhProjeto() ? ajProjeto() : null });
    if (!res.ok) {
        const msg = await ajLerErro(res);
        alvo.appendChild(rdNo('div', 'text-red-400', msg));
        return;
    }
    const r = await res.json();
    const n = r.resumo;
    alvo.appendChild(rdNo('div', 'font-medium', (n.editar + n.inserir + n.excluir + n.mover)
        ? `Pré-visualização: ${n.editar} editada(s), ${n.inserir} inserida(s), ${n.excluir} excluída(s)${n.mover ? ', ordem alterada' : ''}. Nada foi aplicado.`
        : 'Pré-visualização: nada mudaria nas tabelas atuais.'));
    const mostrar = (txt, cls) => alvo.appendChild(rdNo('div', 'text-xs font-mono ' + (cls || ''), txt));
    r.operacoes.slice(0, 40).forEach(o => {
        if (o.op === 'editar') mostrar(`${o.linha_id}: ${o.antes.operacao} ${o.antes.ativo}  →  ${o.depois.operacao} ${o.depois.ativo}`);
        else if (o.op === 'inserir') mostrar(`+ ${o.linha_id}: ${o.depois.operacao} ${o.depois.ativo}`, 'text-[#3fb950]');
        else if (o.op === 'excluir') mostrar(`− ${o.linha_id}: ${o.antes.operacao} ${o.antes.ativo}`, 'text-red-400');
        else mostrar(`↕ ${AJ_TABELAS[o.tabela]}: as linhas seriam reordenadas`);
    });
    if (r.operacoes.length > 40) mostrar(`…e mais ${r.operacoes.length - 40} mudança(s).`, 'text-[#8b949e]');
    r.descartadas.forEach(d => mostrar(`Descartado (${d.linha_id}): ${d.motivo}`, 'text-[#d29922]'));
    r.avisos.forEach(a => mostrar('Atenção: ' + a, 'text-[#d29922]'));
}

/* ── cartão da receita ────────────────────────────────────────────────────── */
function ajRenderizar() {
    const lista = ajEl('ajLista');
    lista.replaceChildren();
    if (!AJ.itens.length) lista.appendChild(rdNo('div', 'text-sm text-[#8b949e]', 'Nenhum ajuste cadastrado.'));
    AJ.itens.forEach((item, i) => lista.appendChild(ajCartao(item, i)));
}

function ajCartao(item, i) {
    const leitura = rdLeitura();
    const herdada = ajEhProjeto() && item.origem !== 'projeto';
    const refazer = () => { ajRenderizar(); ajAgendarFrases(item); };
    const card = rdNo('div', 'border rounded-md p-3');
    card.dataset.ajuste = item.id;
    if (item.oculta) card.classList.add('opacity-50');
    card.addEventListener('input', () => { if (!leitura) { ajMarcarSujo(); ajAgendarFrases(item); } });
    card.addEventListener('change', () => { if (!leitura) ajMarcarSujo(); });

    const topo = rdNo('div', 'flex flex-wrap items-center gap-2');
    const ativa = document.createElement('input'); ativa.type = 'checkbox'; ativa.checked = !!item.ativa; ativa.title = 'Ajuste ligado';
    ativa.disabled = leitura; ativa.onchange = () => { item.ativa = ativa.checked; selo(); };
    const seloEl = rdNo('span', 'text-xs px-2 py-0.5 rounded-full');
    function selo() {
        const o = ajOrigemViva(item);
        seloEl.classList.toggle('hidden', !o);
        if (o) { seloEl.className = 'text-xs px-2 py-0.5 rounded-full ' + RD_SELOS[o][1]; seloEl.textContent = RD_SELOS[o][0]; }
    }
    selo();
    const nome = ajInput(item.nome, v => { item.nome = v; selo(); }, 'w-64');
    nome.readOnly = leitura; nome.placeholder = 'Nome do ajuste';
    const rid = ajInput(item.id, v => { item.id = v.trim(); card.dataset.ajuste = item.id; }, 'w-44 font-mono');
    rid.readOnly = herdada || leitura;
    topo.append(ativa, seloEl, nome, rid);

    const acoes = rdNo('span', 'ml-auto flex flex-wrap gap-2');
    if (!leitura) {
        if (herdada) {
            acoes.appendChild(rdBotao(item.oculta ? '+' : '−', () => { item.oculta = !item.oculta; ajMarcarSujo(); ajRenderizar(); },
                RD_CLS_BTN + ' font-bold w-9', item.oculta ? 'Reexibir este ajuste neste projeto' : 'Ocultar: desliga este ajuste só neste projeto (continua listado)'));
            if (AJ.padrao && AJ.padrao[item.id]) {
                acoes.appendChild(rdBotao('voltar ao padrão', () => {
                    const base = rdClone(AJ.padrao[item.id]);
                    Object.keys(item).forEach(k => { if (!['oculta', 'origem'].includes(k)) delete item[k]; });
                    Object.assign(item, base, { oculta: item.oculta, origem: 'padrao' });
                    ajMarcarSujo(); ajRenderizar();
                }, RD_CLS_BTN, 'Desfaz os ajustes desta receita neste projeto'));
            }
        }
        acoes.appendChild(rdBotao(item._aberta ? 'Fechar' : 'Editar', () => { item._aberta = !item._aberta; ajRenderizar(); }));
        if (!herdada) acoes.appendChild(rdBotao('Excluir', () => {
            if (confirm(`Excluir o ajuste ${item.id}?`)) { AJ.itens.splice(i, 1); ajMarcarSujo(); ajRenderizar(); }
        }, RD_CLS_BTN + ' text-red-400'));
    }
    topo.appendChild(acoes);
    card.appendChild(topo);

    const frases = rdNo('div', 'text-sm mt-2 space-y-1');
    (item.frases || []).forEach(f => frases.appendChild(rdNo('div', '', f)));
    item._frases = frases;
    card.appendChild(frases);
    if (item.descricao) card.appendChild(rdNo('div', 'text-xs text-[#8b949e] mt-1', item.descricao));
    if ((item.regras || []).length) card.appendChild(rdNo('div', 'text-xs text-[#8b949e] mt-1', 'Corrige: ' + item.regras.join(', ')));
    if (item.oculta) card.appendChild(rdNo('div', 'text-xs text-[#8b949e] mt-1', 'Oculto neste projeto. Use + para reexibir.'));

    const previa = rdNo('div', 'mt-2');
    const barra = rdNo('div', 'flex flex-wrap gap-2 mt-2');
    barra.appendChild(rdBotao('Pré-visualizar nas tabelas atuais', () => ajPreVisualizar(item, previa)));
    card.append(barra, previa);

    if (item._aberta && !leitura) card.appendChild(ajCorpoEditor(item, refazer));
    return card;
}

function ajCorpoEditor(item, refazer) {
    const corpo = rdNo('div', 'mt-3 pt-3 border-t space-y-2');
    const desc = ajInput(item.descricao || '', v => { item.descricao = v; }, 'w-full');
    desc.placeholder = 'Descrição (opcional)';
    const reg = ajInput((item.regras || []).join(', '), v => { const l = ajCsv(v); if (l.length) item.regras = l; else delete item.regras; }, 'w-full', 'ids das regras que este ajuste corrige, separados por vírgula', 'ajRegrasSugestoes');
    reg.placeholder = 'Corrige as regras (ex.: C2-CFU-SUPL) — o achado dessa regra oferecerá este ajuste';
    corpo.append(desc, reg);
    corpo.appendChild(rdBotao(item._avancado ? '← Voltar ao editor visual' : 'Modo avançado (JSON)', () => {
        if (item._jsonInvalido) { alert('Corrija o JSON antes de voltar ao editor visual.'); return; }
        item._avancado = !item._avancado; ajRenderizar();
    }));
    if (item._avancado) {
        const ta = rdNo('textarea', RD_CLS_INPUT + ' w-full font-mono text-xs'); ta.rows = 14; ta.spellcheck = false;
        ta.value = JSON.stringify(ajLimpa(item), null, 2);
        ta.oninput = () => {
            try {
                const novo = JSON.parse(ta.value);
                if (!novo || typeof novo !== 'object' || Array.isArray(novo)) throw new Error('objeto esperado');
                Object.keys(item).forEach(k => { if (!['origem', 'oculta', 'frases'].includes(k) && !k.startsWith('_')) delete item[k]; });
                Object.assign(item, novo);
                ta.classList.remove('border-red-500'); item._jsonInvalido = false;
            } catch (e) { ta.classList.add('border-red-500'); item._jsonInvalido = true; }
        };
        corpo.appendChild(ta);
        return corpo;
    }
    (item.acoes || []).forEach((a, i) => corpo.appendChild(ajAcaoEditor(item, a, i, refazer)));
    const novo = rdNo('div', 'flex flex-wrap gap-2 items-center');
    const sel = ajSelect(Object.entries(AJ_ACOES).map(([k, v]) => [k, v.rot]), 'substituir', () => { });
    novo.append(sel, rdBotao('+ ação', () => { item.acoes.push(AJ_ACOES[sel.value].novo()); refazer(); }));
    corpo.appendChild(novo);
    return corpo;
}

/* ── ações da barra ───────────────────────────────────────────────────────── */
function _ajPrefixoNovo() { return ajContexto() ? 'CTX' : (ajEhProjeto() ? 'PROJ' : 'AJ'); }

function ajNovo() {
    let n = AJ.itens.length + 1;
    while (AJ.itens.some(r => r.id === `AJ-NOVO-${n}`)) n++;
    AJ.itens.push({ id: `${_ajPrefixoNovo()}-NOVO-${n}`, nome: 'Novo ajuste', ativa: false, regras: [],
        acoes: [AJ_ACOES.normalizar.novo()], origem: ajEhProjeto() ? 'projeto' : undefined, _aberta: true, frases: [] });
    ajMarcarSujo();
    ajRenderizar();
    ajAtualizarFrases(AJ.itens[AJ.itens.length - 1]);
}

/** Cria um ajuste novo já com as ações dadas (ex.: proposta da IA aceita) — fica em rascunho até o admin salvar. */
function ajNovoComAcoes(acoes, nome) {
    let n = AJ.itens.length + 1;
    while (AJ.itens.some(r => r.id === `AJ-NOVO-${n}`)) n++;
    const item = { id: `${_ajPrefixoNovo()}-NOVO-${n}`, nome: nome || 'Novo ajuste', ativa: false, regras: [],
        acoes: rdClone(acoes), origem: ajEhProjeto() ? 'projeto' : undefined, _aberta: true, frases: [] };
    AJ.itens.push(item);
    ajMarcarSujo();
    ajRenderizar();
    ajAtualizarFrases(item);
    const card = ajEl('ajLista').lastElementChild;
    if (card) card.scrollIntoView({ block: 'center' });
}

function ajParaEnvio() {
    if (AJ.itens.some(r => r._jsonInvalido)) { ajErros(['Há ajustes com JSON inválido (campo em vermelho no modo avançado).']); return null; }
    return AJ.itens.map(r => { const c = ajLimpa(r); if (r.oculta) c.oculta = true; if (r.origem) c.origem = r.origem; return c; });
}

async function ajSalvar() {
    ajErros(null);
    const ajustes = ajParaEnvio();
    if (!ajustes) return;
    const res = await rdPost('/api/validacao/ajustes', { projeto_codigo: ajProjeto(), ajustes, contexto: ajContexto() });
    if (!res.ok) { rpMensagem('error', await ajLerErro(res)); return; }
    rpMensagem('success', 'Ajustes salvos.');
    await ajCarregar();
}

async function ajRestaurarSemente() {
    const msg = ajEhProjeto()
        ? 'Descartar todos os ajustes deste projeto e voltar ao padrão? A versão atual fica no histórico.'
        : 'Substituir os ajustes pelos da semente do sistema (todos desligados)? A versão atual fica no histórico.';
    if (!confirm(msg)) return;
    ajErros(null);
    const res = await rdPost('/api/validacao/ajustes/restaurar-semente', { projeto_codigo: ajProjeto() });
    if (!res.ok) { rpMensagem('error', await ajLerErro(res)); return; }
    rpMensagem('success', ajEhProjeto() ? 'Ajustes do projeto descartados.' : 'Semente restaurada.');
    await ajCarregar();
}

async function ajAlternarHistorico() {
    const box = ajEl('ajHistLista');
    if (!box.classList.contains('hidden')) { box.classList.add('hidden'); return; }
    const res = await fetch(`/api/validacao/ajustes/historico?projeto_codigo=${encodeURIComponent(ajProjeto())}`);
    if (!res.ok) { rpMensagem('error', await ajLerErro(res)); return; }
    const historico = (await res.json()).historico;
    box.replaceChildren();
    if (!historico.length) box.appendChild(rdNo('div', 'p-3 text-[#8b949e]', 'Sem versões anteriores para este projeto.'));
    historico.forEach(h => {
        const linha = rdNo('div', 'flex items-center justify-between gap-3 p-3');
        linha.appendChild(rdNo('span', '', `${h.criado_em} — ${h.criado_por || 'desconhecido'} — ${h.ajustes.length} ajustes, ${h.ajustes.filter(r => r.ativa).length} ligados`));
        const acoes = rdNo('span', 'space-x-3');
        const ver = rdNo('button', 'text-blue-400 hover:text-blue-300', 'Ver no editor');
        ver.onclick = () => { AJ.itens = rdClone(h.ajustes); ajMarcarSujo(); ajRenderizar(); };
        const rev = rdNo('button', 'text-[#d97706] hover:text-[#b45309]', 'Reverter');
        rev.onclick = async () => {
            if (!confirm('Reverter para esta versão? A versão atual fica no histórico.')) return;
            const r = await rdPost('/api/validacao/ajustes/reverter', { projeto_codigo: ajProjeto(), historico_id: h.id });
            if (!r.ok) { rpMensagem('error', await ajLerErro(r)); return; }
            rpMensagem('success', 'Versão revertida.');
            await ajCarregar();
        };
        acoes.append(ver, rev);
        linha.appendChild(acoes);
        box.appendChild(linha);
    });
    box.classList.remove('hidden');
}
