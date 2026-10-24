/*
 * modal_montavel_editor.js (TASK-077, etapa 2) — editor VISUAL de um modal montável (formulário por bloco).
 * Edita o MESMO objeto JSON do editor de texto: criar(container, modal) devolve { obter() } e o gerenciador alterna
 * entre "Visual" e "JSON" sem perder nada. Partes puras (blocoNovo, mover, idUnico, tsv) testadas em Node: tests/test_modal_montavel_editor.py.
 */
(function (raiz) {
    'use strict';

    const TIPOS_BLOCO = [['texto', 'Texto'], ['campo', 'Campo de entrada'], ['grade', 'Grade (colar linhas)'], ['tabela_interna', 'Planilha interna'],
        ['interpretar', 'Interpretar (regras)'], ['destino', 'Destino (Outros/Cabos)'], ['acao_adicionar', 'Botão Adicionar ativos'], ['saida_texto', 'Texto para copiar']];
    const TIPOS_MODAL = [['montagem', 'Montagem'], ['colagem', 'Colagem'], ['interpretativo', 'Interpretativo'], ['processamento', 'Processamento'], ['gerador_texto', 'Gerador de texto']];
    const ENTRADAS = [['texto', 'Texto'], ['numero', 'Número'], ['select', 'Lista'], ['radio', 'Opções']];
    const OPS_COND = [['igual', 'é igual a'], ['diferente', 'é diferente de'], ['vazio', 'está vazio'], ['preenchido', 'está preenchido'], ['em', 'está em (lista, uma por linha)']];

    /* ── partes puras ── */
    function idUnico(base, usados) {
        const b = String(base || 'item').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9_-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 30) || 'item';
        if (!usados.includes(b)) return b;
        let n = 2; while (usados.includes(b + '_' + n)) n++;
        return b + '_' + n;
    }
    function idsDoModal(modal) { return (modal.blocos || []).map(b => b.id).filter(Boolean); }

    /** Bloco novo com valores iniciais que já passam no validador do servidor. */
    function blocoNovo(tipo, modal) {
        const usados = idsDoModal(modal), grades = (modal.blocos || []).filter(b => b.tipo === 'grade'), interps = (modal.blocos || []).filter(b => b.tipo === 'interpretar');
        switch (tipo) {
            case 'texto': return { tipo, texto: 'Texto explicativo.' };
            case 'campo': return { tipo, id: idUnico('campo', usados), rotulo: 'Campo', entrada: 'texto' };
            case 'grade': return { tipo, id: idUnico('grade', usados), linhas_min: 5, colunas: [{ id: 'texto', rotulo: 'TEXTO' }] };
            case 'tabela_interna': return { tipo, id: idUnico('planilha', usados), colunas: ['codigo', 'titulo'], linhas: [] };
            case 'interpretar': return { tipo, id: idUnico('resultado', usados), fonte: grades.length ? grades[0].id : '', regras: [novaRegra()] };
            case 'destino': return { tipo, id: idUnico('destino', usados), opcoes: ['outros', 'cabos'], padrao: 'outros' };
            case 'acao_adicionar': return { tipo, rotulo: 'Adicionar ativos', fonte: interps.length ? interps[0].id : '', destino: 'outros' };
            case 'saida_texto': return { tipo, id: idUnico('texto', usados), rotulo: 'Texto para copiar', modelo: '{{campo}}' };
            default: return null;
        }
    }
    function condicaoNova(campo) { return { campo: campo || '', op: 'igual', valor: '' }; }
    function novaRegra() { return { quando: { regex: 'ABC' }, gera: [{ ativo: 'ATIVO', qtd: 1, operacao: 'I' }] }; }

    /** Move o item i para cima (d=-1) ou para baixo (d=1); devolve novo array (ou o mesmo se não der). */
    function mover(arr, i, d) {
        const j = i + d;
        if (i < 0 || j < 0 || i >= arr.length || j >= arr.length) return arr;
        const c = arr.slice(); const t = c[i]; c[i] = c[j]; c[j] = t; return c;
    }
    const linhasParaTsv = linhas => (linhas || []).map(l => l.map(v => (v == null ? '' : String(v))).join('\t')).join('\n');
    function tsvParaLinhas(texto, nCols) {
        return String(texto || '').replace(/\r\n?/g, '\n').split('\n').filter(l => l.trim() !== '').map(l => {
            const c = l.split('\t'); while (c.length < nCols) c.push(''); return c.slice(0, nCols).map(v => v.trim());
        });
    }

    /* ── DOM ── */
    const doc = raiz.document;
    function el(tag, props, filhos) {
        const e = doc.createElement(tag);
        Object.keys(props || {}).forEach(k => {
            if (k === 'style') e.style.cssText = props[k];
            else if (k === 'text') e.textContent = props[k];
            else if (k.indexOf('on') === 0) e.addEventListener(k.slice(2), props[k]);
            else e.setAttribute(k, props[k]);
        });
        (filhos || []).forEach(f => f && e.appendChild(typeof f === 'string' ? doc.createTextNode(f) : f));
        return e;
    }
    const IN = 'padding:5px 7px;background:#161b22;color:#e6edf3;border:1px solid #30363d;border-radius:6px;box-sizing:border-box;font-size:.8rem;';
    const mini = (txt, fn, extra) => el('button', { type: 'button', class: 'btn-secondary', text: txt, onclick: fn, style: 'padding:2px 8px;font-size:.72rem;' + (extra || '') });
    const rotulo = t => el('label', { text: t, style: 'display:block;font-size:.7rem;color:#8b949e;margin:6px 0 2px;' });

    function criar(container, modal) {
        const m = JSON.parse(JSON.stringify(modal));
        m.blocos = m.blocos || [];
        const topo = el('div'), lista = el('div'), novo = el('div');
        container.appendChild(topo); container.appendChild(lista); container.appendChild(novo);

        function campoTexto(obj, chave, titulo, opts) {
            opts = opts || {};
            const inp = opts.multi ? el('textarea', { rows: String(opts.multi), style: IN + 'width:100%;font-family:monospace;' }) : el('input', { type: opts.tipo || 'text', style: IN + 'width:100%;' });
            inp.value = obj[chave] == null ? '' : obj[chave];
            inp.addEventListener('input', () => { obj[chave] = opts.numero ? (inp.value === '' ? '' : Number(inp.value)) : inp.value; });
            if (opts.recarregar) inp.addEventListener('change', redesenhar);
            return el('div', {}, [rotulo(titulo), inp]);
        }
        function campoLista(obj, chave, titulo, opcoes, opts) {
            opts = opts || {};
            const s = el('select', { style: IN + 'width:100%;' }, (opts.vazio ? [el('option', { value: '', text: opts.vazio })] : []).concat(opcoes.map(o => el('option', { value: o[0], text: o[1] }))));
            s.value = obj[chave] == null ? '' : obj[chave];
            s.addEventListener('change', () => { if (s.value === '' && opts.vazio) delete obj[chave]; else obj[chave] = s.value; if (opts.recarregar) redesenhar(); });
            return el('div', {}, [rotulo(titulo), s]);
        }
        const idsDe = tipo => m.blocos.filter(b => b.tipo === tipo && b.id).map(b => [b.id, b.id]);

        function desenharTopo() {
            topo.innerHTML = '';
            topo.appendChild(el('div', { style: 'display:grid;grid-template-columns:2fr 1fr 1fr;gap:8px;' }, [
                campoTexto(m, 'nome', 'Nome do modal (aparece no botão)'),
                campoLista(m, 'tipo', 'Tipo', TIPOS_MODAL),
                campoTexto(m, 'id', 'Identificador (sem espaços)')]));
        }

        /* Editor de condição em árvore: folha (campo/op/valor), grupo E/OU e NÃO. Mexe no objeto de `obj[chave]`. */
        const OPS = OPS_COND.concat([['contem', 'contém o texto'], ['maior', 'é maior que (número)'], ['menor', 'é menor que (número)'], ['maior_igual', 'é maior ou igual a'], ['menor_igual', 'é menor ou igual a']]);
        function condicao(obj, chave, titulo) {
            const campos = idsDe('campo');
            const box = el('details', { style: 'margin-top:8px;' }, [el('summary', { text: titulo + (obj[chave] ? ' (ativa)' : ''), style: 'cursor:pointer;font-size:.72rem;color:#8b949e;' })]);
            if (!campos.length) { box.appendChild(el('div', { text: 'Crie um bloco "Campo de entrada" para usar condições.', style: 'font-size:.72rem;color:#8b949e;margin-top:4px;' })); return box; }
            const holder = { c: obj[chave] || condicaoNova(campos[0][0]) };
            const ativa = el('input', { type: 'checkbox' }); ativa.checked = !!obj[chave];
            const area = el('div', { style: 'margin-top:6px;' });
            const sync = () => { if (ativa.checked) obj[chave] = holder.c; else delete obj[chave]; };
            ativa.addEventListener('change', sync);
            function no(node, trocar, nivel) {
                const caixa = el('div', { style: 'border-left:2px solid #30363d;padding:4px 0 4px 8px;margin:4px 0;' });
                const tipoNo = node.nao ? 'nao' : node.todos ? 'todos' : node.algum ? 'algum' : 'folha';
                if (tipoNo === 'nao') {
                    caixa.appendChild(el('div', { style: 'display:flex;gap:6px;align-items:center;' }, [el('b', { text: 'NÃO', style: 'font-size:.72rem;color:#f0883e;' }),
                        mini('tirar NÃO', () => { trocar(node.nao); sync(); redesenhar(); })]));
                    caixa.appendChild(no(node.nao, novo => { node.nao = novo; sync(); }, nivel + 1));
                } else if (tipoNo === 'folha') {
                    const lin = el('div', { style: 'display:grid;grid-template-columns:1fr 1.3fr 1fr auto auto;gap:6px;align-items:end;' });
                    const sc = el('select', { style: IN }, campos.map(c => el('option', { value: c[0], text: c[1] }))); sc.value = node.campo;
                    sc.addEventListener('change', () => { node.campo = sc.value; sync(); });
                    const so = el('select', { style: IN }, OPS.map(o => el('option', { value: o[0], text: o[1] }))); so.value = node.op;
                    const val = el('input', { type: 'text', style: IN + 'width:100%;', placeholder: node.op === 'em' ? 'a, b, c' : 'valor' });
                    const setVal = () => { if (['maior', 'menor', 'maior_igual', 'menor_igual'].includes(node.op)) node.valor = val.value === '' ? 0 : Number(val.value) || 0; else if (node.op === 'em') node.valor = val.value.split(',').map(x => x.trim()).filter(Boolean); else node.valor = val.value; };
                    val.value = Array.isArray(node.valor) ? node.valor.join(', ') : (node.valor == null ? '' : node.valor);
                    val.addEventListener('input', () => { setVal(); sync(); });
                    so.addEventListener('change', () => { node.op = so.value; setVal(); val.style.display = ['vazio', 'preenchido'].includes(node.op) ? 'none' : ''; sync(); });
                    val.style.display = ['vazio', 'preenchido'].includes(node.op) ? 'none' : '';
                    lin.appendChild(sc); lin.appendChild(so); lin.appendChild(val);
                    lin.appendChild(mini('NÃO', () => { trocar({ nao: node }); sync(); redesenhar(); }));
                    lin.appendChild(mini('✕', () => { trocar(null); sync(); redesenhar(); }, 'color:#f85149;'));
                    caixa.appendChild(lin);
                } else {
                    const lista = node[tipoNo];
                    const sel = el('select', { style: IN }, [el('option', { value: 'todos', text: 'TODAS verdadeiras (E)' }), el('option', { value: 'algum', text: 'ALGUMA verdadeira (OU)' })]); sel.value = tipoNo;
                    sel.addEventListener('change', () => { const nv = {}; nv[sel.value] = lista; trocar(nv); sync(); redesenhar(); });
                    caixa.appendChild(el('div', { style: 'display:flex;gap:6px;align-items:center;' }, [sel, mini('NÃO', () => { trocar({ nao: node }); sync(); redesenhar(); }),
                        nivel > 0 ? mini('✕ grupo', () => { trocar(null); sync(); redesenhar(); }, 'color:#f85149;') : null]));
                    lista.forEach((f, i) => caixa.appendChild(no(f, novo => { if (novo === null) lista.splice(i, 1); else lista[i] = novo; if (!lista.length) lista.push(condicaoNova(campos[0][0])); }, nivel + 1)));
                    caixa.appendChild(el('div', { style: 'display:flex;gap:6px;margin-top:4px;' }, [mini('+ condição', () => { lista.push(condicaoNova(campos[0][0])); sync(); redesenhar(); }),
                        nivel < 3 ? mini('+ grupo', () => { lista.push({ algum: [condicaoNova(campos[0][0])] }); sync(); redesenhar(); }) : null]));
                }
                return caixa;
            }
            function redesenhar() { area.innerHTML = ''; area.appendChild(no(holder.c, novo => { holder.c = novo || condicaoNova(campos[0][0]); }, 0)); }
            redesenhar();
            box.appendChild(el('label', { style: 'display:flex;gap:6px;align-items:center;font-size:.78rem;margin-top:4px;' }, [ativa, 'Usar condição']));
            box.appendChild(area);
            // 'redesenhar' do formulário inteiro é outro; este só refaz a árvore
            return box;
        }

        function formulario(b) {
            const f = el('div');
            switch (b.tipo) {
                case 'texto': f.appendChild(campoTexto(b, 'texto', 'Texto', { multi: 3 })); break;
                case 'campo':
                    f.appendChild(el('div', { style: 'display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;' }, [
                        campoTexto(b, 'id', 'Identificador', { recarregar: true }), campoTexto(b, 'rotulo', 'Rótulo'), campoLista(b, 'entrada', 'Entrada', ENTRADAS, { recarregar: true })]));
                    if (b.entrada === 'select' || b.entrada === 'radio') {
                        const ta = el('textarea', { rows: '3', style: IN + 'width:100%;' }); ta.value = (b.opcoes || []).join('\n');
                        ta.addEventListener('input', () => { b.opcoes = ta.value.split('\n').map(x => x.trim()).filter(Boolean); });
                        f.appendChild(rotulo('Opções (uma por linha)')); f.appendChild(ta);
                    }
                    break;
                case 'grade': {
                    f.appendChild(el('div', { style: 'display:grid;grid-template-columns:1fr 1fr;gap:8px;' }, [campoTexto(b, 'id', 'Identificador', { recarregar: true }), campoTexto(b, 'linhas_min', 'Linhas iniciais', { numero: true, tipo: 'number' })]));
                    f.appendChild(rotulo('Colunas (na ordem em que o Excel cola)'));
                    b.colunas.forEach((c, i) => f.appendChild(el('div', { style: 'display:flex;gap:6px;margin-bottom:3px;' }, [
                        el('div', { style: 'flex:1;' }, [campoTexto(c, 'id', 'id', { recarregar: true })]), el('div', { style: 'flex:1;' }, [campoTexto(c, 'rotulo', 'título')]),
                        mini('✕', () => { b.colunas.splice(i, 1); redesenhar(); }, 'align-self:flex-end;color:#f85149;')])));
                    f.appendChild(mini('+ Coluna', () => { b.colunas.push({ id: idUnico('col', b.colunas.map(c => c.id)), rotulo: 'COLUNA' }); redesenhar(); }));
                    const col = b.colagem || (b.colagem = { modo: 'colunas' });
                    f.appendChild(el('div', { style: 'display:grid;grid-template-columns:1.4fr 1fr 1fr 1.2fr;gap:8px;align-items:end;margin-top:8px;' }, [
                        campoLista(col, 'modo', 'Ao colar', [['colunas', 'Colunas do Excel (Tab)'], ['linha', 'Linha inteira na 1ª coluna'], ['dividir', 'Dividir a linha por um separador']], { recarregar: true }),
                        col.modo === 'dividir' ? campoTexto(col, 'separador', 'Separador (ex.: ; ou ,)') : el('div'),
                        (() => { const c = el('input', { type: 'checkbox' }); c.checked = !!col.pular_cabecalho; c.addEventListener('change', () => { col.pular_cabecalho = c.checked; });
                            return el('label', { style: 'display:flex;gap:5px;align-items:center;font-size:.75rem;' }, [c, 'Ignorar a 1ª linha colada']); })(),
                        campoLista(b, 'carregar_de', 'Botão "Carregar da tabela"', [['outros', 'Tabela Outros'], ['cabos', 'Tabela Cabos']], { vazio: '(sem botão)' })]));
                    break;
                }
                case 'tabela_interna': {
                    f.appendChild(campoTexto(b, 'id', 'Identificador', { recarregar: true }));
                    const cols = el('input', { type: 'text', style: IN + 'width:100%;' }); cols.value = b.colunas.join(', ');
                    cols.addEventListener('change', () => { b.colunas = cols.value.split(',').map(x => x.trim()).filter(Boolean); b.linhas = tsvParaLinhas(linhasParaTsv(b.linhas), b.colunas.length); redesenhar(); });
                    f.appendChild(rotulo('Colunas (separadas por vírgula)')); f.appendChild(cols);
                    const ta = el('textarea', { rows: '6', style: IN + 'width:100%;font-family:monospace;', placeholder: 'Cole do Excel: uma linha por linha, colunas separadas por Tab' });
                    ta.value = linhasParaTsv(b.linhas);
                    ta.addEventListener('input', () => { b.linhas = tsvParaLinhas(ta.value, b.colunas.length); });
                    f.appendChild(rotulo('Linhas (cole do Excel)')); f.appendChild(ta);
                    break;
                }
                case 'interpretar': {
                    f.appendChild(el('div', { style: 'display:grid;grid-template-columns:1fr 1fr;gap:8px;' }, [campoTexto(b, 'id', 'Identificador', { recarregar: true }),
                        campoLista(b, 'fonte', 'Ler a grade', idsDe('grade'), { vazio: '— escolha —', recarregar: true })]));
                    f.appendChild(campoLista(b, 'normalizar', 'Letras do ativo gerado', [['maiusculo', 'MAIÚSCULAS'], ['minusculo', 'minúsculas']], { vazio: '(como escrito na regra)' }));
                    const grade = m.blocos.find(x => x.tipo === 'grade' && x.id === b.fonte);
                    const colunas = grade ? grade.colunas.map(c => [c.id, c.rotulo]) : [];
                    b.regras.forEach((r, ri) => {
                        const cx = el('div', { style: 'border:1px solid #30363d;border-radius:6px;padding:8px;margin-top:8px;' });
                        cx.appendChild(el('div', { style: 'display:flex;justify-content:space-between;align-items:center;' }, [el('b', { text: 'Regra ' + (ri + 1), style: 'font-size:.78rem;' }),
                            el('div', { style: 'display:flex;gap:4px;' }, [mini('▲', () => { b.regras = mover(b.regras, ri, -1); redesenhar(); }), mini('▼', () => { b.regras = mover(b.regras, ri, 1); redesenhar(); }),
                                mini('✕', () => { b.regras.splice(ri, 1); redesenhar(); }, 'color:#f85149;')])]));
                        cx.appendChild(el('div', { style: 'display:grid;grid-template-columns:1fr 2fr;gap:8px;' }, [
                            campoLista(r.quando, 'coluna', 'Procurar na coluna', colunas, { vazio: '(todas as colunas)' }), campoTexto(r.quando, 'regex', 'Se o texto casar com a regex (use ( ) para capturar números: {1}, {2}…)')]));
                        const parar = el('input', { type: 'checkbox' }); parar.checked = r.parar !== false;
                        parar.addEventListener('change', () => { if (parar.checked) delete r.parar; else r.parar = false; });
                        cx.appendChild(el('label', { style: 'display:flex;gap:6px;align-items:center;font-size:.75rem;margin-top:4px;' }, [parar, 'Parar nesta regra (não testar as seguintes)']));
                        const glob = el('input', { type: 'checkbox' }); glob.checked = !!r.quando.global;
                        glob.addEventListener('change', () => { if (glob.checked) r.quando.global = true; else delete r.quando.global; });
                        cx.appendChild(el('label', { style: 'display:flex;gap:6px;align-items:center;font-size:.75rem;margin-top:2px;' }, [glob, 'Todas as ocorrências na linha (ex.: "P1, P2 e P3" gera 3 itens)']));
                        cx.appendChild(condicao(r, 'se', 'Aplicar esta regra só se (campos do modal)…'));
                        cx.appendChild(rotulo('Então gera (use [campo] na quantidade ou no ativo)'));
                        r.gera.forEach((g, gi) => cx.appendChild(el('div', { style: 'display:grid;grid-template-columns:2fr 1.4fr 0.8fr auto;gap:6px;align-items:end;margin-bottom:3px;' }, [
                            campoTexto(g, 'ativo', 'ativo (ex.: MAC ou P{1})'), campoTexto(g, 'qtd', 'quantidade (ex.: {1}*20)'),
                            campoLista(g, 'operacao', 'operação', ['I', '*I', 'R', '*R', 'M', '*M'].map(x => [x, x])), mini('✕', () => { r.gera.splice(gi, 1); redesenhar(); }, 'color:#f85149;')])));
                        cx.appendChild(mini('+ Item gerado', () => { r.gera.push({ ativo: 'ATIVO', qtd: 1, operacao: 'I' }); redesenhar(); }));
                        f.appendChild(cx);
                    });
                    f.appendChild(el('div', { style: 'margin-top:8px;' }, [mini('+ Regra', () => { b.regras.push(novaRegra()); redesenhar(); })]));
                    break;
                }
                case 'destino': {
                    f.appendChild(campoTexto(b, 'id', 'Identificador', { recarregar: true }));
                    f.appendChild(campoLista(b, 'padrao', 'Destino padrão', [['outros', 'Tabela Outros'], ['cabos', 'Tabela Cabos']]));
                    break;
                }
                case 'acao_adicionar':
                    f.appendChild(el('div', { style: 'display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;' }, [campoTexto(b, 'rotulo', 'Texto do botão'),
                        campoLista(b, 'fonte', 'Adicionar o resultado de', idsDe('interpretar'), { vazio: '— escolha —' }),
                        campoLista(b, 'destino', 'Adicionar em', [['outros', 'Tabela Outros (fixo)'], ['cabos', 'Tabela Cabos (fixo)']].concat(idsDe('destino').map(d => [d[0], 'Escolhido no bloco "' + d[0] + '"'])))]));
                    break;
                case 'saida_texto':
                    f.appendChild(el('div', { style: 'display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;' }, [campoTexto(b, 'id', 'Identificador', { recarregar: true }), campoTexto(b, 'rotulo', 'Título'),
                        campoLista(b, 'fonte', 'Repetir para cada linha da grade', idsDe('grade'), { vazio: '(uma vez só)' })]));
                    f.appendChild(campoTexto(b, 'modelo', 'Modelo — {{coluna}}, {{campo}}, {{buscar(planilha, colunaChave, valor, colunaRetorno)}}', { multi: 3 }));
                    break;
                default: break;
            }
            f.appendChild(condicao(b, 'visivel_se', 'Mostrar este bloco só se…'));
            return f;
        }

        function desenharBlocos() {
            lista.innerHTML = '';
            m.blocos.forEach((b, i) => {
                const nome = (TIPOS_BLOCO.find(t => t[0] === b.tipo) || [0, b.tipo])[1];
                lista.appendChild(el('div', { class: 'mm-bloco', style: 'border:1px solid #30363d;border-radius:8px;padding:8px 10px;margin-top:10px;background:#0d1117;' }, [
                    el('div', { style: 'display:flex;justify-content:space-between;align-items:center;' }, [
                        el('b', { text: (i + 1) + '. ' + nome + (b.id ? ' — ' + b.id : ''), style: 'font-size:.8rem;color:#58a6ff;' }),
                        el('div', { style: 'display:flex;gap:4px;' }, [mini('▲', () => { m.blocos = mover(m.blocos, i, -1); redesenhar(); }), mini('▼', () => { m.blocos = mover(m.blocos, i, 1); redesenhar(); }),
                            mini('✕', () => { m.blocos.splice(i, 1); redesenhar(); }, 'color:#f85149;')])]),
                    formulario(b)]));
            });
            novo.innerHTML = '';
            novo.appendChild(el('div', { style: 'margin-top:12px;font-size:.72rem;color:#8b949e;', text: 'Adicionar bloco:' }));
            novo.appendChild(el('div', { style: 'display:flex;gap:6px;flex-wrap:wrap;margin-top:4px;' }, TIPOS_BLOCO.map(t => mini('+ ' + t[1], () => { m.blocos.push(blocoNovo(t[0], m)); redesenhar(); }))));
        }
        function redesenhar() { desenharTopo(); desenharBlocos(); }
        redesenhar();
        return { obter: () => JSON.parse(JSON.stringify(m)) };
    }

    const api = { criar, blocoNovo, novaRegra, condicaoNova, mover, idUnico, linhasParaTsv, tsvParaLinhas, TIPOS_BLOCO, TIPOS_MODAL };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.ModalMontavelEditor = api;
})(typeof window !== 'undefined' ? window : globalThis);
