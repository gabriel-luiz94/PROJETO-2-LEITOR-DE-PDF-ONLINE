/*
 * ativo_dica.js (TASK-122) — ajuda ao digitar "(" na coluna Ativo da tabela Outros: mostra os PARÂMETROS aceitos do ativo
 * (na ordem da declaração, com padrão e valores cadastrados), ou os avisos "ativo sem parâmetros declarados" / "ativo não encontrado".
 * Parte pura (sem DOM; testada em Node: tests/test_ativo_dica_122.py): contextoDoParentese, construirContexto, dica.
 * Parte de tela: iniciar() liga os eventos (delegados) nos campos de Outros e desenha a caixa. Usa a base já carregada na tela (funciona offline).
 */
(function (raiz) {
    'use strict';
    const AP = typeof require !== 'undefined' && typeof module !== 'undefined' ? require('./ativo_parametros.js') : raiz.AtivoParametros;

    /** Texto até o cursor → o parêntese ainda aberto: {nome, indice (posição do parâmetro), nomeParam, parcial, inicioParcial} ou null. */
    function contextoDoParentese(texto, cursor) {
        const t = String(texto == null ? '' : texto).slice(0, cursor == null ? String(texto).length : cursor);
        const pilha = [];
        let colchete = false;
        for (let i = 0; i < t.length; i++) {
            const c = t[i];
            if (c === '[') colchete = true;
            else if (c === ']') colchete = false;
            else if (!colchete && c === '(') pilha.push(i);
            else if (!colchete && c === ')') pilha.pop();
        }
        if (!pilha.length) return null;
        const abre = pilha[pilha.length - 1];
        const m = /(\[[^\]]*\]|[^\s,(]+)$/.exec(t.slice(0, abre));
        if (!m) return { nome: '', indice: 0, nomeParam: null, parcial: '', inicioParcial: t.length };
        let nome = m[1].replace(/^\*?\d+(?:[.,]\d+)?-(?=\S)/, '');                 // "1-IT_2" → "IT_2"; "[CAZ7,94]" fica literal
        if (/^\[[^\]]*\]$/.test(nome)) nome = nome.slice(1, -1);
        const dentro = t.slice(abre + 1);
        const partes = dentro.split(',');
        let seg = partes[partes.length - 1];
        let nomeParam = null, parcial = seg;
        const eq = seg.indexOf('=');
        if (eq !== -1) { nomeParam = seg.slice(0, eq).trim().toUpperCase() || null; parcial = seg.slice(eq + 1); }
        const semEspaco = parcial.replace(/^\s+/, '');
        return { nome: nome.trim().toUpperCase(), indice: partes.length - 1, nomeParam, parcial: semEspaco, inicioParcial: t.length - semEspaco.length };
    }

    /** Contexto da base para a dica: assinaturas, apelidos e nomes que o cálculo resolve para o projeto e a origem OUTROS (ativo/apelido/componente/código). */
    function construirContexto(linhas, projeto) {
        const p = String(projeto || '').trim().toUpperCase();
        const nomes = new Set();
        (linhas || []).forEach(r => {
            const lp = String(r.projeto || '').trim().toUpperCase();
            if (lp && p && !lp.split('/').map(x => x.trim()).includes(p)) return;
            const ro = String(r.origem || '').trim().toUpperCase();
            if (ro && ro !== 'OUTROS') return;
            const cel = String(r.ativo || '').trim();
            if (cel) { const a = AP.analisarCelulaBase(cel); if (!a.erro && a.base) nomes.add(a.base.toUpperCase()); }
            const comp = String(r.componente || '').trim().toUpperCase(), cod = String(r.codigo || '').trim().toUpperCase();
            if (comp) nomes.add(comp);
            if (cod) nomes.add(cod);
        });
        const apelidos = {};
        Object.entries(AP.apelidosDoProjeto(linhas || [], projeto || '') || {}).forEach(([k, v]) => { apelidos[String(k).toUpperCase()] = String(v).toUpperCase(); nomes.add(String(k).toUpperCase()); });
        return { assinaturas: AP.assinaturasDoProjeto(linhas || [], projeto || ''), apelidos, nomes };
    }

    /** O que mostrar para o nome antes do "(": {tipo: 'parametros'|'sem_parametros'|'nao_encontrado', ...}. */
    function dica(nome, ctx) {
        const n = String(nome || '').trim().toUpperCase();
        const base = (ctx.apelidos && ctx.apelidos[n]) || n;
        const sig = (ctx.assinaturas || {})[base];
        if (sig && ((sig.nomes && sig.nomes.length) || (sig.livres && sig.livres.length))) {
            const params = (sig.nomes || []).map(nm => {
                const valores = Array.from(new Set(((sig.valores || {})[nm] || []))), mascaras = ((sig.mascaras || {})[nm] || []).slice();
                mascaras.forEach(m => {                                                // TASK-149: máscara "%S336%" também vira o valor selecionável "S336" (um miolo só, sem outro % nem !)
                    const miolo = /^!/.test(m) ? null : String(m).replace(/^%+|%+$/g, '');
                    if (miolo && miolo.indexOf('%') === -1 && miolo.indexOf('|') === -1 && !valores.includes(miolo)) valores.push(miolo);
                });
                return { nome: nm, padrao: (sig.padroes || {})[nm] == null ? null : (sig.padroes || {})[nm], valores, mascaras, aberto: !(sig.fechados || []).includes(nm), porNome: false, obra: /^OBRA\./i.test(nm) };           // TASK-165: OBRA.* = da obra, não editável no pedido
            });
            (sig.livres || []).forEach(nm => params.push({ nome: nm, padrao: null, valores: [], mascaras: [], aberto: true, porNome: true }));
            return { tipo: 'parametros', ativo: base, params };
        }
        if (ctx.nomes && (ctx.nomes.has(n) || ctx.nomes.has(base))) return { tipo: 'sem_parametros', ativo: base, mensagem: `${base}: ativo sem parâmetros declarados` };
        return { tipo: 'nao_encontrado', ativo: n, mensagem: `${n}: ativo não encontrado na base` };
    }

    /** TASK-129: fecha os parênteses que faltam (N1(D, X → N1(D, X)); vírgula/espaço pendentes no fim saem antes; [nome literal] não conta; colchete aberto = texto malformado, não mexe. */
    function completarParenteses(texto) {
        const t = String(texto == null ? '' : texto);
        let abertos = 0, colchete = false;
        for (const c of t) {
            if (colchete) { if (c === ']') colchete = false; continue; }
            if (c === '[') colchete = true;
            else if (c === '(') abertos++;
            else if (c === ')' && abertos > 0) abertos--;
        }
        if (!abertos || colchete) return t;
        return t.replace(/[,\s]+$/, '') + ')'.repeat(abertos);
    }

    /** TASK-149: os parâmetros POSICIONAIS (na ordem) e o índice do parâmetro da vez; `atual` = o parâmetro (por posição ou pelo nome já digitado) ou null. */
    function parametroAtual(c, d) {
        if (!d || d.tipo !== 'parametros') return { pos: [], i: -1, atual: null };
        const pos = d.params.filter(p => !p.porNome && !p.obra);
        let i = c.nomeParam ? pos.findIndex(p => p.nome.toUpperCase() === c.nomeParam) : c.indice;
        if (i < 0 || i >= pos.length) i = -1;
        return { pos, i, atual: i === -1 ? null : pos[i] };
    }

    /** Valores selecionáveis (cadastrados) do parâmetro da vez. */
    function valoresDoAtual(c, d) { const a = parametroAtual(c, d).atual; return a ? a.valores.slice() : []; }

    /**
     * Inserir um valor do parâmetro da vez e AVANÇAR: troca o texto do parâmetro (do início do valor até a próxima "," ou ")") por `valor` e acrescenta ", " — ou ")" se for o
     * último parâmetro. Se a "," / ")" já existe logo depois, só pula por cima dela. Devolve {inicio, fim, inserir, cursor}: substituir [inicio, fim) por `inserir` e pôr o cursor em `cursor`.
     */
    function planoDeInsercao(texto, cursor, c, d, valor) {
        const t = String(texto == null ? '' : texto), { pos, i } = parametroAtual(c, d);
        const ult = i === -1 || i >= pos.length - 1;
        let fim = cursor, prof = 0;
        for (let k = cursor; k < t.length; k++) {                       // fim do parâmetro: próxima "," ou ")" do mesmo nível
            const ch = t[k];
            if (ch === '(') prof++;
            else if (ch === ')') { if (prof === 0) { fim = k; break; } prof--; }
            else if (ch === ',' && prof === 0) { fim = k; break; }
            fim = k + 1;
        }
        const resto = t.slice(fim);
        const base = { inicio: c.inicioParcial, fim, inserir: String(valor) };
        if (!ult) {
            if (resto[0] === ',') { const extra = resto[1] === ' ' ? 2 : 1; return Object.assign(base, { cursor: c.inicioParcial + base.inserir.length + extra }); }
            return Object.assign(base, { inserir: base.inserir + ', ', cursor: c.inicioParcial + base.inserir.length + 2 });
        }
        if (resto[0] === ')') return Object.assign(base, { cursor: c.inicioParcial + base.inserir.length + 1 });
        return Object.assign(base, { inserir: base.inserir + ')', cursor: c.inicioParcial + base.inserir.length + 1 });
    }

    /* ───────────── TASK-181: o valor entra PELO NOME (PR(cabo2=S04)); ↑↓ trocam o parâmetro, ←→ o valor ───────────── */
    const ehObraNome = n => /^OBRA\./i.test(String(n || ''));

    /** O par de parênteses (o mais interno) que contém o cursor, com o nome do ativo colado antes: {abre, fecha (-1 se ainda aberto), nome} ou null. */
    function grupoDoCursor(texto, cursor) {
        const t = String(texto == null ? '' : texto), cur = cursor == null ? t.length : cursor;
        const pilha = [], grupos = [];
        let colchete = false;
        for (let i = 0; i < t.length; i++) {
            const c = t[i];
            if (c === '[') colchete = true;
            else if (c === ']') colchete = false;
            else if (!colchete && c === '(') pilha.push(i);
            else if (!colchete && c === ')' && pilha.length) grupos.push({ abre: pilha.pop(), fecha: i });
        }
        pilha.forEach(a => grupos.push({ abre: a, fecha: -1 }));
        const dentro = grupos.filter(g => g.abre < cur && (g.fecha === -1 || cur <= g.fecha)).sort((a, b) => b.abre - a.abre)[0];
        if (!dentro) return null;
        const m = /(\[[^\]]*\]|[^\s,(]+)$/.exec(t.slice(0, dentro.abre));
        if (!m) return null;                                                          // "(" sem nome colado antes: não é parâmetro de ativo
        let nome = m[1].replace(/^\*?\d+(?:[.,]\d+)?-(?=\S)/, '');
        if (/^\[[^\]]*\]$/.test(nome)) nome = nome.slice(1, -1);
        return { abre: dentro.abre, fecha: dentro.fecha, nome: nome.trim().toUpperCase() };
    }

    /** Divide o conteúdo dos parênteses em partes no 1º nível (vírgula), com a posição de cada uma: [{ini, fim, texto, nome|null, valor}] (ini/fim relativos ao conteúdo). */
    function partesDoConteudo(inner) {
        const partes = [];
        let prof = 0, colchete = false, ini = 0;
        const fecha = fim => {
            const bruto = inner.slice(ini, fim), texto = bruto.trim();
            const eq = texto.indexOf('=');
            const nomeOk = eq > 0 && /^[A-Za-z_][\w.]*$/.test(texto.slice(0, eq).trim());
            partes.push({ ini, fim, texto, nome: nomeOk ? texto.slice(0, eq).trim().toUpperCase() : null, valor: nomeOk ? texto.slice(eq + 1).trim() : texto });
        };
        for (let i = 0; i < inner.length; i++) {
            const c = inner[i];
            if (c === '[') colchete = true; else if (c === ']') colchete = false;
            else if (!colchete && c === '(') prof++; else if (!colchete && c === ')') prof--;
            else if (!colchete && prof === 0 && c === ',') { fecha(i); ini = i + 1; }
        }
        fecha(inner.length);
        return partes;
    }

    /** Lê o pedido como o cálculo (resolver_pedido): valor posicional → o próximo parâmetro ainda sem valor (pula OBRA.*). Devolve {valores: {NOME: valor}, extras: [{nome, valor}], partes: [... + alvo]}.
     *  `nomes` = nomes formais na ordem da declaração. Uma parte vazia (cursor depois da última vírgula) fica com `alvo` = o próximo livre, para saber onde o cursor está. */
    function lerPedido(inner, nomes) {
        const formais = (nomes || []).map(n => String(n).toUpperCase());
        const partes = partesDoConteudo(inner), valores = {}, extras = [];
        let pos = 0;
        const proximo = () => { while (pos < formais.length && (formais[pos] in valores || ehObraNome(formais[pos]))) pos++; return pos < formais.length ? formais[pos] : null; };
        partes.forEach(pt => {
            if (pt.nome === null) {
                const alvo = proximo();
                pt.alvo = alvo;
                if (pt.texto === '') return;                                          // parte vazia: só marca onde está
                if (alvo) { valores[alvo] = pt.valor; pos++; } else extras.push({ nome: null, valor: pt.valor });
            } else {
                pt.alvo = pt.nome;
                if (formais.includes(pt.nome)) valores[pt.nome] = pt.valor; else extras.push({ nome: pt.nome, valor: pt.valor });
            }
        });
        return { valores, extras, partes };
    }

    /** Os parâmetros que a caixa deixa escolher (na ordem da declaração; OBRA.* fica de fora — vem do cabeçalho da obra). */
    function parametrosEditaveis(d) { return d && d.tipo === 'parametros' ? d.params.filter(p => !p.obra) : []; }

    /** Qual parâmetro está sob o cursor (nome em maiúsculas) ou null. Usa a mesma regra de posição do cálculo. */
    function parametroSobOCursor(texto, cursor, d) {
        const g = grupoDoCursor(texto, cursor);
        if (!g || !d || d.tipo !== 'parametros') return null;
        const t = String(texto), fimConteudo = g.fecha === -1 ? t.length : g.fecha, inner = t.slice(g.abre + 1, fimConteudo);
        const r = lerPedido(inner, d.params.map(p => p.nome)), rel = Math.max(0, Math.min(cursor - (g.abre + 1), inner.length));
        const pt = r.partes.find(x => rel >= x.ini && rel <= x.fim) || r.partes[r.partes.length - 1];
        return pt ? (pt.alvo || null) : null;
    }

    /** Os valores já escritos no parêntese do cursor: {NOME: valor}. */
    function valoresEscritos(texto, cursor, d) {
        const g = grupoDoCursor(texto, cursor);
        if (!g || !d || d.tipo !== 'parametros') return {};
        const t = String(texto);
        return lerPedido(t.slice(g.abre + 1, g.fecha === -1 ? t.length : g.fecha), d.params.map(p => p.nome)).valores;
    }

    /**
     * Grava (ou tira, com valor vazio/null) o valor de UM parâmetro, sempre pelo nome, reescrevendo só o conteúdo dos parênteses do cursor: PR(cabo2=S04), CE_4(poste=12600D,cabo=A35P,cabo2=S04).
     * Quem estava por posição passa a ficar por nome (para não misturar). Os de fora (OBRA.*, só-fórmula, desconhecidos) ficam como estavam, depois dos formais.
     * Devolve {texto, cursor (fim do valor gravado; no fim do conteúdo se o parâmetro foi removido), fecha} ou null se o cursor não está num parêntese.
     */
    function aplicarValor(texto, cursor, d, nomeParam, valor) {
        const g = grupoDoCursor(texto, cursor);
        if (!g || !d || d.tipo !== 'parametros') return null;
        const t = String(texto), fimConteudo = g.fecha === -1 ? t.length : g.fecha;
        const r = lerPedido(t.slice(g.abre + 1, fimConteudo), d.params.map(p => p.nome));
        const alvo = String(nomeParam).toUpperCase(), v = valor == null ? '' : String(valor).trim();
        if (v === '') delete r.valores[alvo]; else r.valores[alvo] = v;
        const partes = [];
        let fimValor = -1;
        d.params.forEach(p => {
            const n = p.nome.toUpperCase();
            if (n in r.valores) {
                partes.push(`${p.nome.toLowerCase()}=${r.valores[n]}`);
                if (n === alvo) fimValor = partes.join(',').length;
            }
        });
        r.extras.forEach(x => partes.push(x.nome ? `${x.nome.toLowerCase()}=${x.valor}` : x.valor));
        const inner = partes.join(',');
        const novo = t.slice(0, g.abre + 1) + inner + ')' + t.slice(g.fecha === -1 ? t.length : g.fecha + 1);
        const base = g.abre + 1;
        return { texto: novo, cursor: base + (fimValor === -1 ? inner.length : fimValor), fecha: base + inner.length };
    }

    /** "(" recém-digitado logo depois do nome: devolve {texto, cursor} com o ")" já escrito e o cursor no meio — ou null se o ")" já está ali. */
    function fecharParenteseRecemAberto(texto, cursor) {
        const t = String(texto == null ? '' : texto);
        if (cursor <= 0 || t[cursor - 1] !== '(' || t[cursor] === ')') return null;
        return { texto: t.slice(0, cursor) + ')' + t.slice(cursor), cursor };
    }

    /** Parênteses vazios colados a um nome viram só o nome (PR() → PR). */
    function limparParentesesVazios(texto) {
        return String(texto == null ? '' : texto).replace(/([^\s,(])\(\s*\)/g, '$1');
    }

    const api = { contextoDoParentese, construirContexto, dica, completarParenteses, parametroAtual, valoresDoAtual, planoDeInsercao,
        grupoDoCursor, partesDoConteudo, lerPedido, parametrosEditaveis, parametroSobOCursor, valoresEscritos, aplicarValor, fecharParenteseRecemAberto, limparParentesesVazios };
    if (typeof module !== 'undefined' && module.exports) { module.exports = api; return; }

    /* ───────────────────────── tela ───────────────────────── */
    const doc = raiz.document;
    // Estilo do VS Code (dica de parâmetros + lista de sugestões). st = estado da caixa aberta:
    //   {input, nome (parâmetro marcado), manual (marcado pelo teclado), valorI (item realçado da lista; null = nenhum), auto (o realce vem de digitação: 1º item que casa)}
    let cfg = null, caixa = null, ctxCache = null, alvo = null, carregando = false, st = null, dicaVez = null, tocouTeclas = false, escrevendo = false;

    const CSS = '.dica-param{position:fixed;z-index:9000;min-width:300px;max-width:560px;max-height:340px;overflow:auto;background:#252526;color:#cccccc;border:1px solid #454545;border-radius:4px;padding:4px 0;font:12px/1.55 "Cascadia Code",Consolas,Menlo,monospace;box-shadow:0 6px 18px rgba(0,0,0,.55)}'
        + '.dica-param .dp-tit{padding:2px 10px 4px;font-weight:600;color:#e6e6e6}.dica-param .dp-aviso{color:#d7ba7d}.dica-param .dp-erro{color:#f48771}'
        + '.dica-param .dp-assin{padding:2px 10px 5px;border-bottom:1px solid #3c3c3c;color:#9d9d9d}.dica-param .dp-assin b{color:#e6e6e6;text-decoration:underline}.dica-param .dp-assin i{font-style:normal;color:#9cdcfe}'
        + '.dica-param .dp-lin{padding:1px 10px;display:flex;gap:10px;align-items:baseline;border-left:2px solid transparent}.dica-param .dp-lin.atual{background:#2a2d2e;border-left-color:#007acc}'
        + '.dica-param .dp-nome{min-width:72px;color:#9cdcfe}.dica-param .dp-pad{color:#8b8b8b;font:11px Inter,system-ui,sans-serif}.dica-param .dp-pad b{color:#ce9178;font-family:inherit}'
        + '.dica-param .dp-item{display:flex;gap:14px;padding:1px 10px 1px 28px;cursor:pointer;white-space:nowrap}.dica-param .dp-item:hover{background:#2a2d2e}.dica-param .dp-item.sel{background:#094771;color:#fff}'
        + '.dica-param .dp-item .dp-v{font-weight:600}.dica-param .dp-item .dp-d{margin-left:auto;color:#8b8b8b;font:11px Inter,system-ui,sans-serif}.dica-param .dp-item.sel .dp-d{color:#c5d9ea}'
        + '.dica-param .dp-mask{padding:0 10px 0 28px;color:#8b8b8b;font:11px Inter,system-ui,sans-serif}.dica-param .dp-rod{color:#6e7681;font:11px Inter,system-ui,sans-serif;padding:5px 10px 1px;border-top:1px solid #3c3c3c;margin-top:3px}';

    const ehCampo = el => el && el.matches && el.matches('#body-outros .inp-ativo, #body-outros-vazio .inp-ativo-vazio');

    function contexto() {
        const linhas = (cfg.base && cfg.base()) || [];
        const proj = (cfg.projeto && cfg.projeto()) || '';
        if (ctxCache && ctxCache.linhas === linhas && ctxCache.n === linhas.length && ctxCache.proj === proj) return ctxCache.ctx;
        ctxCache = { linhas, n: linhas.length, proj, ctx: construirContexto(linhas, proj) };
        return ctxCache.ctx;
    }

    function fechar() { if (caixa) { caixa.remove(); caixa = null; } alvo = null; st = null; }

    /** Troca o texto do campo (mantém o desfazer nativo quando dá) e avisa a tela (ela guarda o valor no evento "input"). */
    function escrever(input, texto, cursor) {
        escrevendo = true;                                                         // o "input" gerado aqui não redesenha/fecha a caixa (quem chamou cuida disso)
        try {
            if (input.value !== texto) {
                input.focus();
                input.setSelectionRange(0, input.value.length);
                let ok = false;
                try { ok = doc.execCommand('insertText', false, texto); } catch (e) { ok = false; }
                if (!ok || input.value !== texto) { input.value = texto; input.dispatchEvent(new Event('input', { bubbles: true })); }
            }
        } finally { escrevendo = false; }
        const c = Math.max(0, Math.min(cursor == null ? texto.length : cursor, input.value.length));
        input.setSelectionRange(c, c);
    }

    function nomeAtual(input, d) {
        const ed = parametrosEditaveis(d);
        if (st && st.manual && st.nome) return st.nome;
        const n = parametroSobOCursor(input.value, input.selectionStart, d);
        if (n && !ehObraNome(n)) return n;
        return ed.length ? ed[0].nome.toUpperCase() : null;
    }

    /** Valores que a lista mostra para o parâmetro: todos; se há texto digitado que NÃO é um valor exato, só os que casam (começam com ele primeiro, depois os que contêm). */
    function valoresVisiveis(input, d, N) {
        const p = parametrosEditaveis(d).find(x => x.nome.toUpperCase() === N);
        if (!p) return [];
        const digitado = String((valoresEscritos(input.value, input.selectionStart, d) || {})[N] || '').trim().toUpperCase();
        const todos = p.valores.slice();
        if (!digitado || todos.some(v => String(v).toUpperCase() === digitado)) return todos;
        const ini = todos.filter(v => String(v).toUpperCase().startsWith(digitado)), inc = todos.filter(v => !String(v).toUpperCase().startsWith(digitado) && String(v).toUpperCase().includes(digitado));
        return ini.concat(inc);
    }

    /** Grava (ou tira) o valor de um parâmetro pelo nome e deixa o cursor no fim dele. Devolve o plano (ou null). */
    function gravar(input, d, nome, valor) {
        const r = aplicarValor(input.value, input.selectionStart, d, nome, valor);
        if (!r) return null;
        escrever(input, r.texto, r.cursor);
        return r;
    }

    /** Põe o cursor no fim do valor do parâmetro (ou no fim do conteúdo, se ele ainda não foi escrito). */
    function irParaParametro(input, d, nome) {
        const g = grupoDoCursor(input.value, input.selectionStart);
        if (!g) return;
        const t = input.value, fim = g.fecha === -1 ? t.length : g.fecha, base = g.abre + 1;
        const r = lerPedido(t.slice(base, fim), d.params.map(p => p.nome));
        const pt = r.partes.find(x => x.texto !== '' && x.alvo === nome);
        const c = pt ? base + pt.ini + t.slice(base + pt.ini, base + pt.fim).trimEnd().length : fim;
        input.setSelectionRange(c, c);
    }

    /** Fecha a caixa com o cursor depois do ")" do ativo (tirando parênteses vazios: PR() → PR). */
    function concluir(input) {
        const g = grupoDoCursor(input.value, input.selectionStart);
        let cursor = g ? (g.fecha === -1 ? input.value.length : g.fecha + 1) : input.selectionStart;
        fechar();
        const limpo = limparParentesesVazios(input.value);
        if (limpo !== input.value) { cursor = Math.max(0, cursor - (input.value.length - limpo.length)); escrever(input, limpo, cursor); } else input.setSelectionRange(cursor, cursor);
    }

    /** Vai ao parâmetro seguinte (a caixa continua aberta, como os "tab stops" do VS Code); no último, fecha com o cursor depois do ")". `passo` = +1 / -1. */
    function avancar(input, passo) {
        const ed = parametrosEditaveis(dicaVez), nomes = ed.map(p => p.nome.toUpperCase()), i = nomes.indexOf(nomeAtual(input, dicaVez)), n = i + passo;
        if (n < 0) return false;
        if (n >= nomes.length) { concluir(input); return true; }
        st.nome = nomes[n]; st.manual = true; st.valorI = null; st.auto = false;
        irParaParametro(input, dicaVez, st.nome);
        desenhar(input, dicaVez);
        return true;
    }

    /** Aceita o valor realçado (se houver): grava pelo nome e avança. Devolve true se havia valor realçado. */
    function aceitar(input) {
        if (!st || st.valorI === null) return false;
        const N = nomeAtual(input, dicaVez), vis = valoresVisiveis(input, dicaVez, N), v = vis[st.valorI];
        if (st.valorI === vis.length) { avancar(input, 1); return true; }          // linha "digitar": o texto já está escrito no campo; só segue para o próximo parâmetro
        if (v == null) return false;
        gravar(input, dicaVez, N, v);
        if (st) { st.nome = N; st.manual = true; }
        avancar(input, 1);
        return true;
    }

    function desenhar(input, d) {
        if (!caixa) {
            if (!doc.getElementById('dica-param-css')) { const stl = doc.createElement('style'); stl.id = 'dica-param-css'; stl.textContent = CSS; doc.head.appendChild(stl); }
            caixa = doc.createElement('div'); caixa.className = 'dica-param';
            caixa.addEventListener('mousedown', e => e.preventDefault());          // clicar na caixa não tira o foco do campo
            doc.body.appendChild(caixa);
        }
        caixa.innerHTML = '';
        if (d.tipo !== 'parametros') {
            const tit = doc.createElement('div'); tit.className = 'dp-tit ' + (d.tipo === 'nao_encontrado' ? 'dp-erro' : 'dp-aviso');
            tit.textContent = d.mensagem; caixa.appendChild(tit);
        } else {
            const atualNome = nomeAtual(input, d), escritos = valoresEscritos(input.value, input.selectionStart, d);
            const origens = (cfg.origens && cfg.origens(input, d.ativo)) || {};
            const ed = parametrosEditaveis(d);
            const assin = doc.createElement('div'); assin.className = 'dp-assin';                      // assinatura: ATIVO(param1, param2, …) com o parâmetro da vez em destaque (como a dica de parâmetros do VS Code)
            assin.appendChild(doc.createTextNode(d.ativo + '('));
            ed.forEach((p, k) => { if (k) assin.appendChild(doc.createTextNode(', ')); const e = doc.createElement(p.nome.toUpperCase() === atualNome ? 'b' : 'i'); e.textContent = p.nome.toLowerCase(); assin.appendChild(e); });
            assin.appendChild(doc.createTextNode(')')); caixa.appendChild(assin);
            d.params.forEach(p => {
                const N = p.nome.toUpperCase(), atual = !p.obra && N === atualNome;
                const lin = doc.createElement('div'); lin.className = 'dp-lin' + (atual ? ' atual' : '');
                const nm = doc.createElement('span'); nm.className = 'dp-nome'; nm.textContent = p.nome.toLowerCase(); lin.appendChild(nm);
                const pad = doc.createElement('span'); pad.className = 'dp-pad';
                const ori = origens[N], vObra = p.obra && raiz.ObraParams ? raiz.ObraParams.valores()[N] : null;
                const dest = (txt, valor) => { pad.innerHTML = ''; if (valor != null) { const b = doc.createElement('b'); b.textContent = valor; pad.appendChild(b); } pad.appendChild(doc.createTextNode(txt)); };
                if (p.obra) pad.textContent = `da obra (não editável) · ${vObra ? N + '=' + vObra + ' ‹obra›' : (p.padrao != null ? 'sem dado na obra: usa ' + p.padrao : 'sem dado na obra')}`;
                else if (N in escritos) dest(' · manual', escritos[N]);
                else if (ori && ori.valor != null && ori.valor !== '') dest(ori.rotulo ? ' · ' + ori.rotulo : '', ori.valor);
                else pad.textContent = p.porNome ? '(só na fórmula: informe por nome)' : (p.padrao != null ? `padrão ${p.padrao}` : 'sem padrão') + (p.aberto ? ' · aceita qualquer valor' : '');
                lin.appendChild(pad); caixa.appendChild(lin);
                if (atual) {                                                                          // a lista de valores do parâmetro da vez, no estilo da lista de sugestões
                    const vis = valoresVisiveis(input, d, N);
                    if (st && st.auto) {                                                              // digitou: realça o valor igual ao texto (se houver) ou o 1º que casa
                        const igual = vis.findIndex(v => String(v).toUpperCase() === String(escritos[N] || '').trim().toUpperCase());
                        st.valorI = igual >= 0 ? igual : (vis.length ? 0 : vis.length);                  // sem valor que case: a linha "digitar" (a última)
                    }
                    if (st && st.valorI !== null && st.valorI > vis.length) st.valorI = null;
                    vis.forEach((v, k) => {
                        const it = doc.createElement('div'); it.className = 'dp-item' + (st && st.valorI === k ? ' sel' : '');
                        const vs = doc.createElement('span'); vs.className = 'dp-v'; vs.textContent = v; it.appendChild(vs);
                        const det = [];
                        if (p.padrao != null && String(p.padrao).toUpperCase() === String(v).toUpperCase()) det.push('padrão');
                        if (String(escritos[N] || '').toUpperCase() === String(v).toUpperCase()) det.push('escrito');
                        if (det.length) { const ds = doc.createElement('span'); ds.className = 'dp-d'; ds.textContent = det.join(' · '); it.appendChild(ds); }
                        it.addEventListener('click', () => { if (st) st.valorI = k; aceitar(input); });
                        caixa.appendChild(it);
                    });
                    // SEMPRE há uma linha para digitar: o usuário chega nela (↓) e escreve o valor que quiser (mesmo fora da lista)
                    const dig = doc.createElement('div'); dig.className = 'dp-item dp-digitar' + (st && st.valorI === vis.length ? ' sel' : '');
                    const txtDigitado = String(escritos[N] || '').trim();
                    const dv = doc.createElement('span'); dv.className = 'dp-v'; dv.textContent = txtDigitado && !vis.some(v => String(v).toUpperCase() === txtDigitado.toUpperCase()) ? `✎ usar "${txtDigitado}"` : '✎ digitar outro valor…'; dig.appendChild(dv);
                    const dd = doc.createElement('span'); dd.className = 'dp-d'; dd.textContent = p.aberto ? 'aceita qualquer valor' : 'valor livre'; dig.appendChild(dd);
                    dig.addEventListener('click', () => { if (st) { st.valorI = vis.length; st.manual = true; st.nome = N; } input.focus(); irParaParametro(input, d, N); desenhar(input, d); });
                    caixa.appendChild(dig);
                    const outras = p.mascaras.filter(m => !p.valores.includes(String(m).replace(/^%+|%+$/g, '')));            // as máscaras simples já viraram valores
                    if (outras.length) { const m = doc.createElement('div'); m.className = 'dp-mask'; m.title = '% = qualquer texto'; m.textContent = 'também: ' + outras.join('  '); caixa.appendChild(m); }
                    const sel = caixa.querySelector('.dp-item.sel'); if (sel && sel.scrollIntoView) sel.scrollIntoView({ block: 'nearest' });
                }
            });
            const rod = doc.createElement('div'); rod.className = 'dp-rod'; rod.textContent = '↑↓ valores (ou a linha de digitar) · Enter/Tab aceita e avança · Shift+Tab volta · Esc fecha'; caixa.appendChild(rod);
        }
        dicaVez = d;
        const r = input.getBoundingClientRect();
        caixa.style.left = Math.max(4, Math.min(r.left, raiz.innerWidth - 340)) + 'px';
        caixa.style.top = (r.bottom + 3) + 'px';
    }

    function atualizar(input) {
        if (!cfg || !ehCampo(input)) return;
        const c = contextoDoParentese(input.value, input.selectionStart);
        if (!c || !c.nome) { fechar(); return; }
        const linhas = (cfg.base && cfg.base()) || [];
        if (!linhas.length) {                                                      // TASK-128: a base ainda não veio — não desenha nada (antes mostrava "ativo não encontrado" enquanto carregava)
            fechar();
            if (cfg.carregarBase && !carregando) {                                 // carrega uma vez e redesenha quando chegar
                carregando = true;
                Promise.resolve(cfg.carregarBase()).finally(() => { carregando = false; if (doc.activeElement === input) atualizar(input); });
            }
            return;
        }
        if (!st || st.input !== input) st = { input, nome: null, manual: false, valorI: null, auto: false };
        alvo = input;
        desenhar(input, dica(c.nome, contexto()));
    }

    function iniciar(opcoes) {
        cfg = opcoes || {};
        const dentro = e => caixa && st && st.input === e.target && dicaVez && dicaVez.tipo === 'parametros' && grupoDoCursor(e.target.value, e.target.selectionStart);
        // "(" logo depois do nome de um ativo com parâmetros: já escreve o ")" e deixa o cursor no meio
        doc.addEventListener('input', e => {
            if (!ehCampo(e.target) || escrevendo) return;
            if (e.inputType === 'insertText' && e.data === '(' && cfg.base && (cfg.base() || []).length) {
                const input = e.target, cur = input.selectionStart, g = grupoDoCursor(input.value, cur);
                if (g && g.abre === cur - 1 && dica(g.nome, contexto()).tipo === 'parametros') {
                    const f = fecharParenteseRecemAberto(input.value, cur);
                    if (f) { escrever(input, f.texto, f.cursor); atualizar(input); return; }
                }
            }
            if (st && st.input === e.target) { st.auto = true; st.manual = false; }          // digitou: realça o 1º valor que casa (como a lista do VS Code)
            atualizar(e.target);
        });
        // teclado (em captura e ANTES dos demais tratadores, que o consumido não pode alcançar)
        doc.addEventListener('keydown', e => {
            if (e.isComposing || !ehCampo(e.target)) return;
            tocouTeclas = false;
            if (e.key === 'Escape' && caixa) { e.preventDefault(); e.stopImmediatePropagation(); fechar(); return; }       // fecha a caixa e mantém o texto (Ctrl+Z desfaz)
            if (!dentro(e) || e.ctrlKey || e.metaKey || e.altKey) return;
            const input = e.target;
            if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
                const N = nomeAtual(input, dicaVez), n = valoresVisiveis(input, dicaVez, N).length + 1;          // + a linha "digitar"
                if (!N) return;
                e.preventDefault(); e.stopImmediatePropagation(); tocouTeclas = true;
                st.auto = false;
                st.valorI = st.valorI === null ? (e.key === 'ArrowDown' ? 0 : n - 1) : (st.valorI + (e.key === 'ArrowDown' ? 1 : -1) + n) % n;
                if (st.valorI === n - 1) { st.nome = N; st.manual = true; irParaParametro(input, dicaVez, N); }       // chegou na linha "digitar": o cursor vai para o fim do valor, é só escrever
                desenhar(input, dicaVez);
            } else if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault(); e.stopImmediatePropagation(); tocouTeclas = true;
                if (!aceitar(input)) concluir(input);                                // com valor realçado: aceita e avança; sem: só fecha (cursor depois do ")")
            } else if (e.key === 'Tab') {
                e.preventDefault(); e.stopImmediatePropagation(); tocouTeclas = true;
                if (e.shiftKey) { if (!avancar(input, -1)) concluir(input); }
                else if (!aceitar(input)) avancar(input, 1);
            } else if (e.key.length === 1 && st.manual && st.nome) {
                // digitar com um parâmetro marcado pelo teclado: o texto entra nele (se ainda não existe, escreve "nome=" antes)
                const N = st.nome, escritos = valoresEscritos(input.value, input.selectionStart, dicaVez);
                if (!(N in escritos)) {
                    const g = grupoDoCursor(input.value, input.selectionStart), t = input.value, fim = g.fecha === -1 ? t.length : g.fecha, inner = t.slice(g.abre + 1, fim).trim();
                    const rotulo = (inner ? ',' : '') + N.toLowerCase() + '=';
                    escrever(input, t.slice(0, fim) + rotulo + t.slice(fim), fim + rotulo.length);
                } else irParaParametro(input, dicaVez, N);
            }
        }, true);
        doc.addEventListener('keyup', e => { if (ehCampo(e.target) && e.key !== 'Escape' && !tocouTeclas) { if (st && (e.key.indexOf('Arrow') === 0 || e.key === 'Home' || e.key === 'End')) { st.manual = false; st.valorI = null; st.auto = false; } atualizar(e.target); } tocouTeclas = false; });
        doc.addEventListener('click', e => { if (ehCampo(e.target)) { if (st) { st.manual = false; st.valorI = null; st.auto = false; } atualizar(e.target); } });
        doc.addEventListener('focusout', e => { if (ehCampo(e.target)) setTimeout(() => { if (doc.activeElement !== alvo) fechar(); }, 120); });
        // TASK-129: Enter ou sair do campo completa o ")" que falta (em captura: roda antes dos tratadores da própria tela, que leem o valor); TASK-181: parênteses vazios (PR()) viram PR
        const completar = input => {
            const novo = limparParentesesVazios(completarParenteses(input.value));
            if (novo === input.value) return;
            input.value = novo;
            input.dispatchEvent(new Event('input', { bubbles: true }));          // a tela guarda o valor no 'input'
            fechar();
        };
        doc.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.isComposing && ehCampo(e.target)) completar(e.target); }, true);
        doc.addEventListener('blur', e => { if (ehCampo(e.target)) completar(e.target); }, true);
        raiz.addEventListener('scroll', e => { if (caixa && e.target && caixa.contains && caixa.contains(e.target)) return; fechar(); }, true);
        raiz.addEventListener('resize', fechar);
    }

    /** Monta o contexto da base antes do 1º uso (chamado logo depois que a base chega): o primeiro "(" não paga o custo de varrer a base inteira. */
    function aquecer() { if (cfg && cfg.base && (cfg.base() || []).length) contexto(); }

    raiz.AtivoDica = Object.assign({ iniciar, aquecer, contextoAtual: () => (cfg && cfg.base && (cfg.base() || []).length ? contexto() : null) }, api);
})(typeof window !== 'undefined' ? window : globalThis);
