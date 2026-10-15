/*
 * grade_tabela.js (TASK-067) — lógica pura da "grade estilo Excel" das tabelas Cabos e Outros do Resumo.
 * Sem DOM: copiar o intervalo em TSV, ler o que vem do Excel/e-mail, aplicar a colagem sobre os dados
 * (sobrescreve a partir da célula e cria linhas novas ao passar do fim), limpar intervalo.
 * A fonte da verdade continua sendo tableStates[...].data; quem chama (resumo.js) cuida de histórico,
 * reclassificação de entidade, recálculo de qtdAtivos e renderização. Testado em Node: tests/test_grade_tabela.py.
 */
(function (raiz) {
    'use strict';

    const OPERACOES = ['I', '*I', 'R', '*R', 'M', '*M'];
    const COLUNAS = { cabos: ['operacao', 'ativo', 'qtdAtivos'], outros: ['operacao', 'ativo'] };
    const MAX_LINHAS = 5000;

    function colunasDe(tipo) { return COLUNAS[tipo] || COLUNAS.outros; }

    /** TSV de planilha -> matriz de strings. Aceita CRLF, campos entre aspas (Excel) e ignora linhas totalmente vazias. */
    function parseColagem(texto, maxLinhas) {
        const limite = maxLinhas || MAX_LINHAS;
        const t = String(texto == null ? '' : texto).replace(/\r\n?/g, '\n');
        const linhas = [];
        let campo = '', linha = [], aspas = false, i = 0;
        const fechaCampo = () => { linha.push(campo); campo = ''; };
        const fechaLinha = () => { fechaCampo(); linhas.push(linha); linha = []; };
        while (i < t.length) {
            const c = t[i];
            if (aspas) {
                if (c === '"') { if (t[i + 1] === '"') { campo += '"'; i++; } else aspas = false; }
                else campo += c;
            } else if (c === '"' && campo === '') aspas = true;
            else if (c === '\t') fechaCampo();
            else if (c === '\n') fechaLinha();
            else campo += c;
            i++;
        }
        if (campo !== '' || linha.length) fechaLinha();
        const uteis = linhas.filter(l => l.some(x => x.trim() !== ''));
        return { linhas: uteis.slice(0, limite), truncado: uteis.length > limite, total: uteis.length };
    }

    /** O texto tem estrutura de planilha (várias linhas ou colunas)? Se não, o campo trata a colagem sozinho. */
    function ehColagemMultipla(texto) {
        const t = String(texto == null ? '' : texto).replace(/\r\n?/g, '\n').replace(/\n$/, '');
        return t.indexOf('\n') >= 0 || t.indexOf('\t') >= 0;
    }

    function operacaoValida(v) { return OPERACOES.indexOf(String(v == null ? '' : v).trim().toUpperCase()) >= 0; }

    function novaLinha(tipo, operacao) {
        return { entidade: tipo === 'cabos' ? 'CABO' : '0', operacao: operacao || 'I', ativo: '' };
    }

    /**
     * Cola `texto` em `dados` (mutando) a partir de (linha, coluna). Como no Excel: sobrescreve e, passando do fim, cria linhas.
     * Devolve { alteradas:[idx], criadas, opsInvalidas, linhasColadas, truncado }.
     */
    function aplicarColagem(dados, texto, opts) {
        const tipo = opts.tipo, cols = colunasDe(tipo);
        const parsed = parseColagem(texto, opts.maxLinhas);
        const mat = parsed.linhas;
        const res = { alteradas: [], criadas: 0, opsInvalidas: 0, linhasColadas: mat.length, truncado: parsed.truncado };
        if (!mat.length) return res;
        const largura = Math.max.apply(null, mat.map(l => l.length));
        let c0 = Math.max(0, Math.min(opts.coluna | 0, cols.length - 1));
        // uma única coluna colada sobre "Operação": se não forem operações, é uma lista de ativos
        if (largura === 1 && cols[c0] === 'operacao' && !mat.every(l => l[0].trim() === '' || operacaoValida(l[0]))) c0 = cols.indexOf('ativo');
        // "Operação + Ativo" colado com o cursor na coluna Ativo: se a 1ª coluna só tem operações válidas, começa na Operação
        if (largura >= 2 && cols[c0] === 'ativo' && cols.indexOf('operacao') >= 0 && mat.every(l => operacaoValida(l[0]))) c0 = cols.indexOf('operacao');
        const ini = Math.max(0, opts.linha | 0);
        const alt = new Set();
        mat.forEach((celulas, k) => {
            const idx = ini + k;
            if (idx >= dados.length) {
                const ant = dados[dados.length - 1];
                dados.push(novaLinha(tipo, ant ? ant.operacao : 'I'));
                res.criadas++;
            }
            const row = dados[idx];
            celulas.forEach((valor, j) => {
                const campo = cols[c0 + j];
                if (!campo) return;                                  // sobrou coluna: descarta
                const v = String(valor).trim();
                if (campo === 'operacao') {
                    if (v === '') return;
                    if (operacaoValida(v)) row.operacao = v.toUpperCase();
                    else res.opsInvalidas++;                         // mantém a operação que já estava
                } else if (campo === 'ativo') {
                    row.ativo = v.toUpperCase();
                } else if (campo === 'qtdAtivos') {
                    if (/^-/.test(v)) row.qtdAtivos = v;             // só o sinal é manual; o número é recalculado
                    else if (row.qtdAtivos !== undefined && String(row.qtdAtivos).trim().startsWith('-') && v !== '') row.qtdAtivos = v;
                }
            });
            alt.add(idx);
        });
        res.alteradas = Array.from(alt).sort((a, b) => a - b);
        return res;
    }

    /** Intervalo [r1..r2] x [c1..c2] (índices de linha e de coluna da tabela) -> TSV. */
    function textoDoIntervalo(dados, tipo, r1, c1, r2, c2) {
        const cols = colunasDe(tipo), saida = [];
        for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
            const row = dados[r];
            if (!row) continue;
            const cel = [];
            for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
                const v = row[cols[c]];
                cel.push(v === undefined || v === null ? '' : String(v));
            }
            saida.push(cel.join('\t'));
        }
        return saida.join('\n');
    }

    /** Delete no intervalo: esvazia o Ativo (a operação nunca fica vazia; a quantidade é recalculada). Devolve os índices alterados. */
    function limparIntervalo(dados, tipo, r1, c1, r2, c2) {
        const cols = colunasDe(tipo), alt = [];
        for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
            const row = dados[r];
            if (!row) continue;
            let mudou = false;
            for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
                if (cols[c] === 'ativo' && row.ativo !== '') { row.ativo = ''; mudou = true; }
            }
            if (mudou) alt.push(r);
        }
        return alt;
    }


    /* ── TASK-068: colar/copiar para as grades dos modais geradores (Cabos, Postes), colunas configuráveis ── */
    const FASES = ['ABC', 'A', 'B', 'C', 'AC'];
    const RE_MONTADO = /^(.+?)\s+(ABC|AC|AB|BC|A|B|C)\s+(\d+(?:[.,]\d+)?)\s*m$/i;

    /** "CAA 2 ABC 80 m" -> {ativo:'CAA 2', fase:'ABC', comp:80} (ou null). */
    function parseAtivoMontado(texto) {
        const m = RE_MONTADO.exec(String(texto == null ? '' : texto).trim());
        if (!m) return null;
        return { ativo: m[1].trim(), fase: m[2].toUpperCase(), comp: parseFloat(m[3].replace(',', '.')) };
    }

    function numeroDe(v) {
        const n = parseFloat(String(v).trim().replace(',', '.'));
        return isFinite(n) && n > 0 ? n : null;
    }

    /**
     * Cola em `linhas` (mutando) a partir de (linha, coluna). `colunas` = [{campo, tipo}] com tipo op|fase|ativo|comp|qtd.
     * `nova(anterior)` cria a linha que faltar. Devolve { alteradas, criadas, invalidos, linhasColadas, truncado }.
     */
    function aplicarColagemModal(linhas, texto, opts) {
        const cols = opts.colunas, parsed = parseColagem(texto, opts.maxLinhas);
        let mat = parsed.linhas;
        const res = { alteradas: [], criadas: 0, invalidos: 0, linhasColadas: mat.length, truncado: parsed.truncado };
        if (!mat.length) return res;
        const idxTipo = t => cols.findIndex(c => c.tipo === t);
        let c0 = Math.max(0, Math.min(opts.coluna | 0, cols.length - 1));
        const largura = Math.max.apply(null, mat.map(l => l.length));
        const iAtivo = idxTipo('ativo'), iFase = idxTipo('fase'), iComp = idxTipo('comp'), iOp = idxTipo('op');
        if (largura === 1 && iFase >= 0 && iComp >= 0 && mat.every(l => parseAtivoMontado(l[0]))) {
            // coluna única com o ativo já montado: separa em ativo / fase / comprimento
            mat = mat.map(l => { const m = parseAtivoMontado(l[0]); const v = []; v[iAtivo] = m.ativo; v[iFase] = m.fase; v[iComp] = String(m.comp); return v; });
            c0 = 0;
            res.montado = true;
        } else if (largura === 1 && cols[c0].tipo === 'op' && !mat.every(l => l[0].trim() === '' || operacaoValida(l[0]))) {
            c0 = iAtivo;
        } else if (largura >= 2 && cols[c0].tipo === 'ativo' && iOp >= 0 && iOp < c0 && mat.every(l => operacaoValida(l[0]))) {
            c0 = iOp;
        }
        const alt = new Set();
        mat.forEach((celulas, k) => {
            const idx = Math.max(0, opts.linha | 0) + k;
            if (idx >= linhas.length) { linhas.push(opts.nova(linhas[linhas.length - 1])); res.criadas++; }
            const row = linhas[idx];
            celulas.forEach((valor, j) => {
                const col = cols[c0 + j];
                if (!col || valor === undefined) return;
                const v = String(valor).trim();
                if (v === '') return;
                if (col.tipo === 'op') { if (operacaoValida(v)) row[col.campo] = v.toUpperCase(); else res.invalidos++; }
                else if (col.tipo === 'fase') { if (FASES.indexOf(v.toUpperCase()) >= 0) row[col.campo] = v.toUpperCase(); else res.invalidos++; }
                else if (col.tipo === 'ativo') row[col.campo] = v;
                else if (col.tipo === 'comp') { const n = numeroDe(v); if (n !== null) { row[col.campo] = n; row.compManual = true; } else res.invalidos++; }
                else if (col.tipo === 'qtd') { const n = Math.floor(numeroDe(v) || 0); if (n >= 1) row[col.campo] = n; else res.invalidos++; }
            });
            alt.add(idx);
        });
        res.alteradas = Array.from(alt).sort((a, b) => a - b);
        return res;
    }

    /** Intervalo da grade de um modal -> TSV (campos na ordem das colunas). */
    function textoDoIntervaloModal(linhas, colunas, r1, c1, r2, c2) {
        const saida = [];
        for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
            const row = linhas[r];
            if (!row) continue;
            const cel = [];
            for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
                const v = row[colunas[c].campo];
                cel.push(v === undefined || v === null ? '' : String(v));
            }
            saida.push(cel.join('\t'));
        }
        return saida.join('\n');
    }


    /* ── TASK-069: ordenação por coluna (Outros e Cabos) ── */
    const RANK_OPERACAO = { 'I': 0, '*I': 1, 'R': 2, '*R': 3, 'M': 4, '*M': 5 };
    const vazio = v => v === undefined || v === null || String(v).trim() === '';

    function fichas(texto) {                      // "2-U4" -> [2, '-U', 4]; "CAA 2 ABC 80 m" -> ['CAA ', 2, ' ABC ', 80, ' m']
        const out = [], re = /(\d+(?:[.,]\d+)?)|(\D+)/g;
        let m;
        while ((m = re.exec(String(texto))) !== null) out.push(m[1] !== undefined ? parseFloat(m[1].replace(',', '.')) : m[2].toUpperCase());
        return out;
    }
    function compararNatural(a, b) {
        const fa = fichas(a), fb = fichas(b);
        for (let i = 0; i < Math.min(fa.length, fb.length); i++) {
            const x = fa[i], y = fb[i];
            if (typeof x === 'number' && typeof y === 'number') { if (x !== y) return x < y ? -1 : 1; }
            else if (typeof x === 'number') return -1;                // número antes de texto
            else if (typeof y === 'number') return 1;
            else if (x !== y) return x.localeCompare(y, 'pt-BR');
        }
        return fa.length - fb.length;
    }
    function compararCampo(campo, a, b) {         // só chamado com os dois preenchidos
        if (campo === 'operacao') {
            const ra = RANK_OPERACAO[String(a).trim().toUpperCase()], rb = RANK_OPERACAO[String(b).trim().toUpperCase()];
            return (ra === undefined ? 99 : ra) - (rb === undefined ? 99 : rb) || String(a).localeCompare(String(b), 'pt-BR');
        }
        if (campo === 'entidade') {
            const za = String(a) === '0', zb = String(b) === '0';
            if (za !== zb) return za ? 1 : -1;                         // "0" (sem classificação) por último
            return String(a).localeCompare(String(b), 'pt-BR');
        }
        if (campo === 'qtdAtivos') {
            const x = parseFloat(String(a).replace(',', '.')), y = parseFloat(String(b).replace(',', '.'));
            if (isNaN(x) || isNaN(y)) return isNaN(x) === isNaN(y) ? 0 : (isNaN(x) ? 1 : -1);
            return x - y;
        }
        return compararNatural(a, b);                                  // ativo
    }
    function compararLinhas(criterios, la, lb) {
        for (const c of criterios) {
            const a = la[c.campo], b = lb[c.campo], va = vazio(a), vb = vazio(b);
            if (va || vb) { if (va && vb) continue; return va ? 1 : -1; }   // vazio sempre por último, em qualquer sentido
            const r = compararCampo(c.campo, a, b);
            if (r !== 0) return c.dir === 'desc' ? -r : r;
        }
        return 0;
    }
    /** Ordem estável: devolve os índices ANTIGOS na ordem nova. */
    function ordenarLinhas(linhas, criterios) {
        return linhas.map((l, i) => i).sort((i, j) => compararLinhas(criterios, linhas[i], linhas[j]) || (i - j));
    }
    /** ordem (índices antigos por posição nova) -> mapa antigo->novo, para remapear referências por índice (vínculos). */
    function mapaDeOrdem(ordem) {
        const mapa = new Array(ordem.length);
        ordem.forEach((antigo, novo) => { mapa[antigo] = novo; });
        return mapa;
    }
    function estaOrdenado(linhas, criterios) {
        if (!criterios || !criterios.length) return false;
        for (let i = 1; i < linhas.length; i++) if (compararLinhas(criterios, linhas[i - 1], linhas[i]) > 0) return false;
        return true;
    }

    const api = { OPERACOES, COLUNAS, MAX_LINHAS, colunasDe, parseColagem, ehColagemMultipla, operacaoValida, aplicarColagem, textoDoIntervalo, limparIntervalo,
        FASES, parseAtivoMontado, aplicarColagemModal, textoDoIntervaloModal,
        compararLinhas, ordenarLinhas, mapaDeOrdem, estaOrdenado };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    raiz.GradeTabela = api;
})(typeof window !== 'undefined' ? window : globalThis);
