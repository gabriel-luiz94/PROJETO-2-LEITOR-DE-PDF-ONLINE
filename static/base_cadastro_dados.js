/*
 * base_cadastro_dados.js (TASK-144) — parte PURA do cadastro rápido de ativos: gera as linhas da base (como o usuário escreveria na planilha) a partir do rascunho,
 * lê o que se cola do Excel e faz a validação local. A validação completa (a mesma da importação da base) e o simulador do rascunho são do servidor
 * (/api/orcamento/rascunho). Testado em Node: tests/test_base_cadastro_144.py.
 */
(function (raiz) {
    'use strict';
    const AP = typeof require !== 'undefined' && typeof module !== 'undefined' ? require('./ativo_parametros.js') : raiz.AtivoParametros;
    const FD = typeof require !== 'undefined' && typeof module !== 'undefined' ? require('./base_ficha_dados.js') : raiz.BaseFichaDados;
    const T = v => String(v == null ? '' : v).trim();
    const U = v => T(v).toUpperCase();
    const NOME_OK = /^[^\s(),=:|%\[\]]+$/;

    const MDO = { M: 'MATERIAL', MATERIAL: 'MATERIAL', N: 'MAO-DE-OBRA', 'MAO-DE-OBRA': 'MAO-DE-OBRA', 'MÃO-DE-OBRA': 'MAO-DE-OBRA', S: 'MAO-DE-OBRA' };
    const mdoDe = v => MDO[U(v)] || (T(v) ? T(v).toUpperCase() : 'MATERIAL');
    const lvDe = v => { const x = U(v); return x === 'LV' || x === 'LM' ? x : 'A'; };

    function rascunhoVazio() { return { nome: '', apelidos: '', descricao: '', parametros: [], itens: [] }; }
    function itemVazio() { return { condicao: '', codigo: '', desc_codigo: '', mdo: 'M', fator_i: '1', fator_r: '0', lv: 'A', filtro: '', arred: 'NORMAL' }; }
    const listaApelidos = r => String(r.apelidos || '').split(',').map(T).filter(Boolean);
    /** Item que chama outros ativos: o campo CÓDIGO começa com ":=" (ex.: ":= 3*TERM(cabo) + 6*CON(cabo,cabo2)"). */
    const ehChamada = it => T(it.codigo).startsWith(':=');
    const itemPreenchido = it => T(it.codigo) || T(it.condicao);

    /** Linhas da base para o rascunho. `projeto`: nome do projeto das linhas. */
    function gerarLinhas(r, projeto) {
        const nome = T(r.nome), aps = listaApelidos(r), ps = (r.parametros || []).filter(p => T(p.nome));
        const base = { desc_ativo: T(r.descricao), projeto: T(projeto), origem: '', fator_i: 0, fator_r: 0, formula_i: '', formula_r: '', lv: 'A', arred: 'NORMAL', filtro: '', mdo: '', codigo: '', desc_codigo: '', componente: '' };
        const linhas = [];
        if (ps.length) linhas.push(Object.assign({}, base, { ativo: `${nome}=${nome}(${ps.map(p => T(p.nome) + (T(p.padrao) ? '=' + T(p.padrao) : '')).join(',')})${aps.length ? ', ' + aps.join(', ') : ''}` }));
        else if (aps.length) linhas.push(Object.assign({}, base, { ativo: `${nome}, ${aps.join(', ')}` }));
        (r.itens || []).filter(itemPreenchido).forEach(it => {
            const cond = T(it.condicao);
            const celula = cond ? `${nome}(${cond})` : nome;
            if (ehChamada(it)) {
                linhas.push(Object.assign({}, base, { ativo: `${celula}:=${T(it.codigo).slice(2).trim()}` }));
                return;
            }
            const fi = T(it.fator_i).replace(/^=/, '='), fr = T(it.fator_r);
            const formulaI = fi.startsWith('='), formulaR = fr.startsWith('=');
            linhas.push(Object.assign({}, base, {
                ativo: celula, componente: cond ? `${nome}.${cond.toUpperCase().replace(/\s+/g, '')}` : nome, mdo: mdoDe(it.mdo), codigo: T(it.codigo), desc_codigo: T(it.desc_codigo),
                fator_i: formulaI ? 0 : (parseFloat(fi.replace(',', '.')) || 0), formula_i: formulaI ? fi : '', fator_r: formulaR ? 0 : (parseFloat(fr.replace(',', '.')) || 0), formula_r: formulaR ? fr : '',
                lv: lvDe(it.lv), arred: U(it.arred) || 'NORMAL', filtro: T(it.filtro)
            }));
        });
        return linhas;
    }

    /** Texto colado do Excel → itens. Aceita (a) a linha completa da base (13 colunas) ou (b) CÓDIGO ⇥ FATOR I [⇥ FATOR R ⇥ LV]; um cabeçalho na 1ª linha é ignorado. */
    function lerColado(texto) {
        const saida = [];
        String(texto || '').split(/\r\n|\n|\r/).forEach((linha, i) => {
            if (!linha.trim()) return;
            const c = linha.split('\t').map(T);
            if (i === 0 && /^(c[oó]digo|ativo)$/i.test(c[0])) return;
            const it = itemVazio();
            if (c.length >= 10) {                                                  // formato da base: ATIVO;DESC ATIVO;COMPONENTE;PROJETO;MDO;ORIGEM;CODIGO;DESC CODIGO;FATOR I;FATOR R;FILTRO;LV;ARRED
                it.condicao = FD.condicaoDe(c[0]); it.mdo = mdoDe(c[4]) === 'MATERIAL' ? 'M' : 'N'; it.codigo = c[6]; it.desc_codigo = c[7];
                it.fator_i = c[8] || '1'; it.fator_r = c[9] || '0'; it.filtro = c[10] || ''; it.lv = lvDe(c[11]); it.arred = U(c[12]) || 'NORMAL';
                if (/:=/.test(c[0])) { it.condicao = FD.condicaoDe(c[0]); it.codigo = ':=' + c[0].split(':=')[1].split(',')[0].trim(); }
                else if (!c[6]) return;                                           // declaração da base: vira parâmetros/apelidos, não item
            } else {
                it.codigo = c[0]; if (c[1] !== undefined && c[1] !== '') it.fator_i = c[1]; if (c[2] !== undefined && c[2] !== '') it.fator_r = c[2]; if (c[3]) it.lv = lvDe(c[3]);
            }
            saida.push(it);
        });
        return saida;
    }

    /** Combinações coladas ("valor1 ⇥ valor2 ⇥ código" por linha) → itens com condição `p1=%v1%,p2=%v2%` (a tabela por combinação, como o CON). */
    function lerCombinacoes(texto, params) {
        const p = (params || []).map(T).filter(Boolean), saida = [];
        String(texto || '').split(/\r\n|\n|\r/).forEach(linha => {
            if (!linha.trim()) return;
            const c = linha.split('\t').map(T);
            if (c.length < p.length + 1) return;
            const it = itemVazio();
            it.condicao = p.map((nome, i) => `${nome}=%${c[i]}%`).join(',');
            it.codigo = c[p.length];
            if (c[p.length + 1] !== undefined && c[p.length + 1] !== '') it.fator_i = c[p.length + 1];
            saida.push(it);
        });
        return saida;
    }

    /** Validação local (rápida, sem servidor): {erros:[{texto, item?}], avisos:[...]}; `m` = modelo de BaseFichaDados.construir. */
    function validarLocal(r, m) {
        const erros = [], avisos = [], nome = T(r.nome);
        if (!nome) erros.push({ texto: 'Informe o nome do ativo.' });
        else if (!NOME_OK.test(nome)) erros.push({ texto: 'O nome não pode ter espaços, parênteses, vírgula, "=", ":" nem "%".' });
        else if (m && FD.resolver(m, nome)) avisos.push({ texto: `Já existe um ativo "${nome}" neste projeto: ao adicionar à tabela as linhas novas ficam junto das antigas (o cadastro por ativo, que troca as linhas, vem na publicação).` });
        listaApelidos(r).forEach(ap => { if (m && FD.resolver(m, ap) && U(FD.resolver(m, ap)) !== U(nome)) erros.push({ texto: `O apelido "${ap}" já é usado pelo ativo "${FD.resolver(m, ap)}".` }); });
        const nomesParam = (r.parametros || []).map(p => U(p.nome)).filter(Boolean);
        nomesParam.forEach((n, i) => { if (!NOME_OK.test(n)) erros.push({ texto: `Nome de parâmetro inválido: "${n}".` }); if (nomesParam.indexOf(n) !== i) erros.push({ texto: `Parâmetro repetido: "${n}".` }); });
        const novos = new Set();
        (r.itens || []).forEach((it, i) => {
            if (!itemPreenchido(it)) return;
            const n = i + 1, cod = T(it.codigo);
            if (!cod) { erros.push({ texto: `Linha ${n}: informe o código (ou ":= chamada").`, item: i }); return; }
            const cond = T(it.condicao);
            if (cond) {
                const a = AP.analisarCelulaBase(`${nome || 'X'}(${cond})`);
                if (a.erro) erros.push({ texto: `Linha ${n}: condição inválida — ${a.erro}.`, item: i });
                else (a.condicoes || []).forEach(c => { if (nomesParam.length && !nomesParam.includes(U(c.nome))) erros.push({ texto: `Linha ${n}: o parâmetro "${c.nome.toLowerCase()}" não foi declarado no passo 1.`, item: i }); });
                if (!nomesParam.length && !a.erro) erros.push({ texto: `Linha ${n}: há condição, mas o ativo não declara parâmetros (passo 1).`, item: i });
            }
            if (ehChamada(it)) {
                const a = AP.analisarCelulaBase(`${nome || 'X'}${cond ? '(' + cond + ')' : ''}:=${cod.slice(2).trim()}`);
                if (a.erro) erros.push({ texto: `Linha ${n}: chamada inválida — ${a.erro}.`, item: i });
                else (a.itens || []).forEach(x => { const alvo = FD.nomeDoItem(x.ativo); if (U(alvo) !== U(nome) && m && !FD.resolver(m, alvo)) erros.push({ texto: `Linha ${n}: o ativo "${alvo}" não existe na base deste projeto.`, item: i }); });
                return;
            }
            [['fator_i', 'Fator I'], ['fator_r', 'Fator R']].forEach(([k, rot]) => { const v = T(it[k]); if (v && !v.startsWith('=') && isNaN(parseFloat(v.replace(',', '.')))) erros.push({ texto: `Linha ${n}: ${rot} "${v}" não é número nem fórmula (=...).`, item: i }); });
            if (!T(it.desc_codigo) && !(m && m.codigos.get(cod))) avisos.push({ texto: `Linha ${n}: código ${cod} novo na base — informe a descrição.`, item: i });
            const chave = `${U(cond)}|${cod}`;
            if (novos.has(chave)) avisos.push({ texto: `Linha ${n}: código ${cod} repetido na mesma condição (a quantidade soma).`, item: i });
            novos.add(chave);
        });
        if (!(r.itens || []).some(itemPreenchido)) avisos.push({ texto: 'Ainda não há itens (passo 2).' });
        const unico = l => { const v = new Set(); return l.filter(x => !v.has(x.texto) && v.add(x.texto)); };
        return { erros: unico(erros), avisos: unico(avisos) };
    }

    /** TSV das linhas geradas (13 colunas da base, sem cabeçalho) para colar numa planilha. */
    function paraTsv(linhas) {
        const f = v => (v === 0 ? '0' : String(v == null ? '' : v).replace(/[\t\r\n]+/g, ' '));
        return linhas.map(l => [l.ativo, l.desc_ativo, l.componente, l.projeto, l.mdo, l.origem, l.codigo, l.desc_codigo, l.formula_i || l.fator_i, l.formula_r || l.fator_r, l.filtro, l.lv, l.arred].map(f).join('\t')).join('\n');
    }

    const api = { rascunhoVazio, itemVazio, ehChamada, gerarLinhas, lerColado, lerCombinacoes, validarLocal, paraTsv, mdoDe };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.BaseCadastroDados = api;
})(typeof window !== 'undefined' ? window : globalThis);
