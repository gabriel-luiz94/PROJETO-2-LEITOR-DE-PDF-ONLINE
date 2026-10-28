/*
 * ativo_parametros.js (TASK-079 e TASK-080) — ativos com parâmetros por categoria: N1=N1(cabo=XLPE, poste=DT) na base, N1(XLPE,DT) em Outros/Totalizadora.
 * Nos GRUPOS da base também: curinga %, lista | e exceto != (N1(poste=D%, cabo!=XLPE)). Pedido e declaração: só valores concretos.
 * MESMA lógica de services/ativo_parametros.py (paridade testada em tests/test_ativo_parametros.py). Sem DOM. Chaves iguais às do Python.
 * Nomes e valores: letras, números, "_" e "." (sem hífen), sem diferença de caixa.
 */
(function (raiz) {
    'use strict';
    // TASK-096: valores aceitam também "/" (cabos como S1/0). O MESMO nome pode se repetir (N4(cabo, cabo) = dois cabos): a repetição k vira CABO#k por dentro e, na tela, de novo CABO.
    const RE_ID = /^[A-Z0-9_.]+$/, RE_VAL = /^[A-Z0-9_./-]+$/, RE_ALT_BASE = /^[A-Z0-9_.%/-]+$/;      // TASK-100: "-" também (descrição do cabo: 3XS04-7.97KV)
    const decimalDeTensao = c => c.replace(/(\d),(\d+)(?=kv)/gi, '$1.$2');      // 7,97kV = 7.97kV dentro dos parênteses (a vírgula decimal da tensão não separa parâmetros)
    const nomeTela = n => String(n).replace(/#\d+$/, '');
    const nomeRepetido = (vistos, nome) => { vistos[nome] = (vistos[nome] || 0) + 1; return vistos[nome] > 1 ? `${nome}#${vistos[nome]}` : nome; };
    const familia = n => String(n).replace(/#\d+$/, '');
    const norm = t => String(t == null ? '' : t).replace(/\s+/g, '').toUpperCase();
    const cont = (s, c) => s.split(c).length - 1;

    /** Vírgulas FORA de parênteses. */
    function virgulasDeTopo(texto) {
        const pos = []; let prof = 0;
        for (let i = 0; i < texto.length; i++) {
            const ch = texto[i];
            if (ch === '(') prof++;
            else if (ch === ')' && prof > 0) prof--;
            else if (ch === ',' && prof === 0) pos.push(i);
        }
        return pos;
    }

    /** "DT11/300, 11300D" → ["DT11/300", ["11300D"]] (TASK-081). A vírgula dentro de parênteses continua separando parâmetros. */
    function separarApelidos(texto) {
        const t = String(texto == null ? '' : texto), cortes = virgulasDeTopo(t);
        if (!cortes.length) return [t.trim(), []];
        const partes = []; let ini = 0;
        for (const c of cortes) { partes.push(t.slice(ini, c)); ini = c + 1; }
        partes.push(t.slice(ini));
        return [partes[0].trim(), partes.slice(1).map(p => p.trim().toUpperCase())];
    }

    function nomeBase(texto) {
        let t = separarApelidos(texto)[0];
        const i = t.indexOf('('), e = t.indexOf('=');
        if (e !== -1 && (i === -1 || e < i)) t = t.slice(0, e);
        else if (i !== -1) t = t.slice(0, i);
        return t.trim().toUpperCase();
    }

    /** Valor casa o padrão (igualdade, ou "%" = qualquer sequência, inclusive vazia)? Sem diferença de caixa. */
    function casarPadrao(padrao, valor) {
        padrao = String(padrao).toUpperCase(); valor = String(valor).toUpperCase();
        if (padrao.indexOf('%') === -1) return padrao === valor;
        const rx = padrao.split('%').map(p => p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('.*');
        return new RegExp('^(?:' + rx + ')$', 's').test(valor);
    }

    function restricaoVale(restricao, valor) {
        if (typeof restricao === 'string') restricao = { op: '=', alts: [restricao] };
        const acertou = restricao.alts.some(a => casarPadrao(a, valor));
        return restricao.op === '=' ? acertou : !acertou;
    }

    /** Um item de parâmetro (já sem espaços, em maiúsculas). modo: 'pedido' | 'grupo' | 'declaracao'. Retorna [p, erro]. */
    function itemDe(item, modo) {
        let op = '=', nome, valor;
        const nao = item.indexOf('!='), ig = item.indexOf('=');
        if (nao !== -1) { nome = item.slice(0, nao); valor = item.slice(nao + 2); op = '!='; }
        else if (ig !== -1) { nome = item.slice(0, ig); valor = item.slice(ig + 1); }
        else if (modo === 'declaracao') { nome = item; valor = null; }
        else { nome = null; valor = item; }
        const p = { nome, op, valor, alts: [] };
        // TASK-102: a linha de DECLARAÇÃO também pode trazer condições (N1=N1(poste=%D, tipocabo=A)): ela vale como declaração E como o 1º grupo
        const comCondicao = modo === 'grupo' || modo === 'declaracao';
        if (!comCondicao && (op === '!=' || (valor && (valor.indexOf('|') !== -1 || valor.indexOf('%') !== -1)))) return [p, `'${item}': %, | e != só valem nos grupos da base`];
        if (nome !== null && !RE_ID.test(nome)) return [p, `'${nome}' inválido (use letras, números, _ e .)`];
        if (valor === null) return [p, null];
        const alts = comCondicao ? valor.split('|') : [valor];
        const rx = comCondicao ? RE_ALT_BASE : RE_VAL;
        for (const a of alts) if (!rx.test(a)) return [p, `'${a}' inválido (use letras, números, _ e .${comCondicao ? ' ; no grupo da base também % e |' : ''})`];
        p.alts = alts;
        return [p, null];
    }

    function analisarAtivo(texto, modoBase) {
        const original = String(texto == null ? '' : texto).trim();
        const r = { texto: original, base: '', tipo: 'simples', params: [], erro: null, apelidos: [] };
        if (!original) { r.erro = 'ativo vazio'; return r; }
        const i = original.indexOf('('), e = original.indexOf('=');
        if (e !== -1 && (i === -1 || e < i)) {
            const esq = original.slice(0, e).trim().toUpperCase(), dir = original.slice(e + 1).trim();
            const sub = analisarAtivo(dir, false);
            r.base = esq; r.tipo = 'declaracao'; r.params = sub.params;
            if (sub.erro) r.erro = sub.erro;
            else if (sub.tipo !== 'grupo') r.erro = 'declaração precisa ter parâmetros: N1=N1(cabo=XLPE, poste=DT)';
            else if (sub.base !== esq) r.erro = `declaração com nomes diferentes (${esq} = ${sub.base})`;
            return r;
        }
        if (i === -1) { r.base = original.toUpperCase(); return r; }
        r.base = original.slice(0, i).trim().toUpperCase(); r.tipo = 'grupo';
        if (!original.endsWith(')') || cont(original, '(') !== 1 || cont(original, ')') !== 1) { r.erro = 'parênteses mal formados'; return r; }
        const corpo = decimalDeTensao(original.slice(i + 1, -1));
        if (!corpo.trim()) { r.erro = 'parênteses vazios'; return r; }
        const vistos = {};
        for (let item of corpo.split(',')) {
            item = norm(item);
            if (!item) { r.erro = 'parâmetro vazio'; return r; }
            const [p, erro] = itemDe(item, modoBase ? 'grupo' : 'pedido');
            if (erro) { r.erro = erro; return r; }
            if (p.nome === '') { r.erro = 'nome do parâmetro vazio'; return r; }
            if (p.nome !== null && p.valor === '') { r.erro = `parâmetro ${p.nome} sem valor`; return r; }
            if (p.nome !== null) p.nome = nomeRepetido(vistos, p.nome);
            r.params.push(p);
        }
        return r;
    }

    function analisarDeclaracao(texto) {
        const original = String(texto == null ? '' : texto).trim();
        const e = original.indexOf('='), i = original.indexOf('(');
        if (e === -1 || (i !== -1 && e > i)) return analisarAtivo(original, true);
        const esq = original.slice(0, e).trim().toUpperCase(), dir = original.slice(e + 1).trim();
        const r = { texto: original, base: esq, tipo: 'declaracao', params: [], erro: null, apelidos: [] };
        const j = dir.indexOf('(');
        if (j === -1 || !dir.endsWith(')') || cont(dir, '(') !== 1 || cont(dir, ')') !== 1) { r.erro = 'declaração precisa ter parâmetros: N1=N1(cabo=XLPE, poste=DT)'; return r; }
        if (dir.slice(0, j).trim().toUpperCase() !== esq) { r.erro = `declaração com nomes diferentes (${esq} = ${dir.slice(0, j).trim().toUpperCase()})`; return r; }
        const corpo = decimalDeTensao(dir.slice(j + 1, -1));
        if (!corpo.trim()) { r.erro = 'parênteses vazios'; return r; }
        const vistos = {};
        for (let item of corpo.split(',')) {
            item = norm(item);
            if (!item) { r.erro = 'parâmetro vazio'; return r; }
            const [p, erro] = itemDe(item, 'declaracao');
            if (erro) { r.erro = erro; return r; }
            p.nome = nomeRepetido(vistos, p.nome);
            r.params.push(p);
        }
        r.condicionada = r.params.some(p => p.op === '!=' || p.alts.length > 1 || p.alts.some(a => a.indexOf('%') !== -1));
        return r;
    }

    const RE_APELIDO = /^[A-Z0-9_./-]+$/;      // TASK-102: "-" também (B1-MT)

    /** Célula ATIVO da base; o 1º nome é o oficial e os demais são apelidos (só em ativo sem parâmetros na célula). */
    function analisarCelulaBase(texto) {
        const t = String(texto == null ? '' : texto);
        const [principal, apelidos] = separarApelidos(t);
        const e = principal.indexOf('='), i = principal.indexOf('(');
        const a = (e !== -1 && (i === -1 || e < i)) ? analisarDeclaracao(principal) : analisarAtivo(principal, true);
        if (apelidos.length) {
            a.texto = t.trim();
            for (const ap of apelidos) {
                if (!ap) a.erro = a.erro || 'apelido vazio';
                else if (!RE_APELIDO.test(ap)) a.erro = a.erro || `apelido '${ap}' inválido (use letras, números, _, ., - e /)`;
            }
            a.apelidos = apelidos.filter(x => x);
        }
        return a;
    }

    function restricoesDoGrupo(analise) {
        const saida = {};
        if (analise.tipo !== 'grupo' && !(analise.tipo === 'declaracao' && analise.condicionada)) return saida;       // declaração condicionada = também um grupo
        for (const p of analise.params) if (p.nome && (analise.tipo === 'grupo' || p.valor !== null)) saida[p.nome] = { op: p.op, alts: p.alts.slice() };
        return saida;
    }

    function montarAssinaturas(celulas) {
        const saida = {};
        for (const cel of celulas) {
            const a = analisarCelulaBase(cel);
            if (a.erro || !a.params.length) continue;
            const s = saida[a.base] || (saida[a.base] = { nomes: [], padroes: {}, valores: {}, mascaras: {}, fechados: [], declarada: false });
            if (a.tipo === 'declaracao') {
                if (s.declarada) continue;
                s.declarada = true;
                const nomes = a.params.map(p => p.nome);
                s.nomes = nomes.concat(s.nomes.filter(n => !nomes.includes(n)));
                for (const p of a.params) {
                    if (p.valor === null) continue;
                    if (p.op === '=' && p.alts.length === 1 && p.alts[0].indexOf('%') === -1) {          // valor concreto: PADRÃO + valor cadastrado
                        s.padroes[p.nome] = p.valor;
                        const lista = s.valores[p.nome] || (s.valores[p.nome] = []);
                        if (!lista.includes(p.valor)) lista.push(p.valor);
                        if (!s.fechados.includes(p.nome)) s.fechados.push(p.nome);
                    } else {                                                                              // condição (%, |, !=): como num grupo, sem padrão
                        for (const alt of p.alts) {
                            const destino = alt.indexOf('%') !== -1 ? s.mascaras : s.valores;
                            const lista = destino[p.nome] || (destino[p.nome] = []);
                            if (!lista.includes(alt)) lista.push(alt);
                        }
                        if (p.op === '=' && !s.fechados.includes(p.nome)) s.fechados.push(p.nome);
                    }
                }
            } else {
                for (const p of a.params) {
                    if (p.nome === null) continue;
                    if (!s.nomes.includes(p.nome)) s.nomes.push(p.nome);
                    for (const alt of p.alts) {
                        const destino = alt.indexOf('%') !== -1 ? s.mascaras : s.valores;
                        const lista = destino[p.nome] || (destino[p.nome] = []);
                        if (!lista.includes(alt)) lista.push(alt);
                    }
                    if (p.op === '=' && !s.fechados.includes(p.nome)) s.fechados.push(p.nome);
                }
            }
        }
        return saida;
    }

    function resolverPedido(texto, assinaturas) {
        const a = analisarAtivo(texto);
        const r = { base: a.base, texto: a.base, params: {}, faltando: [], erros: [], com_assinatura: false };
        if (a.erro) { r.erros.push(a.erro); r.texto = a.texto.toUpperCase(); return r; }
        if (a.tipo === 'declaracao') { r.erros.push('declaração não pode ser usada como pedido'); return r; }
        const sig = (assinaturas || {})[a.base];
        if (!sig) {
            if (a.params.length) { r.erros.push(`${a.base} não tem variantes cadastradas`); r.texto = a.texto.toUpperCase().replace(/ /g, ''); }
            return r;
        }
        r.com_assinatura = true;
        const nomes = sig.nomes, dados = {};
        let pos = 0;
        for (const p of a.params) {
            if (p.nome === null) {
                while (pos < nomes.length && nomes[pos] in dados) pos++;
                if (pos >= nomes.length) { r.erros.push(`parâmetros demais para ${a.base} (esperado: ${nomes.map(nomeTela).join(', ')})`); continue; }
                dados[nomes[pos]] = p.valor; pos++;
            } else if (!nomes.includes(p.nome)) r.erros.push(`parâmetro ${nomeTela(p.nome)} desconhecido em ${a.base} (existem: ${nomes.map(nomeTela).join(', ')})`);
            else if (p.nome in dados) r.erros.push(`parâmetro ${p.nome} repetido`);
            else dados[p.nome] = p.valor;
        }
        for (const n of nomes) if (!(n in dados) && n in sig.padroes) dados[n] = sig.padroes[n];
        for (const n of nomes) {
            if (n in dados) {
                if (sig.fechados.includes(n)) {
                    const uniao = obj => { const irmaos = nomes.filter(x => familia(x) === familia(n)); return irmaos.length > 1 ? [].concat(...irmaos.map(x => obj[x] || [])) : (obj[n] || []); };      // parâmetro repetido: vale o que qualquer irmão aceita (a ordem não importa)
                    const exatos = uniao(sig.valores), mascaras = uniao(sig.mascaras);
                    if (!exatos.includes(dados[n]) && !mascaras.some(m => casarPadrao(m, dados[n]))) r.erros.push(`${nomeTela(n)}=${dados[n]} não cadastrado em ${a.base} (existem: ${exatos.concat(mascaras).join(', ')})`);
                }
            } else r.faltando.push(n);
        }
        for (const n of nomes) if (n in dados) r.params[n] = dados[n];
        if (!r.faltando.length) r.texto = `${a.base}(${nomes.map(n => dados[n]).join(',')})`;
        else if (Object.keys(dados).length) r.texto = `${a.base}(${nomes.filter(n => n in dados).map(n => nomeTela(n) + '=' + dados[n]).join(',')})`;
        return r;
    }

    /** O grupo vale para o pedido? Grupo comum ({}) sempre; cada parâmetro do grupo precisa existir no pedido e satisfazer a restrição. */
    function linhaVale(paramsLinha, paramsPedido) {
        const fam = {};
        for (const nome of Object.keys(paramsLinha || {})) (fam[familia(nome)] || (fam[familia(nome)] = [])).push(nome);
        for (const f of Object.keys(fam)) {
            const restr = fam[f].map(n => paramsLinha[n]);
            const valores = Object.keys(paramsPedido || {}).filter(n => familia(n) === f).map(n => paramsPedido[n]);
            if (valores.length < restr.length) return false;
            // a ordem não importa: cada restrição do grupo precisa de um valor DIFERENTE do pedido (ex.: cabo=%A%, cabo=%C% x cabo=C16, cabo=A35P)
            const usados = new Set(), cabe = i => {
                if (i === restr.length) return true;
                for (let j = 0; j < valores.length; j++) {
                    if (usados.has(j) || !restricaoVale(restr[i], valores[j])) continue;
                    usados.add(j);
                    if (cabe(i + 1)) return true;
                    usados.delete(j);
                }
                return false;
            };
            if (!cabe(0)) return false;
        }
        return true;
    }

    /** Divide por espaços sem quebrar o que está entre parênteses. */
    function dividirTokens(texto) {
        const saida = []; let atual = '', prof = 0;
        for (const ch of String(texto == null ? '' : texto)) {
            if (ch === '(') prof++;
            else if (ch === ')' && prof > 0) prof--;
            if (/\s/.test(ch) && prof === 0) { if (atual) { saida.push(atual); atual = ''; } }
            else atual += ch;
        }
        if (atual) saida.push(atual);
        return saida;
    }

    /** Assinaturas das linhas da base válidas para o projeto (genéricas + as do projeto), na ordem da tabela. */
    function assinaturasDoProjeto(linhas, projeto) {
        const p = String(projeto || '').trim().toUpperCase();
        const validas = (linhas || []).filter(l => {
            const lp = String(l.projeto || '').trim().toUpperCase();
            return !lp || !p || lp.split('/').map(x => x.trim()).includes(p);
        });
        return montarAssinaturas(validas.map(l => l.ativo || ''));
    }

    /** {APELIDO: NOME_OFICIAL} das linhas da base válidas para o projeto. */
    function apelidosDoProjeto(linhas, projeto) {
        const p = String(projeto || '').trim().toUpperCase(), saida = {};
        for (const l of linhas || []) {
            const lp = String(l.projeto || '').trim().toUpperCase();
            if (lp && p && !lp.split('/').map(x => x.trim()).includes(p)) continue;
            const cel = l.ativo || '';
            if (cel.indexOf(',') === -1) continue;
            const a = analisarCelulaBase(cel);
            if (!a.erro) for (const ap of a.apelidos) if (!(ap in saida)) saida[ap] = a.base;
        }
        return saida;
    }

    const api = { nomeTela, separarApelidos, apelidosDoProjeto, nomeBase, analisarAtivo, analisarDeclaracao, analisarCelulaBase, restricoesDoGrupo, casarPadrao, restricaoVale, montarAssinaturas, resolverPedido, linhaVale, dividirTokens, assinaturasDoProjeto };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.AtivoParametros = api;
})(typeof window !== 'undefined' ? window : globalThis);
