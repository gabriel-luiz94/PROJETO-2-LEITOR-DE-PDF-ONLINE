/*
 * vinculacao_regras.js (TASK-084) — regras de vinculação cabo <-> estrutura: Tabela de Cabos + Tabela de Estruturas por projeto. Sem DOM.
 * MESMA lógica de services/vinculacao_regras.py (paridade testada em tests/test_vinculacao_regras.py).
 * Quem usa: o modal de vinculação (vinculacao_ui.js), a Totalizadora (parâmetro `cabo` das estruturas), a validação do resumo.
 */
(function (raiz) {
    'use strict';
    const ORDEM_MAIOR_PADRAO = ['C', 'P', 'A'];
    const SE_DIFEREM = ['BLOQUEAR', 'MAIOR', 'AVISAR'];
    const RE_REGEX = /[\\^$()\[\]|+*?{}]/;
    const trim = t => String(t == null ? '' : t).trim();

    function casa(padrao, texto) {
        if (padrao == null || padrao === '' || (Array.isArray(padrao) && !padrao.length)) return true;
        const t = trim(texto).toUpperCase();
        if (Array.isArray(padrao)) return padrao.some(x => t === trim(x).toUpperCase());
        const p = trim(padrao);
        if (!p) return true;
        if (!RE_REGEX.test(p)) return p.split(',').map(trim).filter(Boolean).map(x => x.toUpperCase()).includes(t);
        try { return new RegExp(p, 'i').test(trim(texto)); } catch (e) { return false; }
    }

    function configVazia() { return { versao: 2, cabos: [], estruturas: [], ordem_maior: ORDEM_MAIOR_PADRAO.slice(), scripts: [] }; }

    function normalizarConfig(cfg) {
        const c = configVazia();
        if (cfg && typeof cfg === 'object' && !Array.isArray(cfg)) {
            c.cabos = (cfg.cabos || []).filter(r => r && typeof r === 'object' && !Array.isArray(r));
            c.estruturas = (cfg.estruturas || []).filter(r => r && typeof r === 'object' && !Array.isArray(r));
            if (Array.isArray(cfg.ordem_maior) && cfg.ordem_maior.length) c.ordem_maior = cfg.ordem_maior.map(x => trim(x).toUpperCase());
            c.scripts = (cfg.scripts || []).filter(s => s && typeof s === 'object' && !Array.isArray(s));
        }
        return c;
    }

    function opBase(operacao) { const o = trim(operacao).toUpperCase(); return (o === 'M' || o === '*M') ? 'M' : 'I'; }

    function primeira(tabela, campo, valor) {
        for (let i = 0; i < tabela.length; i++) {
            const r = tabela[i];
            if (r.ativa === false) continue;
            if (casa(r[campo], valor)) return [i, r];
        }
        return [null, null];
    }

    function infoCabo(codigo, operacao, cfg) {
        const cod = trim(codigo).toUpperCase();
        const [i, r] = cod ? primeira(cfg.cabos, 'padrao', cod) : [null, null];
        const base = opBase(operacao);
        let pontas = base === 'M' ? 1 : 2;
        if (r && r.pontas && typeof r.pontas === 'object' && Number.isInteger(r.pontas[base])) pontas = r.pontas[base];
        return { codigo: cod, regra: i, valor_cabo: r ? (trim(r.valor_cabo).toUpperCase() || null) : null, pontas, vincula: !(r && r.vincula === false),
            estruturas: r ? (r.estruturas === undefined ? null : r.estruturas) : null, operacoes: r ? (r.operacoes === undefined ? null : r.operacoes) : null,
            op: trim(operacao).toUpperCase(), base };
    }

    function infoEstrutura(codigo, cfg) {
        const cod = trim(codigo).toUpperCase();
        const [i, r] = cod ? primeira(cfg.estruturas, 'padrao', cod) : [null, null];
        const d = { codigo: cod, regra: i, qtd: null, lados: 'LIVRE', amarracao: false, aceita_i: true, cabos: null, parametro: 'CABO', se_diferem: 'MAIOR' };
        if (r) {
            d.qtd = Number.isInteger(r.qtd_cabos) ? r.qtd_cabos : null;
            d.lados = r.lados === 'IGUAIS' ? 'IGUAIS' : 'LIVRE';
            d.amarracao = !!r.amarracao;
            d.aceita_i = r.aceita_i !== false;
            d.cabos = r.cabos === undefined ? null : r.cabos;
            d.parametro = r.parametro === '' ? '' : (trim(r.parametro) || 'cabo').toUpperCase();
            d.se_diferem = SE_DIFEREM.includes(r.se_diferem) ? r.se_diferem : 'MAIOR';
        }
        return d;
    }

    function opFicticio(e) { return (e.regra !== null && e.amarracao && !e.aceita_i) ? 'M' : 'I'; }

    function podeVincular(c, e) {
        if (!c.vincula) return [false, `o cabo ${c.codigo} não vincula (acompanha o cabo principal)`];
        if (Array.isArray(c.operacoes) && c.operacoes.length) {
            const aceitas = c.operacoes.map(x => trim(x).toUpperCase().replace(/\*/g, '')), opn = c.op.replace(/\*/g, '');
            if (!aceitas.includes(opn) && !(opn === 'R' && aceitas.includes('I'))) return [false, `a regra do cabo ${c.codigo} não aceita a operação ${c.op}`];
        }
        if (!casa(c.estruturas, e.codigo)) return [false, `o cabo ${c.codigo} não aceita a estrutura ${e.codigo}`];
        if (!casa(e.cabos, c.codigo)) return [false, `a estrutura ${e.codigo} não aceita o cabo ${c.codigo}`];
        if (c.base === 'M' && e.regra !== null && !e.amarracao) return [false, `cabo M só se liga a estrutura de amarração (${e.codigo} não é)`];
        if (c.base === 'I' && e.regra !== null && !e.aceita_i) return [false, `a estrutura ${e.codigo} só aceita cabo M`];
        return [true, null];
    }

    function podeComOcupantes(c, e, ocupantes) {
        const [ok, motivo] = podeVincular(c, e);
        if (!ok) return [ok, motivo];
        if (e.qtd !== null && ocupantes.length + 1 > e.qtd) return [false, `a estrutura ${e.codigo} já tem os ${e.qtd} cabo(s) que aceita`];
        if (e.lados === 'IGUAIS' && ocupantes.length) {
            const ref = ocupantes[0];
            const mesmo = ref.info.codigo === c.codigo && (ref.ficticio || ref.info.op === c.op);
            if (!mesmo) return [false, `${e.codigo} exige cabos iguais: ${ref.info.codigo} × ${c.codigo}`];
        }
        return [true, null];
    }

    const ehCaboValido = row => !!row && !!trim(row.ativo);
    function estruturasDisponiveis(outros) {
        const r = [];
        outros.forEach((row, i) => { if (row && (row.entidade === 'ESTRUTURA' || row.entidade === 'POSTE')) r.push(i); });
        return r;
    }

    function analisar(cabos, outros, cfg, prefixoFn, codigoFn) {
        cfg = normalizarConfig(cfg);
        const ests = estruturasDisponiveis(outros), infoC = {}, infoE = {}, ocup = {};
        cabos.forEach((c, ic) => { if (ehCaboValido(c)) infoC[ic] = infoCabo(prefixoFn(c.ativo) || '', c.operacao, cfg); });
        ests.forEach(ie => { infoE[ie] = infoEstrutura(codigoFn(outros[ie].ativo), cfg); ocup[ie] = []; });
        cabos.forEach((c, ic) => {
            if (!(ic in infoC)) return;
            (c.vinculoEstruturas || []).forEach(ie => { if (ie in ocup) ocup[ie].push(ic); });
        });
        const links = [], bad = new Set(), resE = {};
        ests.forEach(ie => {
            const e = infoE[ie];
            const fict = (outros[ie].cabosFicticios || []).map(x => trim(x).toUpperCase()).filter(Boolean);
            const tem = ocup[ie].length + fict.length, erros = [], avisos = [], ocupantes = [];
            ocup[ie].forEach(ic => {
                const [ok, motivo] = podeComOcupantes(infoC[ic], e, ocupantes);
                ocupantes.push({ info: infoC[ic], ficticio: false });
                links.push({ cabo: ic, estrutura: ie, ok, motivo });
                if (!ok) { bad.add(ic + '|' + ie); erros.push(motivo); }
            });
            fict.forEach(cod => {
                const icf = infoCabo(cod, opFicticio(e), cfg);
                const [ok, motivo] = podeComOcupantes(icf, e, ocupantes);
                ocupantes.push({ info: icf, ficticio: true });
                if (!ok) erros.push(motivo);
            });
            const valores = ocupantes.map(o => o.info.valor_cabo).filter(v => v !== null);
            let valor = null;
            const distintos = Array.from(new Set(valores)).sort();
            if (distintos.length === 1) valor = distintos[0];
            else if (distintos.length > 1) {
                if (e.se_diferem === 'MAIOR') {
                    const ordem = cfg.ordem_maior, rank = v => ordem.includes(v) ? ordem.indexOf(v) : ordem.length;
                    valor = distintos.slice().sort((a, b) => rank(a) - rank(b) || (a < b ? -1 : a > b ? 1 : 0))[0];
                    avisos.push(`${e.codigo}: cabos diferentes (${distintos.join(', ')}); usado o maior (${valor})`);
                } else avisos.push(`${e.codigo}: cabos diferentes (${distintos.join(', ')}); parâmetro não preenchido`);
            }
            let estado;
            if (erros.length) estado = 'rd'; else if (tem === 0) estado = 'k'; else if (e.qtd !== null && tem < e.qtd) estado = 'pu'; else estado = 'gr';
            resE[ie] = { estado, tem, precisa: e.qtd, valor_cabo: valor, erros, avisos, codigo: e.codigo, regra: e.regra, ficticios: fict.length, parametro: e.parametro };
        });
        const resC = {};
        Object.keys(infoC).forEach(k => {
            const ic = Number(k), c = infoC[ic];
            const mine = (cabos[ic].vinculoEstruturas || []).filter(ie => ie in ocup);
            const invalido = mine.some(ie => bad.has(ic + '|' + ie));
            const precisa = c.vincula ? c.pontas : 0;
            let estado;
            if (invalido) estado = 'rd'; else if (!c.vincula) estado = 'gr'; else if (!mine.length) estado = 'k'; else if (mine.length < precisa) estado = 'pu'; else estado = 'gr';
            resC[ic] = { estado, tem: mine.length, precisa, codigo: c.codigo, valor_cabo: c.valor_cabo, vincula: c.vincula };
        });
        return { estruturas: resE, cabos: resC, links };
    }

    function parametrosPorLinha(cabos, outros, cfg, prefixoFn, codigoFn) {
        const a = analisar(cabos, outros, cfg, prefixoFn, codigoFn), saida = {};
        Object.keys(a.estruturas).forEach(k => {
            const r = a.estruturas[k];
            if (r.valor_cabo && r.parametro) saida[k] = { [r.parametro]: r.valor_cabo };
        });
        return saida;
    }

    function achados(cabos, outros, cfg, prefixoFn, codigoFn) {
        cfg = normalizarConfig(cfg);
        const a = analisar(cabos, outros, cfg, prefixoFn, codigoFn), saida = [];
        const existeM = cabos.some(c => ehCaboValido(c) && opBase(c.operacao) === 'M');
        const existeRetca = outros.some(r => r && /\bRETCA\b/i.test(r.ativo || ''));
        const excecaoZero = existeM || existeRetca;
        const add = (ie, mensagem, explicacao) => saida.push({ linha_id: `OUTROS-${ie}`, tabela: 'outros', regra_id: 'VINCULO-ESTRUTURA', severidade: 'aviso', mensagem, explicacao });
        Object.keys(a.estruturas).forEach(k => {
            const ie = Number(k), r = a.estruturas[k], nome = outros[ie].ativo;
            if (r.regra === null) return;
            if (r.tem === 0) {
                if (!excecaoZero) add(ie, `Estrutura "${nome}" sem nenhum cabo vinculado.`, `${r.codigo} exige ${r.precisa || 1} cabo(s); nenhum encontrado, e não há cabo M nem RETCA no projeto para justificar.`);
                return;
            }
            r.erros.forEach(msg => add(ie, `Estrutura "${nome}": vínculo inválido.`, msg));
            if (!r.erros.length && r.precisa !== null && r.tem !== r.precisa) add(ie, `Estrutura "${nome}" com ${r.tem} cabo(s) vinculado(s), esperado ${r.precisa}.`, `${r.codigo} exige exatamente ${r.precisa} cabo(s) vinculado(s).`);
            r.avisos.forEach(msg => add(ie, `Estrutura "${nome}": ${msg}.`, 'Os cabos vinculados têm valores diferentes para o parâmetro da estrutura.'));
        });
        return saida;
    }

    // ── script de vinculação automática (TASK-084, entrega 2) ──────────────────────────────────────────
    const MODOS_SCRIPT = ['proximidade', 'sequencia'], CONFLITOS_SCRIPT = ['avisar', 'ignorar', 'parar'];

    function regexInvalida(v) {
        if (typeof v !== 'string' || !RE_REGEX.test(v)) return null;
        try { new RegExp(v); return null; } catch (e) { return e.message; }
    }

    function validarScript(script) {
        if (!script || typeof script !== 'object' || Array.isArray(script)) return ['O script precisa ser um objeto JSON.'];
        const erros = [];
        if (script.conflitos != null && !CONFLITOS_SCRIPT.includes(script.conflitos)) erros.push(`'conflitos' precisa ser um de ${JSON.stringify(CONFLITOS_SCRIPT)}.`);
        if (!Array.isArray(script.passos) || !script.passos.length) return erros.concat(["'passos' precisa ser uma lista com pelo menos um passo."]);
        script.passos.forEach((p, i) => {
            const pre = `Passo #${i + 1}`;
            if (!p || typeof p !== 'object' || Array.isArray(p)) { erros.push(`${pre}: precisa ser um objeto.`); return; }
            if (!MODOS_SCRIPT.includes(p.modo)) erros.push(`${pre}: 'modo' precisa ser um de ${JSON.stringify(MODOS_SCRIPT)}.`);
            if (p.raio_m != null && (typeof p.raio_m !== 'number' || !(p.raio_m > 0))) erros.push(`${pre}: 'raio_m' precisa ser um número > 0.`);
            if (p.pontas != null && (!Number.isInteger(p.pontas) || p.pontas < 1)) erros.push(`${pre}: 'pontas' precisa ser um inteiro >= 1.`);
            if (p.cabos != null) {
                if (typeof p.cabos !== 'object' || Array.isArray(p.cabos)) erros.push(`${pre}: 'cabos' precisa ser um objeto ({operacao, ativo}).`);
                else {
                    if (p.cabos.operacao != null && !Array.isArray(p.cabos.operacao)) erros.push(`${pre}: 'cabos.operacao' precisa ser uma lista.`);
                    const m = regexInvalida(p.cabos.ativo); if (m) erros.push(`${pre}: 'cabos.ativo' regex inválida (${m}).`);
                }
            }
            if (p.estruturas != null && p.estruturas !== 'aceitas') {
                if (typeof p.estruturas !== 'object' || Array.isArray(p.estruturas)) erros.push(`${pre}: 'estruturas' precisa ser "aceitas" ou um objeto ({ativo, amarracao}).`);
                else { const m = regexInvalida(p.estruturas.ativo); if (m) erros.push(`${pre}: 'estruturas.ativo' regex inválida (${m}).`); }
            }
        });
        return erros;
    }

    const temXY = r => r._x !== undefined && r._x !== null && r._y !== undefined && r._y !== null;
    const dist = (a, b) => { const dx = a._x - b._x, dy = a._y - b._y; return Math.sqrt(dx * dx + dy * dy); };

    function executarScript(script, cabos, outros, cfg, prefixoFn, codigoFn, aplicar) {
        cfg = normalizarConfig(cfg);
        const modoConf = CONFLITOS_SCRIPT.includes(script.conflitos) ? script.conflitos : 'avisar';
        const ests = estruturasDisponiveis(outros), infoE = {}, infoC = {}, ligados = {}, fict = {};
        ests.forEach(ie => { infoE[ie] = infoEstrutura(codigoFn(outros[ie].ativo), cfg); fict[ie] = (outros[ie].cabosFicticios || []).map(x => trim(x).toUpperCase()).filter(Boolean); });
        cabos.forEach((c, ic) => { if (ehCaboValido(c)) { infoC[ic] = infoCabo(prefixoFn(c.ativo) || '', c.operacao, cfg); ligados[ic] = (c.vinculoEstruturas || []).slice(); } });
        const chavesC = Object.keys(infoC).map(Number).sort((a, b) => a - b);
        const saida = { vinculos: [], conflitos: [], sem_candidato: [], passos: [], parou: false };

        const ocupantesDe = (ie, ignorando) => {
            const lista = chavesC.filter(ic => ic !== ignorando && ligados[ic].includes(ie)).map(ic => ({ info: infoC[ic], ficticio: false }));
            fict[ie].forEach(cod => lista.push({ info: infoCabo(cod, opFicticio(infoE[ie]), cfg), ficticio: true }));
            return lista;
        };
        const pode = (ic, ie) => {
            const c = infoC[ic], e = infoE[ie], oc = ocupantesDe(ie, ic);
            const [ok, motivo] = podeComOcupantes(c, e, oc);
            if (ok && e.qtd === null && oc.length + 1 > c.pontas) return [false, `a estrutura ${e.codigo} já tem os ${c.pontas} cabo(s) esperados`];
            return [ok, motivo];
        };

        (script.passos || []).forEach((passo, ip) => {
            if (saida.parou) return;
            const r = { nome: passo.nome || `Passo ${ip + 1}`, alvo: 0, completos: 0, parciais: 0, sem_candidato: 0, conflitos: 0, ja_tinham: 0 };
            const fc = passo.cabos || {}, ops = (fc.operacao || []).map(x => trim(x).toUpperCase()), opsBase = ops.map(o => o.replace(/\*/g, ''));
            const alvo = [];
            chavesC.forEach(ic => {
                const c = infoC[ic];
                if (!c.vincula) return;
                if (ops.length && !ops.includes(c.op) && !opsBase.includes(c.op.replace(/\*/g, ''))) return;
                if (!casa(fc.ativo, c.codigo)) return;
                if (ligados[ic].length && !passo.substituir) { r.ja_tinham++; return; }
                alvo.push(ic);
            });
            r.alvo = alvo.length;
            const fe = passo.estruturas === undefined ? 'aceitas' : passo.estruturas;
            const estruturaOk = ie => {
                if (fe === 'aceitas' || fe === null) return true;
                if (!casa(fe.ativo, infoE[ie].codigo)) return false;
                if (fe.amarracao !== undefined && fe.amarracao !== null && !!fe.amarracao !== infoE[ie].amarracao) return false;
                return true;
            };
            const candGlobal = ests.filter(estruturaOk), raio = passo.raio_m === undefined ? null : passo.raio_m;
            for (let pos = 0; pos < alvo.length; pos++) {
                const ic = alvo[pos], c = infoC[ic];
                const need = Number.isInteger(passo.pontas) ? passo.pontas : c.pontas;
                const antes = ligados[ic].slice();
                if (passo.substituir) ligados[ic] = [];
                const escolhidas = [], recusas = [];
                let candidatas;
                if (passo.modo === 'sequencia') candidatas = candGlobal.slice(pos, pos + need);
                else {
                    if (!temXY(cabos[ic])) {
                        r.sem_candidato++; saida.sem_candidato.push({ passo: ip, cabo: ic, motivo: 'cabo sem coordenada' }); ligados[ic] = antes; continue;
                    }
                    candidatas = candGlobal.filter(ie => temXY(outros[ie]) && (raio === null || dist(cabos[ic], outros[ie]) <= raio))
                        .sort((a, b) => (dist(cabos[ic], outros[a]) - dist(cabos[ic], outros[b])) || (a - b));
                }
                for (const ie of candidatas) {
                    if (escolhidas.length >= need) break;
                    const [ok, motivo] = pode(ic, ie);
                    if (ok) { escolhidas.push(ie); ligados[ic].push(ie); } else recusas.push(motivo);
                }
                const completo = need > 0 && escolhidas.length === need;
                if (completo || (escolhidas.length && passo.aceitar_parcial)) {
                    saida.vinculos.push({ passo: ip, cabo: ic, estruturas: escolhidas.slice(), completo });
                    if (completo) r.completos++; else r.parciais++;
                } else {
                    ligados[ic] = antes;
                    r.sem_candidato++;
                    saida.sem_candidato.push({ passo: ip, cabo: ic, motivo: 'sem estrutura compatível' + (raio ? ` dentro de ${raio} m` : '') });
                }
                if (recusas.length && modoConf !== 'ignorar') {
                    r.conflitos++; saida.conflitos.push({ passo: ip, cabo: ic, motivo: recusas[0] });
                    if (modoConf === 'parar') { saida.parou = true; break; }
                }
            }
            saida.passos.push(r);
        });
        if (aplicar) saida.vinculos.forEach(v => { cabos[v.cabo].vinculoEstruturas = v.estruturas.slice(); cabos[v.cabo].vinculoOrigem = 'script'; });
        return saida;
    }

    function validarConfig(cfg) {
        if (!cfg || typeof cfg !== 'object' || Array.isArray(cfg)) return ['A configuração precisa ser um objeto.'];
        const erros = [];
        [['cabos', 'Cabo', 'estruturas'], ['estruturas', 'Estrutura', 'cabos']].forEach(([tab, chave, outro]) => {
            const lista = cfg[tab] === undefined ? [] : cfg[tab];
            if (!Array.isArray(lista)) { erros.push(`'${tab}' precisa ser uma lista.`); return; }
            lista.forEach((r, i) => {
                const p = `${chave} #${i + 1}`;
                if (!r || typeof r !== 'object' || Array.isArray(r)) { erros.push(`${p}: precisa ser um objeto.`); return; }
                ['padrao', outro].forEach(campo => {
                    const m = regexInvalida(r[campo]);
                    if (m) erros.push(`${p}: '${campo}' regex inválida (${m}).`);
                    else if (r[campo] != null && typeof r[campo] !== 'string' && !Array.isArray(r[campo])) erros.push(`${p}: '${campo}' precisa ser texto ou lista.`);
                });
                if (tab === 'estruturas') {
                    if (r.qtd_cabos != null && (!Number.isInteger(r.qtd_cabos) || r.qtd_cabos < 1)) erros.push(`${p}: 'qtd_cabos' precisa ser um inteiro >= 1.`);
                    if (r.lados != null && !['IGUAIS', 'LIVRE'].includes(r.lados)) erros.push(`${p}: 'lados' precisa ser IGUAIS ou LIVRE.`);
                    if (r.se_diferem != null && !SE_DIFEREM.includes(r.se_diferem)) erros.push(`${p}: 'se_diferem' precisa ser um de ${JSON.stringify(SE_DIFEREM)}.`);
                } else if (r.pontas != null && (typeof r.pontas !== 'object' || Object.values(r.pontas).some(v => !Number.isInteger(v) || v < 0))) erros.push(`${p}: 'pontas' precisa ser {"I": n, "M": n} com inteiros >= 0.`);
            });
        });
        return erros;
    }

    const api = { opFicticio, validarConfig, casa, configVazia, normalizarConfig, opBase, infoCabo, infoEstrutura, podeVincular, podeComOcupantes, estruturasDisponiveis, validarScript, executarScript, analisar, parametrosPorLinha, achados, ehCaboValido };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.VinculacaoRegras = api;
})(typeof window !== 'undefined' ? window : globalThis);
