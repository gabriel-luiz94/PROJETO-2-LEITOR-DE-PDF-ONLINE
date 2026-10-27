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

    // TASK-092: fases (1x/2x/3x), tensão (13,8kV) e nível (MT/BT) vêm do TEXTO ORIGINAL do cabo ("3xA35P(7,94)-13,8kV 45 m"); sem "Nx" = 1 fase; MT = tensão acima de 1 kV.
    const RE_FASES = /(?<![0-9A-Za-z])([1-9])\s*[xX]\s*(?=[A-Za-z])/;
    const RE_TENSAO = /(\d+(?:[.,]\d+)?)\s*kV/i;
    function derivarCabo(texto) {
        const t = trim(texto), mf = RE_FASES.exec(t), mt = RE_TENSAO.exec(t);
        const tensao = mt ? mt[1].replace('.', ',') + 'kV' : '';
        return { fases: mf ? Number(mf[1]) : 1, tensao, nivel: mt && Number(mt[1].replace(',', '.')) > 1 ? 'MT' : 'BT' };
    }
    /** Dados que as regras do cabo podem usar: o que está na linha (editável) ou, na falta, o derivado do texto original. */
    function extraCabo(row) {
        const d = derivarCabo(row && row.textoOriginal), f = row && trim(row.fases) !== '' ? Number(trim(row.fases)) : d.fases;
        const n = row && trim(row.nivel) ? trim(row.nivel).toUpperCase() : d.nivel;
        return { texto: trim(row && row.textoOriginal), fases: Number.isFinite(f) ? f : d.fases, tensao: row && trim(row.tensao) ? trim(row.tensao) : d.tensao, nivel: n };
    }
    /** `padrao_texto` é SEMPRE regex (busca em qualquer ponto do texto original, sem caixa); inválida não casa. */
    function regexBusca(p, texto) { try { return new RegExp(trim(p), 'i').test(trim(texto)); } catch (e) { return false; } }
    function regraCaboCasa(r, cod, extra) {
        if (!casa(r.padrao, cod)) return false;
        if (!extra) return true;
        if (r.fases !== undefined && r.fases !== null && r.fases !== '' && !(Array.isArray(r.fases) ? r.fases : String(r.fases).split(',')).map(x => Number(trim(x))).includes(extra.fases)) return false;
        if (trim(r.nivel) && trim(r.nivel).toUpperCase() !== extra.nivel) return false;
        if (trim(r.padrao_texto) && !(extra.texto && regexBusca(r.padrao_texto, extra.texto))) return false;
        return true;
    }
    function primeiraCabo(tabela, cod, extra) {
        for (let i = 0; i < tabela.length; i++) { if (tabela[i].ativa === false) continue; if (regraCaboCasa(tabela[i], cod, extra)) return [i, tabela[i]]; }
        return [null, null];
    }

    function infoCabo(codigo, operacao, cfg, extra) {
        const cod = trim(codigo).toUpperCase();
        const [i, r] = cod ? primeiraCabo(cfg.cabos, cod, extra) : [null, null];
        const base = opBase(operacao);
        let pontas = base === 'M' ? 1 : 2;
        if (r && r.pontas && typeof r.pontas === 'object' && Number.isInteger(r.pontas[base])) pontas = r.pontas[base];
        return { codigo: cod, regra: i, valor_cabo: r ? (trim(r.valor_cabo).toUpperCase() || null) : null, pontas, vincula: !(r && r.vincula === false),
            estruturas: r ? (r.estruturas === undefined ? null : r.estruturas) : null, operacoes: r ? (r.operacoes === undefined ? null : r.operacoes) : null,
            op: trim(operacao).toUpperCase(), base, fases: extra ? extra.fases : null, nivel: extra ? extra.nivel : null };
    }

    function infoEstrutura(codigo, cfg) {
        const cod = trim(codigo).toUpperCase();
        const [i, r] = cod ? primeira(cfg.estruturas, 'padrao', cod) : [null, null];
        const d = { codigo: cod, regra: i, qtd: null, qtd_min: null, max_reais: null, lados: 'LIVRE', amarracao: false, aceita_i: true, cabos: null, parametro: 'CABO', se_diferem: 'MAIOR' };
        if (r) {
            // TASK-093: qtd_cabos fixo (min = max) OU faixa qtd_min..qtd_max; max_reais = quantos podem ser cabos REAIS (o excedente só entra como fictício, ex.: SI3 = 1 real + 1 fictício opcional)
            d.qtd = Number.isInteger(r.qtd_max) ? r.qtd_max : (Number.isInteger(r.qtd_cabos) ? r.qtd_cabos : null);
            d.qtd_min = Number.isInteger(r.qtd_min) ? r.qtd_min : d.qtd;
            d.max_reais = Number.isInteger(r.max_reais) ? r.max_reais : null;
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

    function podeComOcupantes(c, e, ocupantes, ficticio) {
        const [ok, motivo] = podeVincular(c, e);
        if (!ok) return [ok, motivo];
        if (e.qtd !== null && ocupantes.length + 1 > e.qtd) return [false, `a estrutura ${e.codigo} já tem os ${e.qtd} cabo(s) que aceita`];
        if (!ficticio && e.max_reais !== null && ocupantes.filter(o => !o.ficticio).length + 1 > e.max_reais) return [false, `a estrutura ${e.codigo} aceita só ${e.max_reais} cabo(s) real(is); o excedente deve ser fictício`];
        if (e.lados === 'IGUAIS' && ocupantes.length) {
            const ref = ocupantes[0];
            const mesmo = ref.info.codigo === c.codigo && (ref.ficticio || ref.info.op === c.op);
            if (!mesmo) return [false, `${e.codigo} exige cabos iguais: ${ref.info.codigo} × ${c.codigo}`];
        }
        return [true, null];
    }

    // TASK-095: só entram no vínculo linhas de operação I, M e as variantes *I e *M. R e *R (retirada/recolocação) ficam de fora — nem listadas, nem contadas.
    const participaVinculo = row => !['R', '*R'].includes(trim(row && row.operacao).toUpperCase());
    const ehCaboValido = row => !!row && !!trim(row.ativo) && participaVinculo(row);
    const ehLinhaEstrutura = row => !!row && (row.entidade === 'ESTRUTURA' || row.entidade === 'POSTE') && participaVinculo(row);
    /** TASK-095: estruturas fictícias do CABO (só o código): preenchem pontas do cabo, sem linha em Outros; entram na Totalizadora como `CABO_ESTRUTURA`. */
    const estruturasFicticiasDe = row => ((row && row.estruturasFicticias) || []).map(x => trim(x).toUpperCase()).filter(Boolean);

    // TASK-094: uma linha de Outros pode ter VÁRIAS estruturas (`9400D, N1, N4` = poste + N1 + N4); cada uma é uma "unidade" vinculável.
    // Identidade da unidade: o índice da linha para a 1ª estrutura (compatível com tudo que já foi salvo) e linha + posição/100 para as seguintes (12 → 12.01 → 12.02).
    const idUnidade = (ie, k) => (k ? Number((ie + k / 100).toFixed(2)) : ie);
    const linhaDe = id => Math.floor(id + 1e-9);
    const slotDe = id => Math.round((id - linhaDe(id)) * 100);
    /** Cabos fictícios da unidade: `cabosFicticios` é da 1ª; as demais ficam em `cabosFicticiosSlots[k]`. */
    const ficticiosDe = (row, k) => (k ? ((row && row.cabosFicticiosSlots && row.cabosFicticiosSlots[k]) || []) : ((row && row.cabosFicticios) || []));
    /** Códigos das estruturas da linha (sem poste). `codigosFn(ativo)` -> lista; sem ele (ou lista vazia) vale o código único de `codigoFn` (linha só com poste). */
    function codigosDaLinha(row, codigoFn, codigosFn) {
        const l = codigosFn ? (codigosFn(row.ativo) || []).filter(Boolean) : [];
        return l.length ? l : [codigoFn(row.ativo, row)];
    }

    function estruturasDisponiveis(outros, codigoFn, codigosFn) {
        const r = [];
        outros.forEach((row, i) => {
            if (!ehLinhaEstrutura(row)) return;
            if (!codigosFn) { r.push(i); return; }
            codigosDaLinha(row, codigoFn || (() => ''), codigosFn).forEach((_, k) => r.push(idUnidade(i, k)));
        });
        return r;
    }

    /** Lista "virtual" de Outros com uma entrada por unidade (as demais linhas passam como estão) e os cabos com os vínculos traduzidos para ela. */
    function expandir(cabos, outros, codigoFn, codigosFn) {
        const virt = [], idDe = [], jDe = {};
        outros.forEach((row, ie) => {
            if (!ehLinhaEstrutura(row)) { idDe.push(ie); virt.push(row); return; }
            const cods = codigosDaLinha(row, codigoFn, codigosFn);
            cods.forEach((cod, k) => {
                const id = idUnidade(ie, k);
                jDe[id] = virt.length; idDe.push(id);
                virt.push(Object.assign({}, row, { cabosFicticios: ficticiosDe(row, k), _codigoUnidade: cod, _linha: ie, _rotulo: cods.length > 1 ? `${row.ativo} [${cod}]` : row.ativo }));
            });
        });
        const cabosV = cabos.map(c => (c && Array.isArray(c.vinculoEstruturas) ? Object.assign({}, c, { vinculoEstruturas: c.vinculoEstruturas.map(id => (jDe[id] !== undefined ? jDe[id] : -1)) }) : c));
        return { virt, idDe, jDe, cabosV, codigoV: (ativo, row) => (row && row._codigoUnidade !== undefined ? row._codigoUnidade : codigoFn(ativo, row)) };
    }

    function analisar(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn) {
        const x = expandir(cabos, outros, codigoFn, codigosFn);
        const a = analisarBase(x.cabosV, x.virt, cfg, prefixoFn, x.codigoV), estruturas = {};
        Object.keys(a.estruturas).forEach(j => { estruturas[x.idDe[Number(j)]] = a.estruturas[j]; });
        return { estruturas, cabos: a.cabos, links: a.links.map(l => Object.assign({}, l, { estrutura: x.idDe[l.estrutura] })) };
    }

    function analisarBase(cabos, outros, cfg, prefixoFn, codigoFn) {
        cfg = normalizarConfig(cfg);
        const ests = estruturasDisponiveis(outros), infoC = {}, infoE = {}, ocup = {};
        cabos.forEach((c, ic) => { if (ehCaboValido(c)) infoC[ic] = infoCabo(prefixoFn(c.ativo) || '', c.operacao, cfg, extraCabo(c)); });
        ests.forEach(ie => { infoE[ie] = infoEstrutura(codigoFn(outros[ie].ativo, outros[ie]), cfg); ocup[ie] = []; });
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
                const [ok, motivo] = podeComOcupantes(icf, e, ocupantes, true);
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
            if (erros.length) estado = 'rd'; else if (tem === 0) estado = 'k'; else if (e.qtd_min !== null && tem < e.qtd_min) estado = 'pu'; else estado = 'gr';
            resE[ie] = { estado, tem, precisa: e.qtd_min, maximo: e.qtd, valor_cabo: valor, erros, avisos, codigo: e.codigo, regra: e.regra, ficticios: fict.length, parametro: e.parametro };
        });
        const resC = {};
        Object.keys(infoC).forEach(k => {
            const ic = Number(k), c = infoC[ic];
            const mine = (cabos[ic].vinculoEstruturas || []).filter(ie => ie in ocup);
            const fictE = c.vincula ? estruturasFicticiasDe(cabos[ic]) : [], errosF = [];
            fictE.forEach(cod => { const [ok, motivo] = podeVincular(c, infoEstrutura(cod, cfg)); if (!ok) errosF.push(motivo); });
            const invalido = mine.some(ie => bad.has(ic + '|' + ie)) || errosF.length > 0;
            const precisa = c.vincula ? c.pontas : 0, tem = mine.length + fictE.length;
            let estado;
            if (invalido) estado = 'rd'; else if (!c.vincula) estado = 'gr'; else if (!tem) estado = 'k'; else if (tem < precisa) estado = 'pu'; else estado = 'gr';
            resC[ic] = { estado, tem, precisa, codigo: c.codigo, valor_cabo: c.valor_cabo, vincula: c.vincula, ficticias: fictE.length, erros: errosF };
        });
        return { estruturas: resE, cabos: resC, links };
    }

    function parametrosPorLinha(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn) {
        const a = analisar(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn), saida = {};
        Object.keys(a.estruturas).sort((x, y) => x - y).forEach(k => {
            const r = a.estruturas[k], linha = linhaDe(Number(k));
            if (r.valor_cabo && r.parametro && !(linha in saida)) saida[linha] = { [r.parametro]: r.valor_cabo };       // várias estruturas na linha: vale a 1ª que tem valor
        });
        return saida;
    }

    function achados(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn) {
        cfg = normalizarConfig(cfg);
        const a = analisar(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn), saida = [];
        const existeM = cabos.some(c => ehCaboValido(c) && opBase(c.operacao) === 'M');
        const existeRetca = outros.some(r => r && /\bRETCA\b/i.test(r.ativo || ''));
        const excecaoZero = existeM || existeRetca;
        const add = (ie, mensagem, explicacao) => saida.push({ linha_id: `OUTROS-${linhaDe(ie)}`, tabela: 'outros', regra_id: 'VINCULO-ESTRUTURA', severidade: 'aviso', mensagem, explicacao });
        Object.keys(a.estruturas).sort((x, y) => x - y).forEach(k => {
            const ie = Number(k), r = a.estruturas[k], linhaO = outros[linhaDe(ie)], nome = codigosFn && (codigosFn(linhaO.ativo) || []).filter(Boolean).length > 1 ? `${linhaO.ativo} [${r.codigo}]` : linhaO.ativo;
            if (r.regra === null) return;
            if (r.tem === 0) {
                if (!excecaoZero) add(ie, `Estrutura "${nome}" sem nenhum cabo vinculado.`, `${r.codigo} exige ${r.precisa || 1} cabo(s); nenhum encontrado, e não há cabo M nem RETCA no projeto para justificar.`);
                return;
            }
            r.erros.forEach(msg => add(ie, `Estrutura "${nome}": vínculo inválido.`, msg));
            if (!r.erros.length && r.precisa !== null && (r.tem < r.precisa || (r.maximo !== null && r.tem > r.maximo))) { const faixa = r.maximo !== null && r.maximo !== r.precisa ? `${r.precisa} a ${r.maximo}` : `${r.precisa}`; add(ie, `Estrutura "${nome}" com ${r.tem} cabo(s) vinculado(s), esperado ${faixa}.`, `${r.codigo} exige ${faixa === String(r.precisa) ? 'exatamente ' : ''}${faixa} cabo(s) vinculado(s).`); }
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

    function executarScript(script, cabos, outros, cfg, prefixoFn, codigoFn, aplicar, codigosFn) {
        const x = expandir(cabos, outros, codigoFn, codigosFn);
        const r = executarScriptBase(script, x.cabosV, x.virt, cfg, prefixoFn, x.codigoV, false);
        r.vinculos.forEach(v => { v.estruturas = v.estruturas.map(j => x.idDe[j]); });
        if (aplicar) r.vinculos.forEach(v => { cabos[v.cabo].vinculoEstruturas = v.estruturas.slice(); cabos[v.cabo].vinculoOrigem = 'script'; });
        return r;
    }

    function executarScriptBase(script, cabos, outros, cfg, prefixoFn, codigoFn, aplicar) {
        cfg = normalizarConfig(cfg);
        const modoConf = CONFLITOS_SCRIPT.includes(script.conflitos) ? script.conflitos : 'avisar';
        const ests = estruturasDisponiveis(outros), infoE = {}, infoC = {}, ligados = {}, fict = {};
        ests.forEach(ie => { infoE[ie] = infoEstrutura(codigoFn(outros[ie].ativo, outros[ie]), cfg); fict[ie] = (outros[ie].cabosFicticios || []).map(x => trim(x).toUpperCase()).filter(Boolean); });
        cabos.forEach((c, ic) => { if (ehCaboValido(c)) { infoC[ic] = infoCabo(prefixoFn(c.ativo) || '', c.operacao, cfg, extraCabo(c)); ligados[ic] = (c.vinculoEstruturas || []).slice(); } });
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
                if ((ligados[ic].length || estruturasFicticiasDe(cabos[ic]).length) && !passo.substituir) { r.ja_tinham++; return; }
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
                const need = Math.max(0, (Number.isInteger(passo.pontas) ? passo.pontas : c.pontas) - (passo.substituir ? 0 : estruturasFicticiasDe(cabos[ic]).length));
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
                    if (ok && outros[ie]._linha !== undefined && escolhidas.some(e2 => outros[e2]._linha === outros[ie]._linha)) continue;      // TASK-094: 2 estruturas do mesmo poste não valem como as 2 pontas
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
                    ['qtd_min', 'qtd_max', 'max_reais'].forEach(k => { if (r[k] != null && (!Number.isInteger(r[k]) || r[k] < 1)) erros.push(`${p}: '${k}' precisa ser um inteiro >= 1.`); });
                    if (Number.isInteger(r.qtd_min) && Number.isInteger(r.qtd_max) && r.qtd_min > r.qtd_max) erros.push(`${p}: 'qtd_min' não pode ser maior que 'qtd_max'.`);
                    if (r.lados != null && !['IGUAIS', 'LIVRE'].includes(r.lados)) erros.push(`${p}: 'lados' precisa ser IGUAIS ou LIVRE.`);
                    if (r.se_diferem != null && !SE_DIFEREM.includes(r.se_diferem)) erros.push(`${p}: 'se_diferem' precisa ser um de ${JSON.stringify(SE_DIFEREM)}.`);
                } else {
                    if (r.fases != null && r.fases !== '' && (Array.isArray(r.fases) ? r.fases : String(r.fases).split(',')).some(x => !/^[1-9]$/.test(trim(x)))) erros.push(`${p}: 'fases' precisa ser números de 1 a 9 (ex.: "3" ou "1,2").`);
                    if (trim(r.nivel) && !['MT', 'BT'].includes(trim(r.nivel).toUpperCase())) erros.push(`${p}: 'nivel' precisa ser MT ou BT.`);
                    let mt = null; if (typeof r.padrao_texto === 'string' && r.padrao_texto) { try { new RegExp(r.padrao_texto); } catch (e) { mt = e.message; } } if (mt) erros.push(`${p}: 'padrao_texto' regex inválida (${mt}).`);
                }
                if (tab === 'cabos' && r.pontas != null && (typeof r.pontas !== 'object' || Object.values(r.pontas).some(v => !Number.isInteger(v) || v < 0))) erros.push(`${p}: 'pontas' precisa ser {"I": n, "M": n} com inteiros >= 0.`);
            });
        });
        return erros;
    }

    const api = { participaVinculo, estruturasFicticiasDe, idUnidade, linhaDe, slotDe, ficticiosDe, codigosDaLinha, derivarCabo, extraCabo, opFicticio, validarConfig, casa, configVazia, normalizarConfig, opBase, infoCabo, infoEstrutura, podeVincular, podeComOcupantes, estruturasDisponiveis, validarScript, executarScript, analisar, parametrosPorLinha, achados, ehCaboValido };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.VinculacaoRegras = api;
})(typeof window !== 'undefined' ? window : globalThis);
