/*
 * vinculacao_regras.js (TASK-084) — regras de vinculação cabo <-> estrutura: Tabela de Cabos + Tabela de Estruturas por projeto. Sem DOM.
 * MESMA lógica de services/vinculacao_regras.py (paridade testada em tests/test_vinculacao_regras.py).
 * Quem usa: o modal de vinculação (vinculacao_ui.js), a Totalizadora (parâmetro `cabo` das estruturas), a validação do resumo.
 */
(function (raiz) {
    'use strict';
    const nodeMod = typeof module !== 'undefined' && module.exports;
    const DC = nodeMod ? require('./dicionario_cabos.js') : raiz.DicionarioCabos;          // TASK-173
    const AC = nodeMod ? require('./anotacao_cabos.js') : raiz.AnotacaoCabos;
    const CP = nodeMod ? require('./convencao_postes.js') : raiz.ConvencaoPostes;          // TASK-174
    const ORDEM_MAIOR_PADRAO = ['C', 'P', 'A'];
    const SE_DIFEREM = ['BLOQUEAR', 'MAIOR', 'AVISAR'], SE_FALTAR = ['AVISAR', 'BLOQUEAR'], ORDENS = ['ORIGEM', 'MAIOR_PRIMEIRO'];          // TASK-169
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

    function configVazia() { return { versao: 2, cabos: [], estruturas: [], ordem_maior: ORDEM_MAIOR_PADRAO.slice(), scripts: [], dicionario_cabos: [] }; }

    function normalizarConfig(cfg) {
        const c = configVazia();
        if (cfg && typeof cfg === 'object' && !Array.isArray(cfg)) {
            c.cabos = (cfg.cabos || []).filter(r => r && typeof r === 'object' && !Array.isArray(r));
            c.estruturas = (cfg.estruturas || []).filter(r => r && typeof r === 'object' && !Array.isArray(r));
            c.dicionario_cabos = DC ? DC.normalizar(cfg.dicionario_cabos) : [];                                          // TASK-173
            if (cfg.postes && typeof cfg.postes === 'object' && !Array.isArray(cfg.postes)) c.postes = cfg.postes;          // TASK-174
            if (Array.isArray(cfg.ordem_maior) && cfg.ordem_maior.length) c.ordem_maior = cfg.ordem_maior.map(x => trim(x).toUpperCase());
            else if (DC && DC.ordemMaiorPadrao(c.dicionario_cabos).length) c.ordem_maior = DC.ordemMaiorPadrao(c.dicionario_cabos);
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
    /** `dic` = dicionário de cabos do projeto (TASK-173): sem tensão/Nx no texto, nível e fases padrão vêm da linha do cabo. */
    function derivarCabo(texto, dic) {
        if (dic === undefined && AC && AC.dicionario) dic = AC.dicionario();          // sem parâmetro: o dicionário do projeto aberto na tela
        const t = trim(texto), mf = RE_FASES.exec(t), mt = RE_TENSAO.exec(t);
        const tensao = mt ? mt[1].replace('.', ',') + 'kV' : '';
        let fases = mf ? Number(mf[1]) : 1, nivel = mt && Number(mt[1].replace(',', '.')) > 1 ? 'MT' : 'BT';
        if (dic && dic.length && DC && AC && (!mt || !mf)) {
            const itens = AC.interpretar(t, dic).itens, p = (itens.filter(i => !i.neutro)[0] || itens[0]);
            if (p) { if (!mt) nivel = DC.nivelPadrao(dic, p.ativo) || nivel; if (!mf) fases = DC.fasesPadrao(dic, p.ativo) || fases; }
        }
        return { fases, tensao, nivel };
    }
    /** Dados que as regras do cabo podem usar: o que está na linha (editável) ou, na falta, o derivado do texto original. */
    function extraCabo(row, dic) {
        const d = derivarCabo(row && row.textoOriginal, dic), f = row && trim(row.fases) !== '' ? Number(trim(row.fases)) : d.fases;
        const n = row && trim(row.nivel) ? trim(row.nivel).toUpperCase() : d.nivel;
        return { texto: trim(row && row.textoOriginal), descricao: descricaoCabo((row && (row.textoOriginal || row.ativo)) || '', ''), fases: Number.isFinite(f) ? f : d.fases, tensao: row && trim(row.tensao) ? trim(row.tensao) : d.tensao, nivel: n };
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
        const dic = cfg.dicionario_cabos, vc = r ? (trim(r.valor_cabo).toUpperCase() || null) : null;
        return { codigo: cod, regra: i, valor_cabo: vc || (dic && dic.length && DC ? DC.valorCabo(dic, cod) : null), pontas, secao: secaoCabo(cod, dic), vincula: !(r && r.vincula === false),
            estruturas: r ? (r.estruturas === undefined ? null : r.estruturas) : null, operacoes: r ? (r.operacoes === undefined ? null : r.operacoes) : null,
            op: trim(operacao).toUpperCase(), base, fases: extra ? extra.fases : null, nivel: extra ? extra.nivel : null, descricao: extra && extra.descricao ? extra.descricao : cod };
    }

    // TASK-096: parâmetros da estrutura com FONTE por cabo vinculado: "cabo=ativo, cabo=ativo, tipocabo=valor_cabo" (ativo = código do cabo sem 3x/tensão; valor_cabo = tipo da regra do cabo).
    // O mesmo nome pode se repetir: a n-ésima ocorrência de um nome recebe o n-ésimo cabo (a ordem dos cabos não importa para o orçamento).
    // TASK-100: o NOME do parâmetro define o que ele recebe — `cabo` = a DESCRIÇÃO COMPLETA do cabo (sem o comprimento, ex.: 3xS04-7,97kV → 3XS04-7.97KV);
    // `tipocabo` = o VALOR DE CABO da Tabela de Cabos (C, P, A…); outro nome = o código do cabo. `nome=fonte` força a fonte (descricao | valor_cabo | ativo).
    const FONTES_PARAMETRO = ['ativo', 'valor_cabo', 'descricao'];
    const fonteDoNome = nome => (/^CABO\d*$/.test(nome) ? 'descricao' : (/^TIPOCABO\d*$/.test(nome) ? 'valor_cabo' : 'ativo'));      // cabo, cabo2, cabo3… / tipocabo, tipocabo2…
    const RE_COMPRIMENTO = /\s+\d+(?:[.,]\d+)?\s*m?\s*$/i;      // o 'm' é opcional (RAMA25 12)
    /** Descrição completa do cabo, sem o comprimento, no formato aceito num parâmetro: maiúsculas, sem espaços, vírgula decimal → ponto, parênteses → "_". */
    /** TASK-170: o cabo fictício de uma estrutura recebe o NÍVEL da primeira regra de cabo que casa com o código E aceita a estrutura (S04 numa LN3 = BT; numa N3 = MT). */
    function infoCaboFicticio(codigo, estrutura, operacao, cfg) {
        const cod = trim(codigo).toUpperCase(), info = infoCabo(cod, operacao, cfg);
        for (const r of cfg.cabos) {
            if (r.ativa === false || !casa(r.padrao, cod) || !casa(r.estruturas, estrutura)) continue;
            const n = trim(r.nivel).toUpperCase();
            info.nivel = (n === 'MT' || n === 'BT') ? n : null;
            break;
        }
        return info;
    }

    function descricaoCabo(texto, codigo) {
        const d = trim(texto).replace(RE_COMPRIMENTO, '').replace(/\s+/g, '').toUpperCase().replace(/,/g, '.').replace(/[()[\]]/g, '_');
        return d || trim(codigo).toUpperCase();
    }
    function normalizarParametros(v) {
        if (v == null || v === '') return null;
        const itens = Array.isArray(v) ? v : String(v).split(',');
        const lista = itens.map(x => {
            if (x && typeof x === 'object') { const n = trim(x.nome).toUpperCase(); return { nome: n, fonte: trim(x.fonte).toLowerCase() || fonteDoNome(n) }; }
            const [n, f] = String(x).split('=');
            const nome = trim(n).toUpperCase();
            return { nome, fonte: trim(f).toLowerCase() || fonteDoNome(nome) };
        }).filter(x => x.nome);
        return lista.length ? lista : null;
    }

    /** TASK-170: [{linha_tem, cabo, param: "nome=valor"}] → [{linha_tem, cabo, param: [NOME, VALOR]}] (item sem condição ou sem param válido é ignorado). */
    function normalizarDefine(v) {
        const saida = [];
        if (!Array.isArray(v)) return saida;
        v.forEach(it => {
            if (!it || typeof it !== 'object' || typeof it.param !== 'string' || !it.param.includes('=')) return;
            const i = it.param.indexOf('='), nome = trim(it.param.slice(0, i)).toUpperCase(), valor = trim(it.param.slice(i + 1)).toUpperCase();
            const lt = (typeof it.linha_tem === 'string' || Array.isArray(it.linha_tem)) && it.linha_tem.length ? it.linha_tem : null;
            const cb = typeof it.cabo === 'string' && trim(it.cabo) ? it.cabo : null;
            if (!nome || !valor || (lt === null && cb === null)) return;
            saida.push({ linha_tem: lt, cabo: cb, param: [nome, valor] });
        });
        return saida;
    }

    function infoEstrutura(codigo, cfg) {
        const cod = trim(codigo).toUpperCase();
        const [i, r] = cod ? primeira(cfg.estruturas, 'padrao', cod) : [null, null];
        const d = { codigo: cod, regra: i, qtd: null, qtd_min: null, max_reais: null, parametros: null, lados: 'LIVRE', amarracao: false, aceita_i: true, cabos: null, parametro: 'CABO', se_diferem: 'MAIOR',
            ocupa: true, herda: null, nivel: null, ordem: 'ORIGEM', distintos: false, se_faltar: 'AVISAR', cabos_vaos: [], define: [] };
        if (r) {
            // TASK-093: qtd_cabos fixo (min = max) OU faixa qtd_min..qtd_max; max_reais = quantos podem ser cabos REAIS (o excedente só entra como fictício, ex.: SI3 = 1 real + 1 fictício opcional)
            d.qtd = Number.isInteger(r.qtd_max) ? r.qtd_max : (Number.isInteger(r.qtd_cabos) ? r.qtd_cabos : null);
            d.qtd_min = Number.isInteger(r.qtd_min) ? r.qtd_min : d.qtd;
            d.max_reais = Number.isInteger(r.max_reais) ? r.max_reais : null;
            d.lados = (r.lados === 'IGUAIS' || r.lados === 'DERIVACAO') ? r.lados : 'LIVRE';          // TASK-170: DERIVACAO = 1 derivação + até 2 do tronco (iguais)
            d.amarracao = !!r.amarracao;
            d.aceita_i = r.aceita_i !== false;
            d.cabos = r.cabos === undefined ? null : r.cabos;
            d.parametro = r.parametro === '' ? '' : (trim(r.parametro) || 'cabo').toUpperCase();
            d.parametros = normalizarParametros(r.parametros);
            if (d.parametros) d.parametro = '';          // a lista substitui o parâmetro único
            d.se_diferem = SE_DIFEREM.includes(r.se_diferem) ? r.se_diferem : 'MAIOR';
            // TASK-169: equipamento (RL, CFU, para-raio…) que NÃO ocupa pontas e lê os cabos do poste; todos opcionais (ausentes = comportamento de antes)
            d.ocupa = r.ocupa !== false;
            d.herda = trim(r.herda).toUpperCase() === 'POSTE' ? 'POSTE' : null;
            d.nivel = ['MT', 'BT'].includes(trim(r.nivel).toUpperCase()) ? trim(r.nivel).toUpperCase() : null;
            d.ordem = trim(r.ordem).toUpperCase() === 'MAIOR_PRIMEIRO' ? 'MAIOR_PRIMEIRO' : 'ORIGEM';
            d.distintos = r.distintos === true;
            d.se_faltar = trim(r.se_faltar).toUpperCase() === 'BLOQUEAR' ? 'BLOQUEAR' : 'AVISAR';
            // TASK-170: um cabo por vão (cada regex precisa de um cabo DIFERENTE) e parâmetros automáticos pela linha do poste / pelo 1º cabo
            if (Array.isArray(r.cabos_vaos) && r.cabos_vaos.some(x => trim(x))) { d.cabos_vaos = r.cabos_vaos.map(trim).filter(Boolean); d.qtd = d.qtd_min = d.cabos_vaos.length; }
            d.define = normalizarDefine(r.define);
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
        if (!(e.cabos_vaos.length ? e.cabos_vaos.some(v => casa(v, c.codigo)) : casa(e.cabos, c.codigo))) return [false, `a estrutura ${e.codigo} não aceita o cabo ${c.codigo}`];
        if (c.base === 'M' && e.regra !== null && !e.amarracao) return [false, `cabo M só se liga a estrutura de amarração (${e.codigo} não é)`];
        if (c.base === 'I' && e.regra !== null && !e.aceita_i) return [false, `a estrutura ${e.codigo} só aceita cabo M`];
        return [true, null];
    }

    /** TASK-170: casamento um a um — cada cabo num vão DIFERENTE cuja regex ele satisfaz. → [índice do vão de cada cabo] ou null. */
    function atribuicaoVaos(vaos, codigos) {
        const n = codigos.length;
        const tenta = (i, usados) => {
            if (i === n) return [];
            for (let j = 0; j < vaos.length; j++) {
                if (usados.has(j) || !casa(vaos[j], codigos[i])) continue;
                const resto = tenta(i + 1, new Set(usados).add(j));
                if (resto !== null) return [j].concat(resto);
            }
            return null;
        };
        return tenta(0, new Set());
    }
    /** TASK-170: índices [i, j] do primeiro par de cabos iguais (o tronco da derivação) ou null. */
    function parDoTronco(codigos) {
        for (let i = 0; i < codigos.length; i++) for (let j = i + 1; j < codigos.length; j++) if (codigos[i] === codigos[j]) return [i, j];
        return null;
    }
    /** `lados: DERIVACAO` → [ocupantes na ordem derivação+tronco, avisos, incompleta]. */
    function listaDerivacao(e, ocupantes) {
        const cods = ocupantes.map(o => o.info.codigo);
        if (ocupantes.length === 0) return [[], [], false];
        if (ocupantes.length === 1) return [ocupantes.slice(), [], false];
        if (ocupantes.length === 2) {
            if (cods[0] === cods[1]) return [[], [`${e.codigo}: falta o cabo da derivação`], true];
            return [[ocupantes[0], ocupantes[1]], [`${e.codigo}: 2 cabos diferentes sem par (${cods[0]}, ${cods[1]}); usado o 1º ligado como derivação`], false];
        }
        const par = parDoTronco(cods);
        if (par === null) return [[], [], false];
        const resto = ocupantes.map((_, i) => i).filter(i => !par.includes(i));
        return [[ocupantes[resto[0]]].concat(par.map(i => ocupantes[i])), [], false];
    }
    /** `cabos_vaos`: os cabos na ORDEM DOS VÃOS (1º nome do parâmetro = 1º vão); vão sem cabo = null. */
    function listaVaos(e, ocupantes) {
        const atrib = atribuicaoVaos(e.cabos_vaos, ocupantes.map(o => o.info.codigo));
        if (atrib === null) return ocupantes.slice();
        const lista = e.cabos_vaos.map(() => null);
        ocupantes.forEach((o, i) => { lista[atrib[i]] = o; });
        while (lista.length && lista[lista.length - 1] === null) lista.pop();
        return lista;
    }
    /** `define`: parâmetros automáticos pela linha do poste e/ou pelo 1º cabo; o 1º item que casa para cada nome vence. → { saida, origem, faltou }. */
    function aplicarDefine(e, ie, outros, ests, infoE, lista) {
        const saida = {}, origem = {}, faltou = [];
        if (!e.define.length) return { saida, origem, faltou };
        const linha = outros[ie]._linha, outras = ests.filter(j => j !== ie && outros[j]._linha === linha).map(j => infoE[j].codigo);
        const prim = lista.find(o => o);
        const primeiro = prim ? prim.info.codigo : null;
        e.define.forEach(item => {
            const [nome, valor] = item.param;
            if (nome in saida) return;
            let achou = null;
            if (item.linha_tem !== null) { achou = outras.find(c => casa(item.linha_tem, c)); if (achou === undefined) return; }
            if (item.cabo !== null && (primeiro === null || !casa(item.cabo, primeiro))) return;
            saida[nome] = valor;
            origem[nome] = item.cabo !== null ? `cabo ${primeiro}` : `linha: ${achou}`;
        });
        e.define.forEach(item => {
            const nome = item.param[0];
            if (nome in saida || faltou.some(f => f.param === nome)) return;
            const lt = item.linha_tem, quem = lt !== null ? (typeof lt === 'string' ? lt : lt.join(', ')).split(',')[0].trim() : '';
            faltou.push({ param: nome, motivo: quem ? `a linha não tem ${quem}` : 'nenhuma condição do define casou' });
        });
        return { saida, origem, faltou };
    }
    /** A ordem em que os cabos vão para os parâmetros. → [lista, avisos_vinculo, incompleta]. */
    function listaDeParametros(e, ocupantes) {
        if (e.lados === 'DERIVACAO') return listaDerivacao(e, ocupantes);
        if (e.cabos_vaos.length) return [listaVaos(e, ocupantes), [], false];
        return [ordenarEJuntar(e, ocupantes), [], false];
    }
    /** Junta os parâmetros do `define` aos do vínculo SEM sobrescrever o que o vínculo já deu. */
    function mesclarDefine(e, parametros, valor, saida) {
        if (!Object.keys(saida).length) return parametros;
        const novo = Object.assign({}, parametros || {});
        if (parametros === null && valor && e.parametro) novo[e.parametro] = valor;
        Object.keys(saida).forEach(n => { if (!(n in novo)) novo[n] = [saida[n]]; });
        return novo;
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
        const codigos = ocupantes.map(o => o.info.codigo).concat([c.codigo]);
        if (e.cabos_vaos.length && atribuicaoVaos(e.cabos_vaos, codigos) === null) return [false, `${e.codigo}: já tem o cabo do vão em que o ${c.codigo} caberia (cada vão aceita um cabo)`];          // TASK-170: um cabo por vão
        if (e.lados === 'DERIVACAO' && codigos.length >= 3 && parDoTronco(codigos) === null) return [false, `${e.codigo}: os dois cabos do tronco precisam ser iguais (${codigos.join(', ')})`];
        return [true, null];
    }

    // TASK-095: só entram no vínculo linhas de operação I, M e as variantes *I e *M. R e *R (retirada/recolocação) ficam de fora — nem listadas, nem contadas.
    const participaVinculo = row => !['R', '*R'].includes(trim(row && row.operacao).toUpperCase());
    const ehCaboValido = row => !!row && !!trim(row.ativo) && participaVinculo(row);
    // TASK-125: o modal lista o que está na tabela Outros — linha ESTRUTURA/POSTE e também a que ainda não foi classificada ('0' ou vazia) mas tem ativo
    // (a classificação automática pode não reconhecer a linha; entidade escolhida à mão — CHAVE, TRAFO, RAMAIS… — continua valendo).
    // Linha de Outros CONGELADA não é estrutura (fora do vínculo); o cabo congelado continua valendo como fonte de um vínculo manual (só não é ligado sozinho pelos algoritmos).
    const ehLinhaEstrutura = row => !!row && !row.congelada && participaVinculo(row) && (row.entidade === 'ESTRUTURA' || row.entidade === 'POSTE' || ((!row.entidade || row.entidade === '0') && !!trim(row.ativo)));
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

    /** TASK-169: seção do cabo para ordenar "maior primeiro" (A185P → 185, S336 → 336; AWG com zero e fração por tabela). */
    const AWG0 = { 2: 33.6, 4: 21.2, 6: 13.3, 8: 8.4 }, AWGX0 = { 1: 53.5, 2: 67.4, 3: 85.0, 4: 107.2 };
    function secaoCabo(codigo, dic) {
        const c = trim(codigo).toUpperCase();
        if (dic && dic.length && DC) { const s = DC.secao(dic, c); if (s !== null) return s; }                  // TASK-173: dicionário primeiro
        let m = /(\d)\/0/.exec(c);
        if (m) return AWGX0[Number(m[1])] !== undefined ? AWGX0[Number(m[1])] : Number(m[1]);
        m = /^[A-Z]+0(\d)$/.exec(c);
        if (m) return AWG0[Number(m[1])] !== undefined ? AWG0[Number(m[1])] : Number(m[1]);
        m = /(\d+)/.exec(c);
        return m ? Number(m[1]) : 0;
    }
    /** `ordem: MAIOR_PRIMEIRO` (estável) e `distintos` (junta os cabos de mesma descrição) ANTES de distribuir nos parâmetros. */
    function ordenarEJuntar(e, ocupantes) {
        let lista = ocupantes.slice();
        if (e.ordem === 'MAIOR_PRIMEIRO') lista = lista.map((o, i) => [o, i]).sort((x, y) => ((y[0].info.secao != null ? y[0].info.secao : secaoCabo(y[0].info.codigo)) - (x[0].info.secao != null ? x[0].info.secao : secaoCabo(x[0].info.codigo))) || (x[1] - y[1])).map(x => x[0]);
        if (e.distintos) { const vistos = new Set(); lista = lista.filter(o => { const k = o.info.descricao; if (vistos.has(k)) return false; vistos.add(k); return true; }); }
        return lista;
    }
    function opAceita(c) {
        if (Array.isArray(c.operacoes) && c.operacoes.length) {
            const aceitas = c.operacoes.map(x => trim(x).toUpperCase().replace(/\*/g, '')), opn = c.op.replace(/\*/g, '');
            return aceitas.includes(opn) || (opn === 'R' && aceitas.includes('I'));
        }
        return true;
    }
    const nomePoste = row => trim(String((row && row.ativo) || '').split(',')[0]);
    function mapaPostes(outros, ests) {
        const porLinha = new Map();
        ests.forEach(ie => { const l = outros[ie]._linha; if (!porLinha.has(l)) porLinha.set(l, []); porLinha.get(l).push(ie); });
        return Array.from(porLinha.keys()).sort((a, b) => a - b).map(l => [l, par(outros[porLinha.get(l)[0]]._poste)]).filter(([, pos]) => pos);
    }
    /** Valor consolidado do cabo e parâmetros por nome a partir dos ocupantes (já ordenados/juntados). */
    function valorEParametros(e, cfg, ocupantes) {
        const avisos = [], reais = ocupantes.filter(o => o), valores = reais.map(o => o.info.valor_cabo).filter(v => v !== null);          // TASK-170: vão sem cabo = null (a posição do parâmetro é mantida)
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
        let parametros = null;
        if (e.parametros) {          // TASK-096: um valor por cabo vinculado (reais e fictícios), na ordem dos ocupantes
            const cont = {}, vezes = {};
            e.parametros.forEach(p => { vezes[p.nome] = (vezes[p.nome] || 0) + 1; });
            e.parametros.forEach(({ nome, fonte }) => {
                const sufixo = /^(?:CABO|TIPOCABO)(\d+)$/.exec(nome), pos = sufixo ? Number(sufixo[1]) - 1 : (cont[nome] || 0);      // cabo2 = o 2º cabo; cabo, cabo = 1º e 2º
                const o = ocupantes[pos]; cont[nome] = (cont[nome] || 0) + 1;
                // `tipocabo` escrito UMA vez numa estrutura com 2 cabos = o tipo consolidado (se_diferem: MAIOR C>P>A…), como no parâmetro único
                let v = o ? (fonte === 'valor_cabo' ? o.info.valor_cabo : (fonte === 'descricao' ? o.info.descricao : o.info.codigo)) : null;
                if (fonte === 'valor_cabo' && vezes[nome] === 1 && reais.length > 1) v = valor;
                if (v) (parametros = parametros || {})[nome] = ((parametros || {})[nome] || []).concat([v]);
            });
        }
        return { valor, avisos, parametros };
    }

    function analisarBase(cabos, outros, cfg, prefixoFn, codigoFn) {
        cfg = normalizarConfig(cfg);
        const ests = estruturasDisponiveis(outros), infoC = {}, infoE = {}, ocup = {};
        cabos.forEach((c, ic) => { if (ehCaboValido(c)) infoC[ic] = infoCabo(prefixoFn(c.ativo) || '', c.operacao, cfg, extraCabo(c, cfg.dicionario_cabos)); });
        ests.forEach(ie => { infoE[ie] = infoEstrutura(codigoFn(outros[ie].ativo, outros[ie]), cfg); });
        const estsOcupa = ests.filter(ie => infoE[ie].ocupa);          // TASK-169: equipamento (ocupa: false) não consome ponta; é resolvido na 2ª passada
        estsOcupa.forEach(ie => { ocup[ie] = []; });
        cabos.forEach((c, ic) => {
            if (!(ic in infoC)) return;
            (c.vinculoEstruturas || []).forEach(ie => { if (ie in ocup) ocup[ie].push(ic); });
        });
        const links = [], bad = new Set(), resE = {};
        estsOcupa.forEach(ie => {
            const e = infoE[ie];
            const fict = (outros[ie].cabosFicticios || []).map(x => trim(x).toUpperCase()).filter(Boolean);
            const tem = ocup[ie].length + fict.length, erros = [], ocupantes = [];
            ocup[ie].forEach(ic => {
                const [ok, motivo] = podeComOcupantes(infoC[ic], e, ocupantes);
                ocupantes.push({ info: infoC[ic], ficticio: false });
                links.push({ cabo: ic, estrutura: ie, ok, motivo });
                if (!ok) { bad.add(ic + '|' + ie); erros.push(motivo); }
            });
            fict.forEach(cod => {
                const icf = infoCaboFicticio(cod, e.codigo, opFicticio(e), cfg);
                const [ok, motivo] = podeComOcupantes(icf, e, ocupantes, true);
                ocupantes.push({ info: icf, ficticio: true });
                if (!ok) erros.push(motivo);
            });
            const [lista, avisosVinculo, incompleta] = listaDeParametros(e, ocupantes);
            const r1 = valorEParametros(e, cfg, lista), valor = r1.valor, avisos = r1.avisos, d = aplicarDefine(e, ie, outros, ests, infoE, lista);
            const parametros = mesclarDefine(e, r1.parametros, valor, d.saida);
            let estado;
            if (erros.length) estado = 'rd'; else if (tem === 0) estado = 'k'; else if ((e.qtd_min !== null && tem < e.qtd_min) || incompleta) estado = 'pu'; else estado = 'gr';
            const defineObj = {}; Object.keys(d.saida).forEach(n => { defineObj[n] = { valor: d.saida[n], origem: d.origem[n] }; });
            resE[ie] = { estado, tem, precisa: e.qtd_min, maximo: e.qtd, valor_cabo: valor, erros, avisos, codigo: e.codigo, regra: e.regra, ficticios: fict.length, parametro: e.parametro, parametros, ocupa: true,
                avisos_vinculo: avisosVinculo, define: defineObj, define_faltou: d.faltou };
        });
        // TASK-169 — 2ª passada: equipamento (RL, CFU, para-raio…) lê os cabos do POSTE, sem ocupar pontas
        const postes = mapaPostes(outros, ests);
        ests.forEach(ie => {
            const e = infoE[ie];
            if (e.ocupa) return;
            const linha = outros[ie]._linha, unidades = estsOcupa.filter(j => outros[j]._linha === linha), candidatos = [], vistos = new Set();
            unidades.forEach(j => ocup[j].forEach(ic => { if (!vistos.has(ic)) { vistos.add(ic); candidatos.push({ info: infoC[ic], ficticio: false }); } }));
            if (par(outros[ie]._poste)) {
                Object.keys(infoC).map(Number).sort((a, b) => a - b).forEach(ic => {
                    const ext = extremos(cabos[ic]);
                    if (!vistos.has(ic) && ext && ext.some(pt => posteMaisPerto(pt, postes, RAIO_POSTE_PADRAO) === linha)) { vistos.add(ic); candidatos.push({ info: infoC[ic], ficticio: false }); }
                });
            }
            unidades.forEach(j => (outros[j].cabosFicticios || []).forEach(x => { if (trim(x)) candidatos.push({ info: infoCaboFicticio(trim(x).toUpperCase(), infoE[j].codigo, opFicticio(infoE[j]), cfg), ficticio: true }); }));
            let lista = candidatos.filter(o => { const c = o.info; return c.vincula && opAceita(c) && casa(e.cabos, c.codigo) && !(e.nivel && c.nivel !== null && c.nivel !== e.nivel); });
            const [listaP, avisosVinculo] = listaDeParametros(e, lista);
            lista = listaP;
            const erros = [], avisosF = [];
            let valor = null, parametros = null;
            if (e.herda === 'POSTE') { const r = valorEParametros(e, cfg, lista); valor = r.valor; parametros = r.parametros; r.avisos.forEach(a => avisosF.push(a)); }
            const d = aplicarDefine(e, ie, outros, ests, infoE, lista);
            parametros = mesclarDefine(e, parametros, valor, d.saida);
            if (e.herda === 'POSTE' && !lista.length) {
                const msg = `${e.codigo}: o poste ${nomePoste(outros[ie])} não tem cabo${e.nivel ? ' ' + e.nivel : ''} ligado`;
                if (e.se_faltar === 'BLOQUEAR') erros.push(msg); else avisosF.push(msg);
            }
            const estado = erros.length ? 'rd' : ((lista.length || e.herda !== 'POSTE') ? 'gr' : 'k');
            const defineObj = {}; Object.keys(d.saida).forEach(n => { defineObj[n] = { valor: d.saida[n], origem: d.origem[n] }; });
            resE[ie] = { estado, tem: lista.filter(o => o).length, precisa: null, maximo: null, valor_cabo: valor, erros, avisos: avisosF, codigo: e.codigo, regra: e.regra, ficticios: 0, parametro: e.parametro, parametros, ocupa: false,
                le_do_poste: { poste: nomePoste(outros[ie]), cabos: lista.filter(o => o).map(o => o.info.codigo) },
                avisos_vinculo: avisosVinculo, define: defineObj, define_faltou: d.faltou };
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

    /** TASK-170: {linha: {CODIGO_DA_ESTRUTURA: {define: {NOME: {valor, origem}}, faltou: {NOME: motivo}}}} — de onde veio cada parâmetro do `define` (e por que não veio), para o balão da tela. */
    function defineDeEstruturas(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn) {
        const a = analisar(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn), saida = {};
        Object.keys(a.estruturas).sort((x, y) => x - y).forEach(k => {
            const r = a.estruturas[k];
            if (!Object.keys(r.define || {}).length && !(r.define_faltou || []).length) return;
            const faltou = {}; (r.define_faltou || []).forEach(f => { faltou[f.param] = f.motivo; });
            ((saida[linhaDe(Number(k))] = saida[linhaDe(Number(k))] || {})[r.codigo]) = { define: r.define || {}, faltou };
        });
        return saida;
    }

    /** TASK-171: {NOME: [valores]} que a estrutura receberia se ESTE cabo fosse o único ligado a ela (mesmas fontes da TASK-100). */
    function parametrosDoCaboNaEstrutura(estrutura, caboRow, cfg, prefixoFn) {
        cfg = normalizarConfig(cfg);
        const e = infoEstrutura(estrutura, cfg), c = infoCabo(prefixoFn((caboRow || {}).ativo) || '', (caboRow || {}).operacao, cfg, extraCabo(caboRow, cfg.dicionario_cabos));
        const r = valorEParametros(e, cfg, [{ info: c, ficticio: false }]);
        let params = r.parametros;
        if (params === null && r.valor && e.parametro) params = { [e.parametro]: [r.valor] };
        return params || {};
    }

    function parametrosPorLinha(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn) {
        const a = analisar(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn), saida = {};
        Object.keys(a.estruturas).sort((x, y) => x - y).forEach(k => {
            const r = a.estruturas[k], linha = linhaDe(Number(k));
            if (r.valor_cabo && r.parametro && !(linha in saida)) saida[linha] = { [r.parametro]: r.valor_cabo };       // várias estruturas na linha: vale a 1ª que tem valor
            else if (r.parametros && !(linha in saida)) saida[linha] = Object.assign({}, r.parametros);
        });
        return saida;
    }

    /** TASK-096: {linha: {CODIGO_DA_ESTRUTURA: {NOME: valor | [valores]}}} — cada estrutura da linha recebe SÓ os parâmetros dela (N1 e N4 do mesmo poste não se misturam). */
    function parametrosPorEstrutura(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn) {
        const a = analisar(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn), saida = {};
        Object.keys(a.estruturas).sort((x, y) => x - y).forEach(k => {
            const r = a.estruturas[k], linha = linhaDe(Number(k));
            const params = r.parametros || (r.valor_cabo && r.parametro ? { [r.parametro]: r.valor_cabo } : null);
            if (!params) return;
            const porCodigo = saida[linha] || (saida[linha] = {});
            if (!(r.codigo in porCodigo)) porCodigo[r.codigo] = params;
        });
        return saida;
    }

    function achados(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn) {
        cfg = normalizarConfig(cfg);
        const a = analisar(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn), saida = [];
        const existeM = cabos.some(c => ehCaboValido(c) && !c.congelada && opBase(c.operacao) === 'M');
        const existeRetca = outros.some(r => r && /\bRETCA\b/i.test(r.ativo || ''));
        const excecaoZero = existeM || existeRetca;
        const add = (ie, mensagem, explicacao) => saida.push({ linha_id: `OUTROS-${linhaDe(ie)}`, tabela: 'outros', regra_id: 'VINCULO-ESTRUTURA', severidade: 'aviso', mensagem, explicacao });
        Object.keys(a.estruturas).sort((x, y) => x - y).forEach(k => {
            const ie = Number(k), r = a.estruturas[k], linhaO = outros[linhaDe(ie)], nome = codigosFn && (codigosFn(linhaO.ativo) || []).filter(Boolean).length > 1 ? `${linhaO.ativo} [${r.codigo}]` : linhaO.ativo;
            if (r.regra === null) return;
            if (r.ocupa === false) {                                           // TASK-169: equipamento lê do poste — só avisa quando falta cabo (ou bloqueia)
                r.erros.forEach(msg => add(ie, `Estrutura "${nome}": ${msg}.`, 'O equipamento lê os cabos do poste e a regra manda bloquear sem cabo (se_faltar = BLOQUEAR).'));
                r.avisos.forEach(msg => add(ie, `Estrutura "${nome}": ${msg}.`, 'O equipamento lê os cabos do poste; sem cabo o parâmetro fica no padrão da base.'));
                return;
            }
            if (r.tem === 0) {
                if (!excecaoZero) add(ie, `Estrutura "${nome}" sem nenhum cabo vinculado.`, `${r.codigo} exige ${r.precisa || 1} cabo(s); nenhum encontrado, e não há cabo M nem RETCA no projeto para justificar.`);
                return;
            }
            r.erros.forEach(msg => add(ie, `Estrutura "${nome}": vínculo inválido.`, msg));
            if (!r.erros.length && r.precisa !== null && (r.tem < r.precisa || (r.maximo !== null && r.tem > r.maximo))) { const faixa = r.maximo !== null && r.maximo !== r.precisa ? `${r.precisa} a ${r.maximo}` : `${r.precisa}`; add(ie, `Estrutura "${nome}" com ${r.tem} cabo(s) vinculado(s), esperado ${faixa}.`, `${r.codigo} exige ${faixa === String(r.precisa) ? 'exatamente ' : ''}${faixa} cabo(s) vinculado(s).`); }
            r.avisos.forEach(msg => add(ie, `Estrutura "${nome}": ${msg}.`, 'Os cabos vinculados têm valores diferentes para o parâmetro da estrutura.'));
            (r.avisos_vinculo || []).forEach(msg => add(ie, `Estrutura "${nome}": ${msg}.`, 'A regra da estrutura (lados DERIVACAO) pede 1 cabo de derivação e até 2 do tronco, iguais entre si.'));          // TASK-170
        });
        return saida;
    }

    // ── script de vinculação automática (TASK-084, entrega 2) ──────────────────────────────────────────
    const MODOS_SCRIPT = ['proximidade', 'sequencia', 'aleatorio'], CONFLITOS_SCRIPT = ['avisar', 'ignorar', 'parar'];

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
            if (p.semente != null && !Number.isInteger(p.semente)) erros.push(`${pre}: 'semente' precisa ser um inteiro.`);
            ['completar', 'puro', 'aceitar_parcial', 'substituir'].forEach(k => { if (p[k] != null && typeof p[k] !== 'boolean') erros.push(`${pre}: '${k}' precisa ser true ou false.`); });
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

    // TASK-099: passo "aleatorio" (reserva do vínculo por coordenadas). Sorteio REPRODUZÍVEL: mesma semente + mesmos dados = mesmo resultado (mulberry32, igual em Python).
    function mulberry32(a) {
        return () => {
            a = (a + 0x6D2B79F5) | 0;
            let t = Math.imul(a ^ (a >>> 15), 1 | a);
            t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
            return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
        };
    }
    function embaralhar(lista, rnd) {
        const l = lista.slice();
        for (let i = l.length - 1; i > 0; i--) { const j = Math.floor(rnd() * (i + 1)); [l[i], l[j]] = [l[j], l[i]]; }
        return l;
    }

    /* ── vinculação POR POSTE (TASK-166) — espelho de services/vinculacao_regras.py ──
       Cada ponta do vão (`_geo`: 1º e último ponto) termina em um poste (`_poste` na linha de Outros): o problema se separa por poste. Cada estrutura do poste procura, entre os
       cabos que CHEGAM nele, os que cumprem as validações dela e consome UMA ponta de cada. Prioridade: ordem da Tabela de Estruturas, depois a ordem na linha do poste. */
    const RAIO_POSTE_PADRAO = 5.0;
    const par = p => (Array.isArray(p) && p.length >= 2 && typeof p[0] !== 'boolean' && Number.isFinite(Number(p[0])) && Number.isFinite(Number(p[1])) && p[0] !== null && p[1] !== null && p[0] !== '' && p[1] !== '' ? [Number(p[0]), Number(p[1])] : null);
    function extremos(cabo) {
        const g = cabo._geo;
        if (!Array.isArray(g) || g.length < 2) return null;
        const a = par(g[0]), b = par(g[g.length - 1]);
        return a && b ? [a, b] : null;
    }
    function posteMaisPerto(pt, postes, raio) {
        let melhor = null, dmin = null;
        postes.forEach(([linha, pos]) => {
            const d = Math.sqrt((pt[0] - pos[0]) ** 2 + (pt[1] - pos[1]) ** 2);
            if (d <= raio && (dmin === null || d < dmin)) { melhor = linha; dmin = d; }
        });
        return melhor;
    }

    function vincularPorPoste(cabos, outros, cfg, prefixoFn, codigoFn, codigosFn, raio, aplicar) {
        const x = expandir(cabos, outros, codigoFn, codigosFn);
        const r = vincularPorPosteBase(x.cabosV, x.virt, cfg, prefixoFn, x.codigoV, typeof raio === 'number' && raio > 0 ? raio : RAIO_POSTE_PADRAO);
        r.vinculos.forEach(v => { v.estruturas = v.estruturas.map(j => x.idDe[j]); });
        r.achados.forEach(a => { if (a.estrutura !== null && a.estrutura !== undefined) a.estrutura = x.idDe[a.estrutura]; });
        if (aplicar) r.vinculos.forEach(v => {
            cabos[v.cabo].vinculoEstruturas = v.estruturas.slice();
            if (v.estruturas.length) cabos[v.cabo].vinculoOrigem = 'poste'; else delete cabos[v.cabo].vinculoOrigem;
        });
        return r;
    }

    function vincularPorPosteBase(cabos, outros, cfg, prefixoFn, codigoFn, raio) {
        cfg = normalizarConfig(cfg);
        const ests = estruturasDisponiveis(outros), infoE = {}, infoC = {}, porLinha = new Map();
        ests.forEach(ie => { infoE[ie] = infoEstrutura(codigoFn(outros[ie].ativo, outros[ie]), cfg); const l = outros[ie]._linha; if (!porLinha.has(l)) porLinha.set(l, []); porLinha.get(l).push(ie); });
        cabos.forEach((c, ic) => { if (ehCaboValido(c)) infoC[ic] = infoCabo(prefixoFn(c.ativo) || '', c.operacao, cfg, extraCabo(c, cfg.dicionario_cabos)); });
        const linhasOrd = Array.from(porLinha.keys()).sort((a, b) => a - b);
        const postes = linhasOrd.map(l => [l, par(outros[porLinha.get(l)[0]]._poste)]).filter(([, pos]) => pos);
        const fict = {};
        ests.forEach(ie => { fict[ie] = (outros[ie].cabosFicticios || []).map(v => trim(v).toUpperCase()).filter(Boolean); });
        const icsOrd = Object.keys(infoC).map(Number).sort((a, b) => a - b);
        const manuais = new Set(icsOrd.filter(ic => cabos[ic].vinculoOrigem === 'manual' && (cabos[ic].vinculoEstruturas || []).length));
        const ocup = {}; ests.forEach(ie => { ocup[ie] = []; });
        Array.from(manuais).sort((a, b) => a - b).forEach(ic => cabos[ic].vinculoEstruturas.forEach(ie => { if (ie in ocup) ocup[ie].push(ic); }));
        const ligados = {}, restante = {}, chegadas = {}, soltas = [], semGeo = [];
        const saida = { vinculos: [], achados: [], resumo: {} };
        icsOrd.forEach(ic => {
            const c = infoC[ic];
            if (!c.vincula || manuais.has(ic) || cabos[ic].congelada) return;          // cabo congelado: só entra no vínculo se já foi ligado à mão
            const fim = extremos(cabos[ic]);
            if (!fim) { semGeo.push(ic); return; }
            ligados[ic] = [];
            restante[ic] = Math.max(0, c.pontas - estruturasFicticiasDe(cabos[ic]).length);
            fim.forEach((pt, k) => {
                const linha = posteMaisPerto(pt, postes, raio);
                if (linha === null) soltas.push([ic, k]); else (chegadas[linha] = chegadas[linha] || []).push(ic);
            });
        });
        const nomePoste = linha => `poste da linha ${linha + 1} (${outros[porLinha.get(linha)[0]].ativo})`;
        const add = (tipo, mensagem, poste, cabo, estrutura) => saida.achados.push({ tipo, poste, cabo, estrutura, mensagem });
        let estrIncompletas = 0;
        postes.forEach(([linha]) => {
            const pend = (chegadas[linha] || []).slice(), ies = porLinha.get(linha);
            const unidades = ies.slice().sort((a, b) => ((infoE[a].regra !== null ? infoE[a].regra : 1e6) - (infoE[b].regra !== null ? infoE[b].regra : 1e6)) || (ies.indexOf(a) - ies.indexOf(b)));
            unidades.forEach(ie => {
                const e = infoE[ie];
                if (e.regra === null || e.qtd === null || !e.ocupa) return;          // sem regra/quantidade (ou equipamento que não ocupa — TASK-169): não exige cabo
                const ocupantes = ocup[ie].map(ic => ({ info: infoC[ic], ficticio: false })).concat(fict[ie].map(cod => ({ info: infoCaboFicticio(cod, e.codigo, opFicticio(e), cfg), ficticio: true })));
                const falta = e.qtd - ocupantes.length;
                if (falta > 0) {
                    let cands = [];
                    pend.forEach(ic => { if ((restante[ic] || 0) > 0 && podeVincular(infoC[ic], e)[0] && !cands.includes(ic)) cands.push(ic); });
                    if (e.lados === 'IGUAIS' && cands.length && !ocupantes.length) {
                        const grupos = new Map();
                        cands.forEach(ic => { const k = infoC[ic].codigo + '|' + infoC[ic].op; if (!grupos.has(k)) grupos.set(k, []); grupos.get(k).push(ic); });
                        let melhor = null, mv = -1;
                        grupos.forEach(g => { const v = Math.min(g.length, falta); if (v > mv) { melhor = g; mv = v; } });          // o grupo de cabos iguais que mais enche a estrutura (empate: o 1º)
                        cands = melhor;
                    }
                    const escolhidos = [];
                    for (const ic of cands) {
                        if (escolhidos.length >= falta) break;
                        if (podeComOcupantes(infoC[ic], e, ocupantes)[0]) { ocupantes.push({ info: infoC[ic], ficticio: false }); escolhidos.push(ic); }
                    }
                    if (cands.length > escolhidos.length && escolhidos.length) add('escolha', `${e.codigo} no ${nomePoste(linha)}: ${cands.length} cabos servem, ligados os ${escolhidos.length} primeiros (ordem da tabela de cabos).`, linha, escolhidos[0], ie);
                    escolhidos.forEach(ic => { ocup[ie].push(ic); ligados[ic].push(ie); restante[ic] -= 1; pend.splice(pend.indexOf(ic), 1); });
                }
                const tem = ocup[ie].length + fict[ie].length;
                if (e.qtd_min !== null && tem < e.qtd_min) { estrIncompletas++; add('incompleta', `${e.codigo} no ${nomePoste(linha)}: precisa de ${e.qtd_min} cabo(s) e só chegaram ${tem} que ela aceita.`, linha, null, ie); }
            });
            pend.forEach(ic => { if ((restante[ic] || 0) > 0) add('sobra', `Cabo ${infoC[ic].codigo} (linha ${ic + 1}) chega ao ${nomePoste(linha)} e nenhuma estrutura o usou.`, linha, ic, null); });
        });
        soltas.forEach(([ic, k]) => add('ponta_solta', `Cabo ${infoC[ic].codigo} (linha ${ic + 1}): a ponta ${k === 0 ? 'A' : 'B'} não toca nenhum poste (raio ${raio}).`, null, ic, null));
        semGeo.forEach(ic => add('sem_geometria', `Cabo ${infoC[ic].codigo} (linha ${ic + 1}) sem geometria (_geo): vínculo não alterado.`, null, ic, null));
        let completos = 0;
        Object.keys(ligados).map(Number).sort((a, b) => a - b).forEach(ic => {
            const full = ligados[ic].length >= Math.max(0, infoC[ic].pontas - estruturasFicticiasDe(cabos[ic]).length);
            if (full && ligados[ic].length) completos++;
            saida.vinculos.push({ cabo: ic, estruturas: ligados[ic].slice(), completo: !!(full && ligados[ic].length) });
        });
        saida.resumo = { cabos: Object.keys(ligados).length, completos, parciais: saida.vinculos.filter(v => v.estruturas.length && !v.completo).length,
            sem_vinculo: saida.vinculos.filter(v => !v.estruturas.length).length, estruturas_incompletas: estrIncompletas, pontas_soltas: soltas.length, manuais_mantidos: manuais.size, sem_geometria: semGeo.length };
        return saida;
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
        cabos.forEach((c, ic) => { if (ehCaboValido(c)) { infoC[ic] = infoCabo(prefixoFn(c.ativo) || '', c.operacao, cfg, extraCabo(c, cfg.dicionario_cabos)); ligados[ic] = (c.vinculoEstruturas || []).slice(); } });
        const chavesC = Object.keys(infoC).map(Number).sort((a, b) => a - b);
        const saida = { vinculos: [], conflitos: [], sem_candidato: [], passos: [], parou: false };

        const ocupantesDe = (ie, ignorando) => {
            const lista = chavesC.filter(ic => ic !== ignorando && ligados[ic].includes(ie)).map(ic => ({ info: infoC[ic], ficticio: false }));
            fict[ie].forEach(cod => lista.push({ info: infoCaboFicticio(cod, infoE[ie].codigo, opFicticio(infoE[ie]), cfg), ficticio: true }));
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
                if ((ligados[ic].length || estruturasFicticiasDe(cabos[ic]).length) && !passo.substituir) {
                    // `completar`: o cabo que ficou com menos pontas do que precisa (ex.: parcial de um passo anterior) volta a ser alvo e só ganha o que falta
                    const falta = (Number.isInteger(passo.pontas) ? passo.pontas : c.pontas) - estruturasFicticiasDe(cabos[ic]).length - ligados[ic].length;
                    if (!(passo.completar && falta > 0)) { r.ja_tinham++; return; }
                }
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
            const candGlobal = ests.filter(ie => estruturaOk(ie) && infoE[ie].ocupa)          // TASK-169: equipamento (ocupa: false) não recebe cabo
                , raio = passo.raio_m === undefined ? null : passo.raio_m;
            for (let pos = 0; pos < alvo.length; pos++) {
                const ic = alvo[pos], c = infoC[ic];
                const need = Math.max(0, (Number.isInteger(passo.pontas) ? passo.pontas : c.pontas) - (passo.substituir ? 0 : estruturasFicticiasDe(cabos[ic]).length));
                const antes = ligados[ic].slice();
                if (passo.substituir) ligados[ic] = [];
                const base = ligados[ic].slice(), falta = Math.max(0, need - base.length);          // só ganha o que falta (base vazia, exceto em `completar`)
                const escolhidas = [], recusas = [];
                let candidatas;
                if (passo.modo === 'sequencia') candidatas = candGlobal.slice(pos, pos + need);
                else if (passo.modo === 'aleatorio') {
                    // sorteio entre as estruturas aceitas; por padrão as com MENOS cabos primeiro (sorteio entre as vazias antes das ocupadas); `puro: true` = sorteio simples
                    const rnd = mulberry32(((Number.isInteger(passo.semente) ? passo.semente : 1) * 1000003 + ic * 7919 + ip * 104729) >>> 0);
                    candidatas = embaralhar(candGlobal.filter(ie => !base.includes(ie)), rnd);
                    if (!passo.puro) candidatas = candidatas.map((ie, k) => [ie, k, ocupantesDe(ie).length]).sort((a, b) => (a[2] - b[2]) || (a[1] - b[1])).map(x => x[0]);
                } else {
                    if (!temXY(cabos[ic])) {
                        r.sem_candidato++; saida.sem_candidato.push({ passo: ip, cabo: ic, motivo: 'cabo sem coordenada' }); ligados[ic] = antes; continue;
                    }
                    candidatas = candGlobal.filter(ie => temXY(outros[ie]) && (raio === null || dist(cabos[ic], outros[ie]) <= raio))
                        .sort((a, b) => (dist(cabos[ic], outros[a]) - dist(cabos[ic], outros[b])) || (a - b));
                }
                for (const ie of candidatas) {
                    if (escolhidas.length >= falta) break;
                    if (base.includes(ie)) continue;
                    const [ok, motivo] = pode(ic, ie);
                    if (ok && outros[ie]._linha !== undefined && base.concat(escolhidas).some(e2 => outros[e2]._linha === outros[ie]._linha)) continue;      // TASK-094: 2 estruturas do mesmo poste não valem como as 2 pontas
                    if (ok) { escolhidas.push(ie); ligados[ic].push(ie); } else recusas.push(motivo);
                }
                const completo = need > 0 && base.length + escolhidas.length === need;
                if (completo || (escolhidas.length && passo.aceitar_parcial)) {
                    saida.vinculos.push({ passo: ip, cabo: ic, estruturas: base.concat(escolhidas), completo });
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
        const erros = (DC ? DC.validar(cfg.dicionario_cabos) : []).concat(CP ? CP.validar(cfg.postes) : []);          // TASK-173/174
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
                    if (r.parametros != null && r.parametros !== '') {
                        const np = normalizarParametros(r.parametros), brutos = Array.isArray(r.parametros) ? r.parametros.length : String(r.parametros).split(',').length;
                        if (!np || np.length !== brutos || np.some(x => !/^[A-Z0-9_.]+$/.test(x.nome) || !FONTES_PARAMETRO.includes(x.fonte))) erros.push(`${p}: 'parametros' precisa ser uma lista como "cabo=ativo, cabo=ativo, tipocabo=valor_cabo" (fonte: descricao, valor_cabo ou ativo).`);
                    }
                    if (r.lados != null && !['IGUAIS', 'LIVRE', 'DERIVACAO'].includes(r.lados)) erros.push(`${p}: 'lados' precisa ser IGUAIS, LIVRE ou DERIVACAO.`);
                    if (r.cabos_vaos != null && r.cabos_vaos !== '') {                                                                // TASK-170
                        if (!Array.isArray(r.cabos_vaos)) erros.push(`${p}: 'cabos_vaos' precisa ser uma lista de regex (uma por vão).`);
                        else r.cabos_vaos.forEach(x => { let m = null; try { new RegExp(String(x)); } catch (e) { m = e.message; } if (m) erros.push(`${p}: 'cabos_vaos' tem regex inválida (${m}).`); });
                    }
                    if (r.define != null && r.define !== '') {
                        if (!Array.isArray(r.define)) erros.push(`${p}: 'define' precisa ser uma lista de {linha_tem?, cabo?, param}.`);
                        else r.define.forEach((it, k) => {
                            const q = `${p}: define #${k + 1}`;
                            if (!it || typeof it !== 'object') { erros.push(`${q}: precisa ser um objeto.`); return; }
                            if ((it.linha_tem == null || it.linha_tem === '' || (Array.isArray(it.linha_tem) && !it.linha_tem.length)) && !trim(it.cabo)) erros.push(`${q}: precisa de pelo menos uma condição (linha_tem ou cabo).`);
                            if (typeof it.param !== 'string' || !/^\s*[A-Za-z0-9_.]+\s*=\s*[A-Za-z0-9_./-]+\s*$/.test(it.param)) erros.push(`${q}: 'param' precisa ser nome=valor (ex.: tipo=RL).`);
                            ['linha_tem', 'cabo'].forEach(campo => { const v = it[campo]; if (typeof v === 'string' && RE_REGEX.test(v)) { let m = null; try { new RegExp(v); } catch (e) { m = e.message; } if (m) erros.push(`${q}: '${campo}' regex inválida (${m}).`); } });
                        });
                    }
                    if (r.se_diferem != null && !SE_DIFEREM.includes(r.se_diferem)) erros.push(`${p}: 'se_diferem' precisa ser um de ${JSON.stringify(SE_DIFEREM)}.`);
                    ['ocupa', 'distintos'].forEach(k => { if (r[k] != null && typeof r[k] !== 'boolean') erros.push(`${p}: '${k}' precisa ser true ou false.`); });          // TASK-169
                    if (r.herda != null && r.herda !== '' && r.herda !== 'POSTE') erros.push(`${p}: 'herda' precisa ser "POSTE" (ou vazio).`);
                    if (r.herda === 'POSTE' && r.ocupa !== false) erros.push(`${p}: 'herda' só vale com 'ocupa': false.`);
                    if (trim(r.nivel) && !['MT', 'BT'].includes(trim(r.nivel).toUpperCase())) erros.push(`${p}: 'nivel' precisa ser MT ou BT.`);
                    if (r.ordem != null && r.ordem !== '' && !ORDENS.includes(r.ordem)) erros.push(`${p}: 'ordem' precisa ser um de ${JSON.stringify(ORDENS)}.`);
                    if (r.se_faltar != null && r.se_faltar !== '' && !SE_FALTAR.includes(r.se_faltar)) erros.push(`${p}: 'se_faltar' precisa ser um de ${JSON.stringify(SE_FALTAR)}.`);
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

    const api = { parametrosDoCaboNaEstrutura, defineDeEstruturas, vincularPorPoste, RAIO_POSTE_PADRAO, descricaoCabo, normalizarParametros, parametrosPorEstrutura, participaVinculo, estruturasFicticiasDe, idUnidade, linhaDe, slotDe, ficticiosDe, codigosDaLinha, derivarCabo, extraCabo, opFicticio, validarConfig, casa, configVazia, normalizarConfig, opBase, infoCabo, infoEstrutura, podeVincular, podeComOcupantes, estruturasDisponiveis, validarScript, executarScript, analisar, parametrosPorLinha, achados, ehCaboValido };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.VinculacaoRegras = api;
})(typeof window !== 'undefined' ? window : globalThis);
