/*
 * anotacao_cabos.js (TASK-083) — notação do LABEL na tabela Cabos: "3xA35P(7,94)-13,8kV 45 m". Sem DOM.
 * MESMA lógica de services/anotacao_cabos.py (paridade testada em tests/test_anotacao_cabos.py).
 * Formato antigo (CAA2 ABC 35 m) continua como sempre; a tradução da notação é implícita (a célula nunca é reescrita).
 */
(function (raiz) {
    'use strict';
    const DC = (typeof module !== 'undefined' && module.exports) ? require('./dicionario_cabos.js') : raiz.DicionarioCabos;   // TASK-173
    let dicionarioAtual = null;          // TASK-173: dicionário do projeto aberto na tela (o resumo.js define ao carregar a configuração); parâmetro `dic` explícito sempre vale mais
    function definirDicionario(dic) { dicionarioAtual = Array.isArray(dic) && dic.length ? dic : null; }
    const dicDe = dic => (dic === undefined ? dicionarioAtual : dic);
    const RE_SEPARA = /^(.*\S)\s+(\d+(?:[.,]\d+)?)\s*[mM]?$/;
    const RE_TENSAO = /\s*[-+]\s*(\d+(?:[.,]\d+)?)\s*KV\s*$/i;      // 13,8kV ou 13.8KV
    const RE_SOLTO = /^[(\[]\s*([A-Z]\w*)\s*[)\]]$/;
    const RE_EXTRA = /^(.*\S)\s*\+\s*(\d+(?:[.,]\d+)?(?:\/\d+)?)$/;          // "3xC16(16)+16": o "+16" é mais um condutor da mesma família (sem kV: com kV é a tensão)
    const RE_POTENCIA = /^C\s*\d[xX]/;
    const RE_CODIGO = /^(?:(\d)[xX])?([A-Z]{0,4}\d[\dA-Z/,.]*)(?:\(([^)]*)\))?$/;
    const RE_SECAO = /^([A-Z]*)(\d+)([A-Z]*)$/;
    const RE_SECAO_AWG = /^([A-Z]*)(\d+\/\d+)([A-Z]*)$/;               // TASK-251: A1/0, A4/0
    const RE_CURTO = /^(?:RAMA|PB)\d+$/;
    const RE_NUMERO_SOZINHO = /^\d{2,3}$/;
    const RE_MARCA = /^\d[xX]|^C\s*\d[xX]|[(\[]|KV/i;

    function separar(texto) {
        const t = String(texto == null ? '' : texto).replace(/\s+/g, ' ').trim();
        const m = t.match(RE_SEPARA);
        return m ? [m[1].trim(), m[2]] : [t, null];
    }

    function semTensao(r) { const m = r.match(RE_TENSAO); return m ? r.slice(0, m.index).trim() : r; }
    // TASK-173: com dicionário, o código do rótulo ("3xALU-50(35)" → "ALU-50") é qualquer texto que uma linha do dicionário reconheça.
    const RE_CODIGO_DIC = /^(?:(\d)[xX])?([^()\s]+)(?:\(([^)]*)\))?$/;
    function porDicionario(rotulo, dic, solta) {
        if (!dic || !DC) return null;
        const c = rotulo.replace(/ /g, '').match(RE_CODIGO_DIC);
        return c && DC.reconheceCodigo(dic, c[2], !!solta) ? c : null;
    }

    function ehNotacao(texto, dic) {
        dic = dicDe(dic);
        let r = separar(texto)[0].toUpperCase();
        if (!r) return false;
        if (RE_MARCA.test(r) || RE_CURTO.test(r) || RE_NUMERO_SOZINHO.test(r)) return true;
        if (dic && DC) { r = semTensao(r); return DC.apelido(dic, r)[0] || porDicionario(r, dic, true) !== null; }
        return false;
    }

    const num = t => parseFloat(String(t).replace(',', '.'));

    function neutroDe(codigo, valor, dic) {
        const v = valor.trim().toUpperCase();
        if (!v) return null;
        if (v.includes(',')) return dic && DC ? DC.neutroTemplate(dic, codigo).replace('{valor}', v) : 'CAZ' + v;
        if (/^\d+$/.test(v)) {
            const m = codigo.match(RE_SECAO);
            return m ? `${m[1]}${v}${m[3]}` : codigo;
        }
        if (/^\d+\/\d+$/.test(v)) {                                   // TASK-251: seção AWG → mesma família do cabo; cabo que não é AWG: valor literal
            const m = codigo.match(RE_SECAO_AWG);
            return m ? `${m[1]}${v}${m[3]}` : v;
        }
        return v;
    }

    function interpretar(texto, dic) {
        dic = dicDe(dic);
        const r = { formato: 'antigo', itens: [], tensao: null, erros: [] };
        if (!ehNotacao(texto, dic)) return r;
        r.formato = 'notacao';
        let [rotulo, comp] = separar(texto);
        rotulo = rotulo.toUpperCase();
        const comprimento = comp !== null ? num(comp) : null;
        if (comprimento === null) r.erros.push(`falta o comprimento em '${rotulo}' (use: rótulo comprimento m)`);
        const mt = rotulo.match(RE_TENSAO);
        if (mt) { r.tensao = mt[1]; rotulo = rotulo.slice(0, mt.index).trim(); }
        const add = (ativo, fases, neutro) => r.itens.push({ ativo, fases, comprimento, neutro });

        const solto = rotulo.match(RE_SOLTO);
        if (solto) add(solto[1], 1, true);
        else if (RE_POTENCIA.test(rotulo)) add(rotulo.replace(/\s+/g, ''), 1, false);
        else if (dic && DC && DC.apelido(dic, rotulo)[0]) { const ap = DC.apelido(dic, rotulo)[1]; add(ap, DC.fasesPadrao(dic, ap) || 1, false); }
        else if (RE_NUMERO_SOZINHO.test(rotulo)) add('C' + rotulo, 1, false);
        else {
            let c = null, extra = null;
            const me = rotulo.match(RE_EXTRA);
            if (me) { c = me[1].replace(/ /g, '').match(RE_CODIGO) || porDicionario(me[1].trim(), dic); if (c) extra = me[2]; }
            if (!c) c = rotulo.replace(/ /g, '').match(RE_CODIGO) || porDicionario(rotulo, dic);
            if (!c) { r.erros.push(`rótulo '${rotulo}' não reconhecido (ex.: 3xA35P(7,94)-13,8kV 45 m)`); return r; }
            const codigo = c[2], fases = c[1] ? parseInt(c[1], 10) : ((dic && DC ? DC.fasesPadrao(dic, codigo) : null) || 1);
            add(codigo, fases, false);
            if (c[3] !== undefined) {
                const neutro = neutroDe(codigo, c[3], dic);
                if (neutro === null) r.erros.push(`neutro vazio em '${rotulo}'`);
                else add(neutro, 1, true);
            }
            if (extra !== null) {                                                         // "+16": mais um condutor (seção 16) da mesma família — 3xC16(16)+16 = 3 fases + neutro + 1 = 5 ativos
                const ex = neutroDe(codigo, extra, dic);
                if (ex === null) r.erros.push(`condutor extra vazio em '${rotulo}'`); else add(ex, 1, true);
            }
        }
        return r;
    }

    function principal(r) {
        const p = r.itens.filter(i => !i.neutro);
        return (p.length ? p : r.itens)[0] || null;
    }

    function fasesDaLinha(texto, dic) {
        const p = principal(interpretar(texto, dic));
        return p ? p.fases : 1;
    }

    function prefixoEComprimento(texto, dic) {
        const p = principal(interpretar(texto, dic));
        return p && p.comprimento !== null ? [p.ativo, p.comprimento] : null;
    }

    /** TASK-116: linha que traz SÓ o comprimento ("45", "45 m", "7,5M"): vale para o cabo da linha de cima que ficou sem comprimento. */
    function comprimentoSozinho(texto) {
        const m = String(texto == null ? '' : texto).trim().match(/^(\d+(?:[.,]\d+)?)\s*[mM]?$/);
        return m ? num(m[1]) : null;
    }

    const api = { definirDicionario, dicionario: () => dicionarioAtual, ehNotacao, interpretar, fasesDaLinha, prefixoEComprimento, comprimentoSozinho };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.AnotacaoCabos = api;
})(typeof window !== 'undefined' ? window : globalThis);
