/*
 * anotacao_cabos.js (TASK-083) — notação do LABEL na tabela Cabos: "3xA35P(7,94)-13,8kV 45 m". Sem DOM.
 * MESMA lógica de services/anotacao_cabos.py (paridade testada em tests/test_anotacao_cabos.py).
 * Formato antigo (CAA2 ABC 35 m) continua como sempre; a tradução da notação é implícita (a célula nunca é reescrita).
 */
(function (raiz) {
    'use strict';
    const RE_SEPARA = /^(.*\S)\s+(\d+(?:[.,]\d+)?)\s*[mM]?$/;
    const RE_TENSAO = /\s*[-+]\s*(\d+(?:[.,]\d+)?)\s*KV\s*$/i;      // 13,8kV ou 13.8KV
    const RE_SOLTO = /^[(\[]\s*([A-Z]\w*)\s*[)\]]$/;
    const RE_POTENCIA = /^C\s*\d[xX]/;
    const RE_CODIGO = /^(?:(\d)[xX])?([A-Z]{0,4}\d[\dA-Z/,.]*)(?:\(([^)]*)\))?$/;
    const RE_SECAO = /^([A-Z]*)(\d+)([A-Z]*)$/;
    const RE_CURTO = /^(?:RAMA|PB)\d+$/;
    const RE_NUMERO_SOZINHO = /^\d{2,3}$/;
    const RE_MARCA = /^\d[xX]|^C\s*\d[xX]|[(\[]|KV/i;

    function separar(texto) {
        const t = String(texto == null ? '' : texto).replace(/\s+/g, ' ').trim();
        const m = t.match(RE_SEPARA);
        return m ? [m[1].trim(), m[2]] : [t, null];
    }

    function ehNotacao(texto) {
        const r = separar(texto)[0].toUpperCase();
        if (!r) return false;
        return RE_MARCA.test(r) || RE_CURTO.test(r) || RE_NUMERO_SOZINHO.test(r);
    }

    const num = t => parseFloat(String(t).replace(',', '.'));

    function neutroDe(codigo, valor) {
        const v = valor.trim().toUpperCase();
        if (!v) return null;
        if (v.includes(',')) return 'CAZ' + v;
        if (/^\d+$/.test(v)) {
            const m = codigo.match(RE_SECAO);
            return m ? `${m[1]}${v}${m[3]}` : codigo;
        }
        return v;
    }

    function interpretar(texto) {
        const r = { formato: 'antigo', itens: [], tensao: null, erros: [] };
        if (!ehNotacao(texto)) return r;
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
        else if (RE_NUMERO_SOZINHO.test(rotulo)) add('C' + rotulo, 1, false);
        else {
            const c = rotulo.replace(/ /g, '').match(RE_CODIGO);
            if (!c) { r.erros.push(`rótulo '${rotulo}' não reconhecido (ex.: 3xA35P(7,94)-13,8kV 45 m)`); return r; }
            const fases = c[1] ? parseInt(c[1], 10) : 1, codigo = c[2];
            add(codigo, fases, false);
            if (c[3] !== undefined) {
                const neutro = neutroDe(codigo, c[3]);
                if (neutro === null) r.erros.push(`neutro vazio em '${rotulo}'`);
                else add(neutro, 1, true);
            }
        }
        return r;
    }

    function principal(r) {
        const p = r.itens.filter(i => !i.neutro);
        return (p.length ? p : r.itens)[0] || null;
    }

    function fasesDaLinha(texto) {
        const p = principal(interpretar(texto));
        return p ? p.fases : 1;
    }

    function prefixoEComprimento(texto) {
        const p = principal(interpretar(texto));
        return p && p.comprimento !== null ? [p.ativo, p.comprimento] : null;
    }

    const api = { ehNotacao, interpretar, fasesDaLinha, prefixoEComprimento };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.AnotacaoCabos = api;
})(typeof window !== 'undefined' ? window : globalThis);
