/*
 * dicionario_cabos.js (TASK-173, ADR-007) — dicionário de cabos por projeto. Sem DOM.
 * MESMA lógica de services/dicionario_cabos.py (paridade testada em tests/test_dicionario_cabos_173.py).
 * Sem dicionário (lista vazia), as regras fixas de Pernambuco continuam valendo (quem chama decide o padrão).
 */
(function (raiz) {
    'use strict';
    const NEUTRO_PADRAO = 'CAZ{valor}';
    const trim = t => String(t == null ? '' : t).trim();
    function rx(p) { try { return new RegExp(trim(p), 'i'); } catch (e) { return null; } }
    function normalizar(dic) { return Array.isArray(dic) ? dic.filter(r => r && typeof r === 'object' && !Array.isArray(r)).map(r => Object.assign({}, r)) : []; }

    function validar(dic) {
        const erros = [];
        if (dic === undefined || dic === null) return erros;
        if (!Array.isArray(dic)) return ["'dicionario_cabos' precisa ser uma lista de linhas."];
        dic.forEach((r, n) => {
            const p = `Dicionário de cabos, linha ${n + 1}`;
            if (!r || typeof r !== 'object' || Array.isArray(r)) { erros.push(`${p}: precisa ser um objeto.`); return; }
            if (!trim(r.padrao)) erros.push(`${p}: falta o 'padrao' (regex do código do cabo).`);
            else if (!rx(r.padrao)) erros.push(`${p}: 'padrao' não é uma regex válida.`);
            if (trim(r.nivel) && !['MT', 'BT'].includes(trim(r.nivel).toUpperCase())) erros.push(`${p}: 'nivel' precisa ser MT, BT ou vazio.`);
            const fp = r.fases_padrao;
            if (fp !== undefined && fp !== null && fp !== '' && !(Number.isInteger(fp) && fp >= 1)) erros.push(`${p}: 'fases_padrao' precisa ser um número inteiro maior que zero.`);
            const s = r.secao;
            if (s !== undefined && s !== null && s !== '' && typeof s !== 'number' && !/^\$[0-9]$/.test(trim(s))) erros.push(`${p}: 'secao' precisa ser um número ou $1…$9 (grupo da regex).`);
            if (r.apelidos !== undefined && r.apelidos !== null && !Array.isArray(r.apelidos)) erros.push(`${p}: 'apelidos' precisa ser uma lista.`);
            (Array.isArray(r.apelidos) ? r.apelidos : []).forEach(a => {
                const pa = a && typeof a === 'object' ? a.padrao : a;
                if (!trim(pa) || !rx(pa)) erros.push(`${p}: apelido '${pa}' não é uma regex válida.`);
            });
            if (r.notacao_solta !== undefined && r.notacao_solta !== null && typeof r.notacao_solta !== 'boolean') erros.push(`${p}: 'notacao_solta' precisa ser true ou false.`);
            if (r.neutro !== undefined && r.neutro !== null && r.neutro !== '' && !trim(r.neutro).includes('{valor}')) erros.push(`${p}: 'neutro' precisa conter {valor} (ex.: CAZ{valor}).`);
            const aw = r.awg;
            if (aw !== undefined && aw !== null && aw !== '' && !(typeof aw === 'object' && !Array.isArray(aw) && Object.values(aw).every(v => typeof v === 'number'))) erros.push(`${p}: 'awg' precisa ser {"4": 21.2, "1/0": 53.5}.`);
        });
        return erros;
    }

    /** [índice, linha, match] da primeira linha ativa cujo `padrao` casa o código; [null, null, null] se nenhuma. */
    function linha(dic, codigo) {
        const cod = trim(codigo).toUpperCase();
        if (!cod) return [null, null, null];
        const l = normalizar(dic);
        for (let i = 0; i < l.length; i++) {
            if (l[i].ativa === false) continue;
            const r = rx(l[i].padrao), m = r ? r.exec(cod) : null;
            if (m) return [i, l[i], m];
        }
        return [null, null, null];
    }
    function secao(dic, codigo) {
        const [, r, m] = linha(dic, codigo);
        if (!r) return null;
        const s = r.secao;
        if (s === undefined || s === null || s === '') return null;
        if (typeof s === 'number') return s;
        const g = /^\$([0-9])$/.exec(trim(s));
        if (!g || Number(g[1]) > m.length - 1 || m[Number(g[1])] === undefined) return null;
        const txt = m[Number(g[1])].toUpperCase(), aw = r.awg && typeof r.awg === 'object' ? r.awg : {};
        if (Object.prototype.hasOwnProperty.call(aw, txt)) return Number(aw[txt]);
        const n = Number(txt.replace(',', '.'));
        return /^\s*[+-]?(\d+\.?\d*|\.\d+)\s*$/.test(txt.replace(',', '.')) ? n : null;
    }
    function valorCabo(dic, codigo) { const [, r] = linha(dic, codigo); return r ? (trim(r.valor_cabo).toUpperCase() || null) : null; }
    function nivelPadrao(dic, codigo) { const [, r] = linha(dic, codigo); const n = r ? trim(r.nivel).toUpperCase() : ''; return n === 'MT' || n === 'BT' ? n : null; }
    function fasesPadrao(dic, codigo) { const [, r] = linha(dic, codigo); const f = r ? r.fases_padrao : null; return Number.isInteger(f) && f >= 1 ? f : null; }
    function neutroTemplate(dic, codigo) { const [, r] = linha(dic, codigo); const t = r ? trim(r.neutro) : ''; return t && t.includes('{valor}') ? t : NEUTRO_PADRAO; }

    /** Se o texto é apelido do dicionário: [true, código]; senão [false, null]. */
    function apelido(dic, rotulo) {
        const t = trim(rotulo).toUpperCase().replace(/ /g, '');
        if (!t) return [false, null];
        for (const r of normalizar(dic)) {
            if (r.ativa === false || !Array.isArray(r.apelidos)) continue;
            for (const a of r.apelidos) {
                const pa = a && typeof a === 'object' ? a.padrao : a, re = rx(pa), m = re ? re.exec(t) : null;
                if (!m) continue;
                const cod = a && typeof a === 'object' ? trim(a.codigo).toUpperCase() : '';
                if (!cod) return [true, t];
                return [true, cod.replace(/\$([0-9])/g, (_, g) => (Number(g) <= m.length - 1 ? (m[Number(g)] || '') : '')).toUpperCase()];
            }
        }
        return [false, null];
    }
    function reconheceCodigo(dic, codigo, solta) { const [, r] = linha(dic, codigo); return r !== null && (!solta || r.notacao_solta === true); }
    function ordemMaiorPadrao(dic) {
        const ordem = [];
        normalizar(dic).forEach(r => { const v = trim(r.valor_cabo).toUpperCase(); if (v && !ordem.includes(v)) ordem.push(v); });
        return ordem;
    }

    const api = { normalizar, validar, linha, secao, valorCabo, nivelPadrao, fasesPadrao, neutroTemplate, apelido, reconheceCodigo, ordemMaiorPadrao, NEUTRO_PADRAO };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.DicionarioCabos = api;
})(typeof window !== 'undefined' ? window : globalThis);
