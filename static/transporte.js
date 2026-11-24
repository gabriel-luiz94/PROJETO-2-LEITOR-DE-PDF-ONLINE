/*
 * transporte.js (TASK-245) — transporte de material (TRANSP / TRANSP_KM) igual à aba TRANSPORT da ICP.
 * MESMA função de services/transporte.py (paridade testada em tests/test_transporte_245.py). A parte de cálculo não usa DOM;
 * o modal (abrirModal) só existe no navegador. Não "consertar" as bordas da ICP: a regra é dar o mesmo resultado que ela.
 */
(function (raiz) {
    'use strict';

    // #VALOR! do Excel: propaga em contas, SOMASE e comparações. Conta como TEXTO (como o _Erro(str) do Python).
    const ERRO = { toString() { return '#VALUE!'; } };
    const ehErro = x => x === ERRO;
    const algumErro = (...xs) => xs.some(ehErro);
    const ehTexto = x => typeof x === 'string' || ehErro(x);
    const ehNum = x => typeof x === 'number';
    class Propaga extends Error {}

    const BANDAS_ATE50 = ['SDETU5001SI', 'SDETU5003SI', 'SDETU5005SI', 'SDETU5007SI', 'SDETU5009SI'];
    const BANDAS_MAIS50 = ['SDETU5002SI', 'SDETU5004SI', 'SDETU5006SI', 'SDETU5008SI', 'SDETU5010SI'];

    /** Comparação do Excel: texto é maior que qualquer número; "" (vazio de fórmula) também é texto. */
    function gt(a, b) {
        if (algumErro(a, b)) throw new Propaga();
        if (ehTexto(a) && !ehTexto(b)) return true;
        if (!ehTexto(a) && ehTexto(b)) return false;
        return a > b;
    }
    const ge = (a, b) => gt(a, b) || a === b;
    const le = (a, b) => !gt(a, b);
    const lt = (a, b) => !ge(a, b);

    /** Texto decimal exato (sem notação científica) de um número com 15 algarismos significativos. */
    function decimal15(x) {
        const s = Math.abs(x).toPrecision(15);
        const m = s.match(/^(\d)(?:\.(\d+))?e([+-]\d+)$/);
        if (!m) return s;
        const dig = m[1] + (m[2] || ''), exp = parseInt(m[3], 10);
        if (exp >= 0) return dig.length > exp + 1 ? dig.slice(0, exp + 1) + '.' + dig.slice(exp + 1) : dig + '0'.repeat(exp + 1 - dig.length);
        return '0.' + '0'.repeat(-exp - 1) + dig;
    }
    /** ARRED do Excel: meio para longe do zero, sobre o número com 15 algarismos significativos (0,185 → 0,19). */
    function arred(x, casas) {
        casas = casas === undefined ? 2 : casas;
        const neg = x < 0, s = decimal15(x);
        let [int, frac = ''] = s.split('.');
        frac = frac.padEnd(casas + 1, '0');
        const sobe = frac[casas] >= '5';
        let digitos = (int + frac.slice(0, casas)).split('').map(Number);
        if (sobe) {
            let i = digitos.length - 1;
            while (i >= 0) { if (digitos[i] === 9) { digitos[i] = 0; i--; } else { digitos[i]++; break; } }
            if (i < 0) digitos.unshift(1);
        }
        const txt = digitos.join(''), corte = txt.length - casas;
        const v = Number(casas ? txt.slice(0, corte) + '.' + txt.slice(corte) : txt);
        return neg && v !== 0 ? -v : v;
    }

    const num = x => (ehNum(x) ? x : 0);

    /** T3 (carga total O5, até 50 km). Repare no 4.001: entre 4 e 4,001 t não sai serviço. */
    function bandaPrimeira(o5) {
        if (!gt(o5, 0.0001)) return '';
        if (gt(o5, 0.001) && le(o5, 0.6)) return 'SDETU5001SI';
        if (gt(o5, 0.6) && le(o5, 4)) return 'SDETU5003SI';
        if (gt(o5, 4.001) && le(o5, 8)) return 'SDETU5005SI';
        if (gt(o5, 8) && le(o5, 14)) return 'SDETU5007SI';
        if (gt(o5, 14) && le(o5, 25)) return 'SDETU5009SI';
        if (gt(o5, 25)) return 'SDETU5009SI';
        return false;                                   // o IF mais interno sem "senão" devolve FALSO
    }

    /** W3/Z3/AC3/...: bandas das cargas seguintes. Aqui 25 exato não cai em nenhuma (V3<25 / V3>25). */
    function bandaResto(v) {
        if (v === '') return '';
        if (!gt(v, 0.001)) return '';
        if (gt(v, 0.001) && le(v, 0.6)) return 'SDETU5001SI';
        if (gt(v, 0.6) && le(v, 4)) return 'SDETU5003SI';
        if (gt(v, 4) && le(v, 8)) return 'SDETU5005SI';
        if (gt(v, 8) && le(v, 14)) return 'SDETU5007SI';
        if (gt(v, 14) && lt(v, 25)) return 'SDETU5009SI';
        if (gt(v, 25)) return 'SDETU5009SI';
        return false;
    }

    /** T4/W4/...: serviço acima de 50 km para a carga x. `condV` = célula usada no último ramo (>25). */
    function bandaKm(x, r3, condV) {
        if (condV === undefined) condV = x;
        if (le(x, 0.6) && ge(r3, 50.001)) return 'SDETU5002SI';
        if (ge(x, 0.6001) && le(x, 4) && ge(r3, 50.001)) return 'SDETU5004SI';
        if (gt(x, 4) && le(x, 8) && ge(r3, 50.001)) return 'SDETU5006SI';
        if (gt(x, 8) && le(x, 14) && ge(r3, 50.001)) return 'SDETU5008SI';
        if (gt(x, 14) && le(x, 25) && ge(r3, 50.001)) return 'SDETU5010SI';
        if (gt(condV, 25) && gt(r3, 50)) return 'SDETU5010SI';
        return '';
    }

    /**
     * qtdOdi / extra: {codigo: quantidade}; km: KMx2 da UTD ou null (município/UTD não informado = H15 vazio na ICP);
     * pesos: [[codigo, descricao, kg]] na ordem da aba. Devolve {toneladas, km, servicos: [[codigo, qtd 2 casas] | ['#VALUE!', '#VALUE!']]}.
     */
    function calcular(qtdOdi, extra, km, pesos, h15) {
        if (h15 === undefined || h15 === null) h15 = km !== null && km !== undefined;
        qtdOdi = qtdOdi || {}; extra = extra || {};
        // ---- O5: toneladas (H = E*G + F*G; O5 = soma/1000) ----
        let h = 0;
        for (const [cod, , kg] of pesos) h += num(qtdOdi[String(cod)] || 0) * kg + num(extra[String(cod)] || 0) * kg;
        const o5 = h / 1000;
        // ---- R3, R4, O6 ----
        const r3 = h15 && km !== null && km !== undefined ? km : '-';
        const s1 = h15 && !ehTexto(r3) ? r3 - 50 : (!h15 ? '' : '#VALOR');
        const r4 = !h15 ? '-' : (gt(r3, 50) ? s1 : '');
        const o6 = r3;
        // ---- cadeia de cargas de 25 t ----
        const resto = prev => (prev === '' ? '' : (gt(prev, 25) ? prev - 25 : ''));
        const v3 = gt(o5, 25) ? o5 - 25 : '';
        const y3 = resto(v3), ab3 = resto(y3), ae3 = resto(ab3), ah3 = resto(ae3), ak3 = resto(ah3), an3 = resto(ak3);
        const t3 = bandaPrimeira(o5);
        const [w3, z3, ac3, af3, ai3, al3] = [v3, y3, ab3, ae3, ah3, ak3].map(bandaResto);
        const t4 = bandaKm(o5, r3);
        const w4 = v3 !== '' ? bandaKm(v3, r3) : '';
        const z4 = y3 !== '' ? bandaKm(y3, r3, v3) : '';          // Z4 usa V3 no último ramo
        const ac4 = ab3 !== '' ? bandaKm(ab3, r3, y3) : '';       // AC4 usa Y3
        const af4 = ae3 !== '' ? bandaKm(ae3, r3, ab3) : '';      // AF4 usa AB3
        const ai4 = ah3 !== '' ? bandaKm(ah3, r3, ae3) : '';      // AI4 usa AE3
        const al4 = ak3 !== '' ? bandaKm(ak3, r3, ak3) : '';
        // ---- O10 (ton x km excedente) ----
        let o10;
        if (!h15) o10 = '-';
        else if (lt(o6, 50)) o10 = 'DISTÂNCIA INFERIOR A 50 Km';
        else if (gt(r3, 50) || gt(o5, 25)) o10 = !ehTexto(r4) ? o5 * r4 : ERRO;
        else o10 = '';
        // ---- quantidades por carga (coluna V) ----
        const loadQty = (x, prox) => {
            if (x === '') return '';
            if (gt(x, 0.001) && lt(x, 25)) return x;
            if (gt(x, 25)) return prox !== '' ? x - prox : ERRO;
            return '';
        };
        const v13 = le(o5, 25) ? o5 : o5 - v3;
        const v15 = loadQty(v3, y3), v17 = loadQty(y3, ab3), v19 = loadQty(ab3, ae3), v21 = loadQty(ae3, ah3);
        const v23 = loadQty(ah3, ak3), v25 = loadQty(ak3, an3);
        const mult = (a, b) => (ehTexto(a) || ehTexto(b) ? ERRO : a * b);   // texto ("" ou "-") x número = #VALOR!
        let t10, v10, y10, ab10, af10, ai10, al10;
        if (!h15) t10 = v10 = y10 = ab10 = af10 = ai10 = al10 = '';
        else {
            t10 = lt(o6, 50) ? '' : (v3 !== '' ? mult(v13, r4) : '');
            v10 = lt(o6, 50) ? '' : (v3 !== '' ? mult(v15, r4) : '');
            y10 = y3 !== '' ? mult(v17, r4) : '';
            ab10 = ab3 !== '' ? mult(v19, r4) : '';
            af10 = ae3 !== '' ? mult(v21, r4) : '';
            ai10 = ah3 !== '' ? mult(v23, r4) : '';
            al10 = ak3 !== '' ? mult(v25, r4) : '';
        }
        const linhas = [];                              // (U, V) linhas 13..26
        linhas.push([t3 !== '' ? t3 : '', v13]);
        const u14 = t4 !== '' && gt(o5, 0.001) ? t4 : '';
        const v14 = t4 !== '' ? (lt(o5, 25) ? o10 : t10) : '';
        linhas.push([u14, v14]);
        linhas.push([v3 !== '' ? w3 : '', v15]);
        const u16 = w4 !== '' && v3 !== '' && gt(v3, 0.001) ? w4 : '';
        linhas.push([u16, w4 !== '' ? v10 : '']);
        linhas.push([z3 !== '' ? z3 : '', v17]); linhas.push([z4, z4 !== '' ? y10 : '']);
        linhas.push([ac3 !== '' ? ac3 : '', v19]); linhas.push([ac4, ac4 !== '' ? ab10 : '']);
        linhas.push([af3 !== '' ? af3 : '', v21]); linhas.push([af4, af4 !== '' ? af10 : '']);
        linhas.push([ai3 !== '' ? ai3 : '', v23]); linhas.push([ai4, ai4 !== '' ? ai10 : '']);
        linhas.push([al3 !== '' ? al3 : '', v25]); linhas.push([al4, al4 !== '' ? al10 : '']);
        // ---- AF: SOMASE por código (erro numa célula que casa = erro) ----
        const af = {};
        for (const cod of BANDAS_ATE50.concat(BANDAS_MAIS50)) {
            const casam = linhas.filter(([u]) => u === cod).map(([, v]) => v);
            af[cod] = algumErro(...casam) ? ERRO : casam.filter(ehNum).reduce((s, v) => s + v, 0);
        }
        const ac23 = algumErro(...BANDAS_ATE50.map(c => af[c])) ? ERRO : BANDAS_ATE50.reduce((s, c) => s + af[c], 0);
        const primeiro = ordem => {                     // IF(AF>0,001;cod;IF(…)) com curto-circuito
            for (const c of ordem) {
                if (ehErro(af[c])) return ERRO;
                if (af[c] > 0.001) return c;
            }
            return '';
        };
        const e = (...conds) => {                       // E(…) do Excel: erro em qualquer argumento = erro
            const vals = [];
            for (const f of conds) {
                try { vals.push(f()); } catch (x) { if (x instanceof Propaga) return ERRO; throw x; }
            }
            return vals.every(Boolean);
        };
        const semErro = x => { if (ehErro(x)) throw new Propaga(); return x; };
        const ak13 = primeiro(BANDAS_ATE50);
        const ak14 = !h15 ? '' : primeiro(BANDAS_MAIS50);
        const c15 = e(() => gt(ac23, 25), () => gt(af.SDETU5009SI, 0.001), () => semErro(ak13) !== 'SDETU5009SI');
        const ak15 = ehErro(c15) ? ERRO : (c15 ? 'SDETU5009SI' : '');
        const k14 = ak13;
        const k15 = ehErro(ak14) ? ERRO : (ak14 === '' ? ak15 : ak14);
        const k16 = ehErro(ak14) ? ERRO : (ak14 === '' ? '' : (ehErro(c15) ? ERRO : (c15 ? 'SDETU5009SI' : '')));
        const c17 = e(() => semErro(ak15) === 'SDETU5009SI', () => gt(af.SDETU5010SI, 0.001), () => semErro(ak13) !== 'SDETU5010SI');
        const k17 = ehErro(c17) ? ERRO : (c17 ? 'SDETU5010SI' : '');
        const out = [];
        for (const k of [k14, k15, k16, k17]) {
            if (ehErro(k)) out.push(['#VALUE!', '#VALUE!']);
            else if (k) out.push([k, ehErro(af[k]) ? '#VALUE!' : arred(af[k], 2)]);
        }
        return { toneladas: o5, km: r3, servicos: out };
    }

    // ── dados da parte `transporte` ─────────────────────────────────────────────────────────────
    /** Mesma normalização de services/obra_params.py (maiúsculas, sem acento, espaços simples). */
    const norm = t => String(t == null ? '' : t).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim().replace(/\s+/g, ' ');

    /** UTD da obra → {utd, km, origem}: `obra.utd`; senão a do bairro; senão a do município. utd '' = não resolvida. */
    function resolverUtd(obra, dados) {
        obra = obra || {};
        const kmUtd = {};
        Object.keys((dados && dados.km_por_utd) || {}).forEach(k => { kmUtd[norm(k)] = [k, dados.km_por_utd[k]]; });
        const linhas = (dados && dados.municipio_utd) || [];
        const comKm = (utd, origem) => { const k = kmUtd[norm(utd)]; return k ? { utd: k[0], km: k[1], origem } : null; };
        if (String(obra.utd || '').trim()) { const r = comKm(obra.utd, 'UTD da obra'); if (r) return r; }
        for (const [campo, rotulo] of [['bairro', 'bairro'], ['municipio', 'município']]) {
            const alvo = norm(obra[campo]);
            if (!alvo) continue;
            for (const l of linhas) {
                if (norm(l[campo]) === alvo && (campo !== 'bairro' || !norm(obra.municipio) || norm(l.municipio) === norm(obra.municipio))) {
                    const r = comKm(l.utd, `${rotulo} ${l[campo]}`);
                    if (r) return r;
                }
            }
        }
        return { utd: '', km: null, origem: '' };
    }

    const RE_LINHA = /^\s*\*?[0-9]+(?:[.,][0-9]+)?[-xX]\s*([^\s(]+)/;
    /** A linha de Outros é um TRANSP/TRANSP_KM (de um cálculo anterior)? */
    function ehLinhaTransporte(ativo, dados) {
        const nomes = new Set(Object.values((dados && dados.ativos) || {}).map(v => String(v).toUpperCase()));
        const t = String(ativo == null ? '' : ativo).trim().toUpperCase();
        const m = t.match(RE_LINHA);
        return nomes.has(m ? m[1] : t.split('(')[0].trim());
    }

    const qtdTexto = q => { const s = q.toFixed(2).replace(/0+$/, '').replace(/\.$/, ''); return s || '0'; };
    /** [{codigo, faixa, quantidade, texto, erro}] na ordem dos serviços: '37.6-TRANSP(faixa=25T)'; os #VALOR! vêm com erro=true. */
    function linhasParaOutros(servicos, dados) {
        const ativos = (dados && dados.ativos) || {}, porCodigo = {};
        ((dados && dados.faixas) || []).forEach(f => { porCodigo[f.ate_50km] = [ativos.ate_50km, f.faixa]; porCodigo[f.acima_50km] = [ativos.acima_50km, f.faixa]; });
        return servicos.map(([cod, q]) => {
            if (cod === '#VALUE!' || q === '#VALUE!' || !porCodigo[cod]) return { codigo: String(cod), quantidade: null, texto: '', erro: true };
            const [ativo, faixa] = porCodigo[cod];
            return { codigo: cod, faixa, quantidade: q, texto: `${qtdTexto(q)}-${ativo}(faixa=${faixa})`, erro: false };
        });
    }

    /** Qtd ODI por código de material: soma do `total` das linhas de instalação (I) do resultado do cálculo. */
    function qtdOdiDoResultado(resultado, dados) {
        const codigos = new Set(((dados && dados.pesos) || []).map(p => String(p.codigo)));
        const saida = {};
        (resultado || []).forEach(r => {
            const c = String(r.codigo || '').trim();
            if (r.operacao === 'I' && codigos.has(c)) saida[c] = (saida[c] || 0) + (Number(r.total) || 0);
        });
        return saida;
    }

    const pesosDe = dados => ((dados && dados.pesos) || []).map(p => [String(p.codigo), p.descricao || '', Number(p.kg) || 0]);

    // ── modal (só no navegador) ───────────────────────────────────────────────────────────────
    /**
     * opts: { dados, obra, obterResultado: async () => ({resultado} | {erro}), lancar: linhas[] → string (mensagem) }
     */
    function abrirModal(opts) {
        const doc = raiz.document, dados = opts.dados;
        const antigo = doc.getElementById('modal-transporte');
        if (antigo) antigo.remove();
        const el = (tag, attrs, filhos) => {
            const n = doc.createElement(tag);
            Object.entries(attrs || {}).forEach(([k, v]) => { if (v === undefined || v === null) return; if (k === 'text') n.textContent = v; else if (k === 'style') n.style.cssText = v; else n.setAttribute(k, v); });
            (filhos || []).forEach(f => { if (f) n.appendChild(typeof f === 'string' ? doc.createTextNode(f) : f); });
            return n;
        };
        const fundo = el('div', { id: 'modal-transporte', role: 'dialog', 'aria-modal': 'true', 'aria-labelledby': 'transp-titulo',
            style: 'position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:10000;display:flex;align-items:center;justify-content:center;' });
        const caixa = el('div', { style: 'background:var(--bg-canvas,#0d1117);color:var(--text-primary,#e6edf3);border:1px solid var(--border-default,#30363d);border-radius:10px;padding:18px 20px;width:min(980px,96vw);max-height:92vh;display:flex;flex-direction:column;gap:10px;font-family:Inter,sans-serif;font-size:.85rem;' });
        const fechar = () => { fundo.remove(); doc.removeEventListener('keydown', tecla, true); };
        const tecla = e => { if (e.key === 'Escape') { e.stopPropagation(); fechar(); } };
        doc.addEventListener('keydown', tecla, true);
        fundo.addEventListener('mousedown', e => { if (e.target === fundo) fechar(); });

        const utds = Object.keys(dados.km_por_utd || {}).sort();
        const auto = resolverUtd(opts.obra, dados);
        const selUtd = el('select', { id: 'transp-utd', 'aria-label': 'UTD da obra', style: 'padding:4px 6px;background:#1c1f26;color:#c9d1d9;border:1px solid #30363d;border-radius:6px;' },
            [el('option', { value: '', text: '— sem UTD (sem a parte acima de 50 km) —' })].concat(utds.map(u => el('option', { value: u, text: `${u} · ${dados.km_por_utd[u]} km` }))));
        selUtd.value = auto.utd || '';
        const infoUtd = el('span', { id: 'transp-utd-info', style: 'color:#8b949e;' });
        const aviso = el('p', { id: 'transp-aviso', role: 'status', style: 'margin:0;color:#d29922;min-height:1em;' });

        const extra = {};
        let qtdOdi = {}, mostrarTodas = false, resultadoAtual = null;
        const corpoTabela = el('tbody');
        const tabela = el('table', { id: 'transp-materiais', style: 'width:100%;border-collapse:collapse;font-size:.8rem;' }, [
            el('thead', {}, [el('tr', {}, ['Código', 'Descrição', 'kg', 'Qtd ODI', 'Qtd extra', 'Peso (kg)'].map((t, i) =>
                el('th', { text: t, style: 'position:sticky;top:0;background:var(--bg-canvas,#0d1117);text-align:' + (i < 2 ? 'left' : 'right') + ';padding:4px 6px;border-bottom:1px solid #30363d;' })))]),
            corpoTabela
        ]);
        const btnTodas = el('button', { type: 'button', class: 'btn-secondary', id: 'transp-todas', text: 'Mostrar todas' });
        const total = el('strong', { id: 'transp-total' });
        const listaServicos = el('div', { id: 'transp-servicos', style: 'display:flex;flex-direction:column;gap:3px;' });
        const btnLancar = el('button', { type: 'button', class: 'btn-primary', id: 'transp-lancar', text: 'Lançar em Outros' });
        const btnFechar = el('button', { type: 'button', class: 'btn-secondary', text: 'Fechar' });
        btnFechar.onclick = fechar;

        const fmt = (v, c) => Number(v).toLocaleString('pt-BR', { minimumFractionDigits: c, maximumFractionDigits: c });
        function kmAtual() { const u = selUtd.value; return u ? dados.km_por_utd[u] : null; }
        function recalcular() {
            const km = kmAtual();
            infoUtd.textContent = selUtd.value ? `UTD ${selUtd.value} → ${km} km (ida e volta)` + (selUtd.value === auto.utd && auto.origem ? ` · pelo ${auto.origem}` : ' · escolhida aqui') : 'Informe o município/UTD da obra: sem UTD o cálculo segue sem a parte acima de 50 km.';
            corpoTabela.textContent = '';
            (dados.pesos || []).forEach((p, i) => {
                const cod = String(p.codigo), q = qtdOdi[cod] || 0, x = Number(extra[cod]) || 0;
                if (!mostrarTodas && !q && !x) return;
                const inp = el('input', { type: 'number', step: 'any', min: '0', value: extra[cod] || '', 'aria-label': 'Quantidade extra de ' + cod, 'data-codigo': cod, style: 'width:70px;text-align:right;background:#1c1f26;color:#c9d1d9;border:1px solid #30363d;border-radius:4px;padding:2px 4px;' });
                inp.oninput = () => { extra[cod] = inp.value === '' ? '' : Number(inp.value); atualizarTotal(); };
                const td = (t, dir) => el('td', { text: t, style: 'padding:3px 6px;border-bottom:1px solid #21262d;text-align:' + (dir || 'right') + ';' });
                corpoTabela.appendChild(el('tr', { 'data-linha': String(i) }, [td(cod, 'left'), td(p.descricao || '', 'left'), td(fmt(p.kg, 0)), td(fmt(q, 2)),
                    el('td', { style: 'padding:3px 6px;border-bottom:1px solid #21262d;text-align:right;' }, [inp]), el('td', { class: 'transp-peso', text: fmt((q + x) * p.kg, 1), style: 'padding:3px 6px;border-bottom:1px solid #21262d;text-align:right;' })]));
            });
            if (!corpoTabela.children.length) corpoTabela.appendChild(el('tr', {}, [el('td', { colspan: '6', text: 'Nenhum material pesado com quantidade nesta obra. Use "Mostrar todas" para digitar quantidades extras.', style: 'padding:10px;color:#8b949e;text-align:center;' })]));
            atualizarTotal();
        }
        function atualizarTotal() {
            const km = kmAtual(), num = {};
            Object.keys(extra).forEach(k => { if (extra[k] !== '') num[k] = Number(extra[k]) || 0; });
            corpoTabela.querySelectorAll('tr[data-linha]').forEach(tr => {
                const p = dados.pesos[Number(tr.dataset.linha)], cod = String(p.codigo);
                tr.querySelector('.transp-peso').textContent = fmt(((qtdOdi[cod] || 0) + (num[cod] || 0)) * p.kg, 1);
            });
            const r = calcular(qtdOdi, num, km, pesosDe(dados));
            resultadoAtual = linhasParaOutros(r.servicos, dados);
            total.textContent = fmt(r.toneladas, 4) + ' t';
            listaServicos.textContent = '';
            if (!resultadoAtual.length) listaServicos.appendChild(el('span', { text: 'Sem serviço de transporte (obra sem material pesado).', style: 'color:#8b949e;' }));
            resultadoAtual.forEach(s => listaServicos.appendChild(s.erro
                ? el('div', { class: 'transp-servico erro', style: 'color:#f85149;', text: 'Na ICP esta linha dá #VALOR! — não será lançada.' })
                : el('div', { class: 'transp-servico' }, [el('code', { text: s.codigo }), ` · faixa ${s.faixa} · ${fmt(s.quantidade, 2)} → `, el('code', { text: s.texto })])));
        }
        selUtd.onchange = recalcular;
        btnTodas.onclick = () => { mostrarTodas = !mostrarTodas; btnTodas.textContent = mostrarTodas ? 'Só com quantidade' : 'Mostrar todas'; recalcular(); };
        btnLancar.onclick = () => {
            const linhas = (resultadoAtual || []).filter(s => !s.erro).map(s => s.texto);
            aviso.textContent = opts.lancar(linhas) || '';
        };

        caixa.append(
            el('h3', { id: 'transp-titulo', text: 'Transporte de material', style: 'margin:0;' }),
            el('p', { style: 'margin:0;color:#8b949e;line-height:1.4;', text: 'Frete dos materiais pesados (postes, cruzetas, blocos e placas de concreto), igual à aba TRANSPORT da ICP. A Qtd ODI vem do orçamento do Resumo (instalação); a Qtd extra é digitada aqui.' }),
            el('div', { style: 'display:flex;gap:8px;align-items:center;flex-wrap:wrap;' }, [el('label', { for: 'transp-utd', text: 'UTD:' }), selUtd, infoUtd]),
            aviso,
            el('div', { style: 'overflow:auto;flex:1;min-height:120px;border:1px solid #30363d;border-radius:6px;' }, [tabela]),
            el('div', { style: 'display:flex;gap:10px;align-items:center;' }, [btnTodas, el('span', { text: 'Total:' }), total]),
            el('div', {}, [el('h4', { text: 'Serviços', style: 'margin:0 0 4px;' }), listaServicos]),
            el('div', { style: 'display:flex;gap:8px;justify-content:flex-end;' }, [btnFechar, btnLancar])
        );
        fundo.appendChild(caixa);
        doc.body.appendChild(fundo);
        aviso.textContent = 'Calculando o orçamento do Resumo…';
        recalcular();
        selUtd.focus();
        Promise.resolve(opts.obterResultado()).then(r => {
            if (!doc.body.contains(fundo)) return;
            if (r && r.erro) { aviso.textContent = r.erro + ' A Qtd ODI ficou 0.'; return; }
            qtdOdi = qtdOdiDoResultado(r && r.resultado, dados);
            aviso.textContent = '';
            recalcular();
        }).catch(e => { aviso.textContent = 'Não foi possível calcular o orçamento do Resumo (' + (e && e.message || e) + '). A Qtd ODI ficou 0.'; });
        return { fechar };
    }

    const api = { calcular, arred, resolverUtd, ehLinhaTransporte, linhasParaOutros, qtdOdiDoResultado, pesosDe, abrirModal, BANDAS_ATE50, BANDAS_MAIS50 };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.Transporte = api;
})(typeof window !== 'undefined' ? window : globalThis);
