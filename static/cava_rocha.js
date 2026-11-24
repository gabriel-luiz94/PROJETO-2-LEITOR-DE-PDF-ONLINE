/*
 * cava_rocha.js (TASK-248) — cava em rocha (SDEMU3007II) igual à aba CAVA_ROCHA da ICP.
 * MESMA função de services/cava_rocha.py (paridade testada em tests/test_cava_rocha_248.py). A parte de cálculo não usa DOM;
 * o modal (abrirModal) só existe no navegador. Não "consertar" a conta: a regra é dar o mesmo resultado que a ICP.
 *   F     = estais em solo + postes de 9 m para cima (os códigos de `contagem`, na instalação da ODI)
 *   cavas = F se ARREDMULTB(MARRED(F × índice; 0,1); 1) > F, senão ARREDMULTB(MARRED(F × índice; 0,1); 1)
 */
(function (raiz) {
    'use strict';

    /** Número com 15 algarismos significativos (como o Excel e o Decimal(format(x, '.15g')) do Python) → {neg, dig: BigInt, exp}. */
    function decimal15(x) {
        const neg = x < 0, s = Math.abs(x).toPrecision(15);
        const m = s.match(/^(\d+)(?:\.(\d+))?(?:e([+-]\d+))?$/);
        const frac = m[2] || '';
        return { neg, dig: BigInt(m[1] + frac), exp: (m[3] ? parseInt(m[3], 10) : 0) - frac.length };
    }
    /** round(x × 10^k) com o modo dado ('meio' = meio para longe do zero; 'piso' = para baixo), sobre o número com 15 algarismos. */
    function inteiroEscalado(x, k, modo) {
        const { neg, dig, exp } = decimal15(x);
        const e = exp + k;
        let q;
        if (e >= 0) q = dig * (10n ** BigInt(e));
        else {
            const div = 10n ** BigInt(-e), r = dig % div;
            q = dig / div;
            if (modo === 'meio' && r * 2n >= div) q += 1n;
            if (modo === 'piso' && neg && r > 0n) q += 1n;
        }
        return neg ? -q : q;
    }
    const comoNumero = (q, k) => Number(k > 0 ? `${q}e-${k}` : String(q));

    /** MARRED(x; m): múltiplo de m mais próximo, meio para longe do zero (m = 10^-k: 1, 0,1, 0,01…). */
    function mround(x, m) {
        if (x === 0) return 0;
        const k = Math.round(-Math.log10(m));
        if (Math.abs(Math.pow(10, -k) - m) > 1e-15) return Math.round(x / m) * m;       // múltiplo que não é potência de 10 (não usado pela ICP)
        return comoNumero(inteiroEscalado(x, k, 'meio'), k);
    }
    /** ARREDMULTB(x; 1). */
    const floor1 = x => comoNumero(inteiroEscalado(x, 0, 'piso'), 0);

    /** Quantidade de cavas em rocha de UMA localidade (linha da aba). */
    function cavas(postesEstais, indice) {
        const f = postesEstais, g = floor1(mround(f * indice, 0.1));
        return g > f ? f : g;
    }

    /** {f, cavas}: localidade = DASHBORD!C4; qtdServicos = {codigo: quantidade}; tabela = [[localidade, indice], …] na ordem da aba. */
    function calcular(localidade, qtdServicos, tabela, contagem) {
        let f = 0;
        (contagem || ['SDEMU3002II', 'SDEMU1004II', 'SDEMU1005II']).forEach(c => { f += Number((qtdServicos || {})[c]) || 0; });
        let total = 0;
        (tabela || []).forEach(([loc, indice]) => { if (loc === localidade) total += cavas(f, indice); });
        return { f, cavas: total };
    }

    // ── dados da parte `cava_rocha` ─────────────────────────────────────────────────────────────
    /** Mesma normalização de services/obra_params.py (maiúsculas, sem acento, espaços simples; «_» = espaço). */
    const norm = t => String(t == null ? '' : t).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim().replace(/\s+/g, ' ').replace(/_/g, ' ');
    const tabelaDe = dados => ((dados && dados.localidades) || []).map(l => [l.localidade, Number(l.indice)]);

    /** Localidade da obra → {localidade, indice, relacao, origem}: pelo bairro; senão pelo município. localidade '' = fora da tabela. */
    function resolverLocalidade(obra, dados) {
        obra = obra || {};
        const locs = (dados && dados.localidades) || [];
        for (const [campo, rotulo] of [['bairro', 'bairro'], ['municipio', 'município']]) {
            const alvo = norm(obra[campo]);
            if (!alvo) continue;
            const l = locs.find(x => norm(x.localidade) === alvo);
            if (l) return { localidade: l.localidade, indice: l.indice, relacao: l.relacao || '', origem: `${rotulo} ${obra[campo]}` };
        }
        return { localidade: '', indice: null, relacao: '', origem: '' };
    }
    /** A localidade digitada/escolhida (sem acento, maiúsculas) → a linha da tabela, ou null. */
    const localidadePorNome = (nome, dados) => ((dados && dados.localidades) || []).find(l => norm(l.localidade) === norm(nome)) || null;

    const RE_LINHA = /^\s*\*?[0-9]+(?:[.,][0-9]+)?[-xX]\s*([^\s(,]+)/;
    /** A linha de Outros é a cava em rocha (de um cálculo anterior)? */
    function ehLinhaCava(ativo, dados) {
        const nome = String((dados && dados.ativo) || '').trim().toUpperCase();
        if (!nome) return false;
        const t = String(ativo == null ? '' : ativo).trim().toUpperCase();
        const m = t.match(RE_LINHA);
        return (m ? m[1] : t.split('(')[0].trim()) === nome;
    }
    const qtdTexto = q => { const s = q.toFixed(2).replace(/0+$/, '').replace(/\.$/, ''); return s || '0'; };
    /** '6-CAVA_ROCHA'; '' quando a quantidade é 0 (só remove a antiga). */
    const linhaParaOutros = (quantidade, dados) => (quantidade ? `${qtdTexto(quantidade)}-${dados.ativo}` : '');

    /** Qtd ODI por código contado: soma do `total` das linhas de instalação (I) do resultado do cálculo; e a descrição de cada código. */
    function qtdDoResultado(resultado, dados) {
        const codigos = new Set(((dados && dados.contagem) || []).map(String)), qtd = {}, descricao = {};
        (resultado || []).forEach(r => {
            const c = String(r.codigo || '').trim();
            if (r.operacao === 'I' && codigos.has(c)) { qtd[c] = (qtd[c] || 0) + (Number(r.total) || 0); if (r.desc_codigo) descricao[c] = r.desc_codigo; }
        });
        return { qtd, descricao };
    }

    // ── modal (só no navegador) ───────────────────────────────────────────────────────────────
    const DESCRICAO_PADRAO = { SDEMU3002II: 'INSTALAR ESTAI EM SOLO', SDEMU1004II: 'INSTALAR POSTE DE 9 A 14 M', SDEMU1005II: 'INSTALAR POSTE DE 14 M OU MAIS' };
    /**
     * opts: { dados, obra, obterResultado: async () => ({resultado} | {erro}), lancar: (texto, dados) → string (mensagem) }
     */
    function abrirModal(opts) {
        const doc = raiz.document, dados = opts.dados;
        const antigo = doc.getElementById('modal-cava-rocha');
        if (antigo) antigo.remove();
        const el = (tag, attrs, filhos) => {
            const n = doc.createElement(tag);
            Object.entries(attrs || {}).forEach(([k, v]) => { if (v === undefined || v === null) return; if (k === 'text') n.textContent = v; else if (k === 'style') n.style.cssText = v; else n.setAttribute(k, v); });
            (filhos || []).forEach(f => { if (f) n.appendChild(typeof f === 'string' ? doc.createTextNode(f) : f); });
            return n;
        };
        const fundo = el('div', { id: 'modal-cava-rocha', role: 'dialog', 'aria-modal': 'true', 'aria-labelledby': 'cava-titulo',
            style: 'position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:10000;display:flex;align-items:center;justify-content:center;' });
        const caixa = el('div', { style: 'background:var(--bg-canvas,#0d1117);color:var(--text-primary,#e6edf3);border:1px solid var(--border-default,#30363d);border-radius:10px;padding:18px 20px;width:min(760px,96vw);max-height:92vh;overflow:auto;display:flex;flex-direction:column;gap:10px;font-family:Inter,sans-serif;font-size:.85rem;' });
        const fechar = () => { fundo.remove(); doc.removeEventListener('keydown', tecla, true); };
        const tecla = e => { if (e.key === 'Escape' && doc.body.contains(fundo)) { e.stopPropagation(); fechar(); } };
        doc.addEventListener('keydown', tecla, true);
        fundo.addEventListener('mousedown', e => { if (e.target === fundo) fechar(); });

        const fmt = (v, c) => Number(v).toLocaleString('pt-BR', { minimumFractionDigits: c, maximumFractionDigits: c });
        const fmtInd = v => Number(v).toLocaleString('pt-BR', { maximumFractionDigits: 4 });
        const auto = resolverLocalidade(opts.obra, dados);
        const lista = el('datalist', { id: 'cava-localidades' }, (dados.localidades || []).map(l => el('option', { value: l.localidade, text: `índice ${fmtInd(l.indice)}${l.relacao ? ' · ' + l.relacao : ''}${l.utd ? ' · UTD ' + l.utd : ''}` })));
        const inpLoc = el('input', { id: 'cava-localidade', list: 'cava-localidades', 'aria-label': 'Localidade da obra', autocomplete: 'off', placeholder: 'digite para buscar…', value: auto.localidade || '',
            style: 'padding:4px 6px;min-width:240px;background:#1c1f26;color:#c9d1d9;border:1px solid #30363d;border-radius:6px;' });
        const infoLoc = el('span', { id: 'cava-localidade-info', style: 'color:#8b949e;' });
        const aviso = el('p', { id: 'cava-aviso', role: 'status', style: 'margin:0;color:#d29922;min-height:1em;' });
        const avisoBase = dados.ativo_na_base === false
            ? el('p', { id: 'cava-aviso-base', style: 'margin:0;color:#f85149;', text: `A base deste projeto não tem o ativo ${dados.ativo}: o Montar Orçamento vai avisar "ativo não encontrado".` }) : null;

        const extra = {};
        let qtd = {}, descricao = {}, atual = { f: 0, cavas: 0 }, linha = '';
        const corpo = el('tbody');
        const tabela = el('table', { id: 'cava-contagem', style: 'width:100%;border-collapse:collapse;font-size:.8rem;' }, [
            el('thead', {}, [el('tr', {}, ['Código', 'Descrição', 'Qtd ODI', 'Qtd extra', 'Total'].map((t, i) =>
                el('th', { text: t, style: 'text-align:' + (i < 2 ? 'left' : 'right') + ';padding:4px 6px;border-bottom:1px solid #30363d;' })))]),
            corpo
        ]);
        const conta = el('div', { id: 'cava-conta', style: 'line-height:1.6;' });
        const btnLancar = el('button', { type: 'button', class: 'btn-primary', id: 'cava-lancar', text: 'Lançar em Outros' });
        const btnFechar = el('button', { type: 'button', class: 'btn-secondary', text: 'Fechar' });
        btnFechar.onclick = fechar;
        const td = (t, dir, cls) => el('td', { text: t, class: cls, style: 'padding:3px 6px;border-bottom:1px solid #21262d;text-align:' + (dir || 'right') + ';' });

        function desenharTabela() {
            corpo.textContent = '';
            (dados.contagem || []).forEach(cod => {
                const inp = el('input', { type: 'number', step: 'any', min: '0', value: extra[cod] === undefined ? '' : extra[cod], 'aria-label': 'Quantidade extra de ' + cod, 'data-codigo': cod,
                    style: 'width:70px;text-align:right;background:#1c1f26;color:#c9d1d9;border:1px solid #30363d;border-radius:4px;padding:2px 4px;' });
                inp.oninput = () => { extra[cod] = inp.value === '' ? undefined : Number(inp.value); recalcular(); };
                corpo.appendChild(el('tr', { 'data-codigo': cod }, [td(cod, 'left'), td(descricao[cod] || DESCRICAO_PADRAO[cod] || '', 'left'), td(fmt(qtd[cod] || 0, 2), 'right', 'cava-odi'),
                    el('td', { style: 'padding:3px 6px;border-bottom:1px solid #21262d;text-align:right;' }, [inp]), td('', 'right', 'cava-total')]));
            });
        }
        function recalcular() {
            const loc = localidadePorNome(inpLoc.value, dados);
            const soma = {};
            (dados.contagem || []).forEach(c => { soma[c] = (qtd[c] || 0) + (Number(extra[c]) || 0); });
            corpo.querySelectorAll('tr[data-codigo]').forEach(tr => { tr.querySelector('.cava-total').textContent = fmt(soma[tr.dataset.codigo] || 0, 2); });
            if (loc) {
                infoLoc.textContent = `Localidade ${loc.localidade} → índice ${fmtInd(loc.indice)}` + (loc.relacao ? ` (${loc.relacao})` : '') + (loc.localidade === auto.localidade && auto.origem ? ` · pelo ${auto.origem}` : ' · escolhida aqui');
                infoLoc.style.color = '#8b949e';
            } else {
                infoLoc.textContent = inpLoc.value.trim() ? 'Esta localidade não tem cava em rocha na ICP: quantidade 0.' : 'Informe o bairro/localidade da obra (sem localidade da tabela, a quantidade é 0).';
                infoLoc.style.color = '#d29922';
            }
            atual = calcular(loc ? loc.localidade : '', soma, tabelaDe(dados), dados.contagem);
            linha = linhaParaOutros(atual.cavas, dados);
            conta.textContent = '';
            const bruto = loc ? atual.f * Number(loc.indice) : 0, arred01 = loc ? mround(bruto, 0.1) : 0, piso = loc ? floor1(arred01) : 0;
            conta.append(
                el('div', {}, ['F (estais em solo + postes de 9 m ou mais) = ', el('strong', { id: 'cava-f', text: fmt(atual.f, 2) })]),
                loc ? el('div', {}, [`F × índice = ${fmt(atual.f, 2)} × ${fmtInd(loc.indice)} = ${fmt(bruto, 4)} → arredondado a 0,1 = ${fmt(arred01, 1)} → parte inteira = ${fmt(piso, 0)}` + (piso > atual.f ? ` → limitado a F = ${fmt(atual.f, 2)}` : '')]) : null,
                el('div', {}, ['Cavas em rocha (', el('code', { text: dados.servico || '' }), '): ', el('strong', { id: 'cava-qtd', text: fmt(atual.cavas, 0) }),
                    ' → ', linha ? el('code', { id: 'cava-linha', text: linha }) : el('span', { id: 'cava-linha', text: 'nada a lançar (remove a linha anterior, se houver)', style: 'color:#8b949e;' })])
            );
        }
        inpLoc.oninput = recalcular;
        btnLancar.onclick = () => {
            if (dados.ativo_na_base === false && linha && !raiz.confirm(`A base deste projeto não tem o ativo ${dados.ativo}. Lançar mesmo assim?`)) return;
            aviso.textContent = opts.lancar(linha, dados) || '';
        };

        caixa.append(
            el('h3', { id: 'cava-titulo', text: 'Cava em rocha', style: 'margin:0;' }),
            el('p', { style: 'margin:0;color:#8b949e;line-height:1.4;', text: `Estimativa de cavas em rocha (${dados.servico || ''}) pela localidade, igual à aba CAVA_ROCHA da ICP: índice da localidade × (estais em solo + postes de 9 m ou mais instalados). A Qtd ODI vem do orçamento do Resumo; a Qtd extra é digitada aqui.` }),
            el('div', { style: 'display:flex;gap:8px;align-items:center;flex-wrap:wrap;' }, [el('label', { for: 'cava-localidade', text: 'Localidade:' }), inpLoc, lista, infoLoc]),
            avisoBase, aviso,
            el('div', { style: 'border:1px solid #30363d;border-radius:6px;overflow:auto;' }, [tabela]),
            el('div', {}, [el('h4', { text: 'Resultado', style: 'margin:0 0 4px;' }), conta]),
            el('div', { style: 'display:flex;gap:8px;justify-content:flex-end;' }, [btnFechar, btnLancar])
        );
        fundo.appendChild(caixa);
        doc.body.appendChild(fundo);
        aviso.textContent = 'Calculando o orçamento do Resumo…';
        desenharTabela(); recalcular();
        inpLoc.focus();
        Promise.resolve(opts.obterResultado()).then(r => {
            if (!doc.body.contains(fundo)) return;
            if (r && r.erro) { aviso.textContent = r.erro + ' A Qtd ODI ficou 0.'; return; }
            const q = qtdDoResultado(r && r.resultado, dados);
            qtd = q.qtd; descricao = q.descricao;
            aviso.textContent = Object.keys(qtd).length ? '' : 'O orçamento do Resumo não tem estais em solo nem postes de 9 m ou mais instalados: a Qtd ODI ficou 0.';
            desenharTabela(); recalcular();
        }).catch(e => { aviso.textContent = 'Não foi possível calcular o orçamento do Resumo (' + (e && e.message || e) + '). A Qtd ODI ficou 0.'; });
        return { fechar };
    }

    const api = { calcular, cavas, mround, floor1, resolverLocalidade, localidadePorNome, ehLinhaCava, linhaParaOutros, qtdDoResultado, tabelaDe, abrirModal };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.CavaRocha = api;
})(typeof window !== 'undefined' ? window : globalThis);
