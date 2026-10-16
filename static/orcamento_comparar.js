/*
 * orcamento_comparar.js (TASK-071) — lógica pura do modo comparação do Resultado do Orçamento: compara os itens do REC atual
 * com os de um REC de referência. Identidade do item = operação | MDO | código (a mesma chave do consolidar). Status:
 *   igual · diferente (quantidades distintas) · so_atual (falta na referência) · so_referencia (falta no atual).
 * Sem DOM; testado em Node (tests/test_orcamento_comparar.py).
 */
(function (raiz) {
    'use strict';

    const normOp = op => (String(op == null ? '' : op).trim().toUpperCase() === 'R' ? 'R' : 'I');
    function normMdo(mdo) {
        const m = String(mdo == null ? '' : mdo).trim().toUpperCase();
        if (m === 'MATERIAL' || m === 'M') return 'MATERIAL';
        return 'MAO-DE-OBRA';                    // vazio, '-', N, S, MAO/MÃO-DE-OBRA: como o consolidar do REC
    }
    const chave = (op, mdo, codigo) => normOp(op) + '|' + normMdo(mdo) + '|' + String(codigo == null ? '' : codigo).trim().toUpperCase();
    const arred = n => Math.round((Number(n) || 0) * 100);
    const igualQtd = (a, b) => arred(a) === arred(b);
    const qtdPtBr = txt => {
        if (typeof txt === 'number') return txt;
        const n = parseFloat(String(txt == null ? '' : txt).trim().replace(/\./g, '').replace(',', '.'));
        return isNaN(n) ? 0 : n;
    };

    /** Texto de um .rec (`codigo;qtd;M|S;IN|DR;...`) -> itens consolidados por chave (repetidos são somados). */
    function parseRec(texto) {
        const mapa = new Map();
        String(texto == null ? '' : texto).split(/\r?\n/).forEach(linha => {
            if (!linha.trim()) return;
            const p = linha.split(';');
            if (p.length < 6) return;
            const codigo = p[0].trim(), qtd = parseFloat(p[1].replace(',', '.')) || 0;
            const operacao = p[3].trim().toUpperCase() === 'IN' ? 'I' : 'R';
            const mdo = p[2].trim().toUpperCase() === 'S' ? 'MAO-DE-OBRA' : 'MATERIAL';
            const k = chave(operacao, mdo, codigo);
            if (mapa.has(k)) mapa.get(k).qtd = Math.round((mapa.get(k).qtd + qtd) * 100) / 100;
            else mapa.set(k, { chave: k, operacao, mdo, codigo, qtd: Math.round(qtd * 100) / 100, desc: '' });
        });
        return Array.from(mapa.values());
    }

    /**
     * atual: [{operacao, mdo, codigo, qtd (número)}] na ordem da tela; referencia: [{operacao, mdo, codigo, qtd, desc}] (parseRec).
     * -> { linhas:[{status, ref}], fantasmas:[item da referência], contagem:{igual,diferente,so_atual,so_referencia} }.
     * Linhas repetidas da mesma chave: compara a SOMA; a referência aparece só na 1ª (ref = null nas demais).
     */
    function comparar(atual, referencia) {
        const ref = new Map((referencia || []).map(i => [chave(i.operacao, i.mdo, i.codigo), i]));
        const soma = new Map(), primeira = new Map();
        atual.forEach((l, i) => {
            const k = chave(l.operacao, l.mdo, l.codigo);
            soma.set(k, (soma.get(k) || 0) + (Number(l.qtd) || 0));
            if (!primeira.has(k)) primeira.set(k, i);
        });
        const contagem = { igual: 0, diferente: 0, so_atual: 0, so_referencia: 0 };
        const status = new Map();
        soma.forEach((q, k) => {
            const r = ref.get(k);
            const s = !r ? 'so_atual' : (igualQtd(q, r.qtd) ? 'igual' : 'diferente');
            status.set(k, s);
            contagem[s]++;
        });
        const linhas = atual.map((l, i) => {
            const k = chave(l.operacao, l.mdo, l.codigo), r = ref.get(k);
            return { status: status.get(k), ref: primeira.get(k) === i ? (r ? r.qtd : 0) : null };
        });
        const fantasmas = [];
        ref.forEach((r, k) => { if (!soma.has(k)) { fantasmas.push(Object.assign({}, r, { chave: k })); contagem.so_referencia++; } });
        return { linhas, fantasmas, contagem };
    }

    /** Ordem padrão da tabela REC: MDO (mão de obra antes de material) → operação → descrição. */
    function comparadorPadrao(a, b) {
        const rank = m => { const x = normMdo(m); return x === 'MAO-DE-OBRA' ? 1 : 2; };
        return (rank(a.mdo) - rank(b.mdo))
            || normOp(a.operacao).localeCompare(normOp(b.operacao), 'pt-BR')
            || String(a.desc || '').toUpperCase().localeCompare(String(b.desc || '').toUpperCase(), 'pt-BR');
    }
    /** Índice onde `item` entra em `linhas` (já na ordem da tela) segundo `cmp`: antes da primeira linha que vem depois dele. */
    function posicaoInsercao(linhas, item, cmp) {
        for (let i = 0; i < linhas.length; i++) if (cmp(linhas[i], item) > 0) return i;
        return linhas.length;
    }

    const STATUS_CLASSE = { igual: 'cmp-igual', diferente: 'cmp-diferente', so_atual: 'cmp-so-atual', so_referencia: 'cmp-so-ref' };
    const STATUS_ROTULO = { igual: 'iguais', diferente: 'diferentes', so_atual: 'só no atual', so_referencia: 'só na referência' };

    const api = { normOp, normMdo, chave, igualQtd, qtdPtBr, parseRec, comparar, comparadorPadrao, posicaoInsercao, STATUS_CLASSE, STATUS_ROTULO };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    raiz.OrcamentoComparar = api;
})(typeof window !== 'undefined' ? window : globalThis);
