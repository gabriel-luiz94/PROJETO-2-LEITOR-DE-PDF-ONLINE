/*
 * resumo_rede.js (TASK-192) — as 4 contagens do «Resumo da rede» (postes, rede de média, rede de baixa e equipamentos instalando), sem DOM.
 * Estava dentro de resumo.js com regras que só serviam à Paraíba; aqui valem para os dois estados, testado em Node (tests/test_resumo_rede_192.py):
 *   · instalando = operação `I` OU `*I` (linha viva também instala); R, *R, M, *M e vazia não contam;
 *   · poste = nome que começa por DT/RC/CV/FV (Paraíba) OU que a CONVENÇÃO DE POSTES do projeto reconhece, com ou sem parâmetros:
 *     em Pernambuco a base cadastra o poste com o nome compacto (`12600D=12600D(fases=TRI), DT12/600`), então o item chega como `12600D`;
 *   · equipamento = TR…, CFU…, CFA…;  média/baixa: cabo com tensão (kV) no rótulo = média (o neutro não entra), sem tensão = baixa; no formato clássico CAA/CAL/P = média, M2X/M3X = baixa.
 */
(function (raiz) {
    'use strict';

    const instalando = op => { const o = String(op == null ? '' : op).trim().toUpperCase(); return o === 'I' || o === '*I'; };
    const semParametros = a => String(a == null ? '' : a).replace(/\(.*$/, '').trim();
    const RE_POSTE_FIXO = /^(DT|RC|CV|FV)/i, RE_EQUIPAMENTO = /^(TR|CFU|CFA)/i;

    /** O item (já com o nome oficial da base) é poste: pela lista fixa da Paraíba ou pela convenção de postes do projeto (`d.convencaoPostes`; sem ela, a de Pernambuco). */
    function ehPoste(ativo, d) {
        const nome = semParametros(ativo);
        if (!nome) return false;
        if (RE_POSTE_FIXO.test(nome)) return true;
        return !!(d.AnotacaoOutros && d.AnotacaoOutros.ehPoste && d.AnotacaoOutros.ehPoste(nome, d.convencaoPostes));
    }

    /**
     * d = { cabos, outros (linhas da tela), ctxOutros, convencaoPostes, AnotacaoOutros, AnotacaoCabos, extrairPares, extrairPrefixoEComprimento, extrairPrefixoSemComprimento }.
     * Devolve { postes, redeMedia, redeBaixa, equipamentos }.
     */
    function calcular(d) {
        const resumo = { postes: 0, redeMedia: 0, redeBaixa: 0, equipamentos: 0 };

        (d.outros || []).forEach(row => {
            if (!row || row.congelada || !instalando(row.operacao)) return;
            d.AnotacaoOutros.interpretarLinha(row.ativo, 'I', d.extrairPares, d.ctxOutros).itens.forEach(({ ativo, qtd }) => {
                if (ehPoste(ativo, d)) resumo.postes += qtd;
                else if (RE_EQUIPAMENTO.test(semParametros(ativo))) resumo.equipamentos += qtd;
            });
        });

        const linhasCabos = d.cabos || [];
        linhasCabos.forEach((row, idx) => {
            if (!row || row.congelada || !instalando(row.operacao)) return;
            // TASK-116/118: cabo reconhecido, mas sem comprimento na própria linha → vale o comprimento da linha de baixo: só o número ("45 m")
            // ou outro cabo com comprimento ("CAZ9,5 abc 50 m"). A linha de baixo continua sendo contada por si mesma.
            const abaixo = () => {
                const p = linhasCabos[idx + 1];
                if (!p || p.congelada) return null;
                const solto = d.AnotacaoCabos.comprimentoSozinho(p.ativo);
                if (solto !== null) return solto;
                const info = d.extrairPrefixoEComprimento(p.ativo);
                return info ? info.comprimento : null;
            };
            if (d.AnotacaoCabos.ehNotacao(row.ativo)) {                                           // TASK-083: com tensão = média; sem = baixa
                const nr = d.AnotacaoCabos.interpretar(row.ativo);
                nr.itens.forEach(it => {
                    const comp = it.comprimento !== null ? it.comprimento : abaixo();
                    if (comp === null) return;
                    if (nr.tensao !== null) { if (!it.neutro) resumo.redeMedia += comp; }          // TASK-116: o neutro não entra na Média
                    else resumo.redeBaixa += comp;
                });
                return;
            }
            let info = d.extrairPrefixoEComprimento(row.ativo);
            if (!info) {
                const pref = d.extrairPrefixoSemComprimento(row.ativo);
                const comp = pref ? abaixo() : null;
                if (comp === null) return;
                info = { prefixo: pref, comprimento: comp };
            }
            if (/^(CAA|CAL|P)/i.test(info.prefixo)) resumo.redeMedia += info.comprimento;
            else if (/^(M2X|M3X)/i.test(info.prefixo)) resumo.redeBaixa += info.comprimento;
        });

        return resumo;
    }

    const api = { calcular, ehPoste, instalando, semParametros };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.ResumoRede = api;
})(typeof window !== 'undefined' ? window : globalThis);
