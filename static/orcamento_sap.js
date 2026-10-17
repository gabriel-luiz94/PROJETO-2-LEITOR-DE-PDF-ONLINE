/*
 * orcamento_sap.js (TASK-070) — lógica pura da visão SAP do Resultado do Orçamento: duas tabelas, ODI (instalando, operação I)
 * e ODD (removendo, operação R), com TIPO MDO abreviado (M = MATERIAL, N = MAO-DE-OBRA). É só visualização: o dado e a
 * gravação continuam sendo os do REC (tabela da tela). Testado em Node: tests/test_orcamento_sap.py.
 */
(function (raiz) {
    'use strict';

    /** MATERIAL -> M; MAO-DE-OBRA (ou vazio/'-'/S) -> N; o que já vier M/N é mantido; outro texto fica como está (maiúsculo). */
    function rotuloMdo(mdo) {
        const m = String(mdo == null ? '' : mdo).trim().toUpperCase();
        if (m === 'MATERIAL' || m === 'M') return 'M';
        if (m === '' || m === '-' || m === 'MAO-DE-OBRA' || m === 'MÃO-DE-OBRA' || m === 'N' || m === 'S') return 'N';
        return m;
    }
    /** O que o usuário digitou -> valor interno do REC, ou null se não for M/N (nem os nomes por extenso). */
    function mdoDeRotulo(txt) {
        const m = String(txt == null ? '' : txt).trim().toUpperCase();
        if (m === 'M' || m === 'MATERIAL') return 'MATERIAL';
        if (m === 'N' || m === 'S' || m === 'MAO-DE-OBRA' || m === 'MÃO-DE-OBRA') return 'MAO-DE-OBRA';
        return null;
    }

    /** CENTRO da visão SAP: 1600 para material, 1000 para mão de obra (derivado do TIPO MDO; outro tipo = vazio). */
    function centroDe(mdo) {
        const r = rotuloMdo(mdo);
        return r === 'M' ? '1600' : (r === 'N' ? '1000' : '');
    }

    /** linhas do REC ({operacao, ...}) -> { odi, odd }. Só 'R' é remoção; tudo o mais é instalação (como o REC mostra). */
    function separar(linhas) {
        const odi = [], odd = [];
        (linhas || []).forEach(l => (String(l.operacao || 'I').trim().toUpperCase() === 'R' ? odd : odi).push(l));
        return { odi, odd };
    }

    function numeroPtBr(txt) {
        if (typeof txt === 'number') return txt;
        const n = parseFloat(String(txt == null ? '' : txt).trim().replace(/\./g, '').replace(',', '.'));
        return isNaN(n) ? 0 : n;
    }

    function textoDaColuna(l, coluna) {
        switch (coluna) {
            case 'mdo': return rotuloMdo(l.mdo);
            case 'codigo': return String(l.codigo || '');
            case 'centro': return centroDe(l.mdo);
            case 'desc': return String(l.desc || '');
            case 'qtd': return String(l.qtd == null ? '' : l.qtd);
            case 'filtro': return String(l.filtro || '');
            default: return '';
        }
    }

    /** Comparador de duas linhas por coluna (texto A-Z, quantidade numérica). Positivo = a depois de b (no sentido `asc`). */
    function compararColuna(a, b, coluna, asc) {
        const dir = asc ? 1 : -1;
        const r = coluna === 'qtd' ? numeroPtBr(a.qtd) - numeroPtBr(b.qtd)
            : textoDaColuna(a, coluna).localeCompare(textoDaColuna(b, coluna), 'pt-BR');
        return r * dir;
    }

    /** Ordenação de exibição (não mexe no REC): texto A-Z/Z-A, quantidade numérica. Estável. */
    function ordenar(linhas, coluna, asc) {
        return linhas.map((l, i) => ({ l, i })).sort((a, b) => compararColuna(a.l, b.l, coluna, asc) || (a.i - b.i)).map(x => x.l);
    }

    /**
     * Filtro dinâmico de uma linha. f = { global, mdo, codigo, desc, filtro } (textos; vazio = sem filtro). O TIPO MDO casa tanto o
     * rótulo (M/N) quanto o nome por extenso; FILTRO é igualdade exata (como o seletor do REC).
     */
    function passaFiltro(l, f) {
        const up = s => String(s == null ? '' : s).toUpperCase();
        const g = up(f.global);
        const mdoTxt = rotuloMdo(l.mdo) + ' ' + up(l.mdo);
        if (g && (mdoTxt + ' ' + up(l.codigo) + ' ' + centroDe(l.mdo) + ' ' + up(l.desc) + ' ' + up(l.qtd) + ' ' + up(l.filtro)).indexOf(g) === -1) return false;
        if (f.mdo) {                                  // "M"/"N" sozinhos são o tipo; mais letras casam o nome por extenso
            const fm = up(f.mdo).trim();
            if (fm.length === 1 && (fm === 'M' || fm === 'N')) { if (rotuloMdo(l.mdo) !== fm) return false; }
            else if (mdoTxt.indexOf(fm) === -1) return false;
        }
        if (f.codigo && up(l.codigo).indexOf(up(f.codigo)) === -1) return false;
        if (f.centro && centroDe(l.mdo).indexOf(String(f.centro).trim()) === -1) return false;
        if (f.desc && up(l.desc).indexOf(up(f.desc)) === -1) return false;
        if (f.filtro && up(l.filtro).trim() !== up(f.filtro).trim()) return false;
        return true;
    }

    /** Colunas que o botão Copiar leva (TASK-072): a DESCRIÇÃO fica de fora. */
    const COLUNAS_COPIA = ['mdo', 'codigo', 'qtd', 'centro'];

    function valorDaColuna(l, coluna) {
        if (coluna === 'qtd') return l.qtd == null ? '' : String(l.qtd);
        return textoDaColuna(l, coluna);
    }

    /** TSV de uma tabela, sem cabeçalho: TIPO MDO (M/N) ⇥ CÓDIGO ⇥ QUANTIDADE ⇥ CENTRO (sem a descrição). `colunas` muda o conjunto. */
    function tsv(linhas, colunas) {
        const cols = colunas || COLUNAS_COPIA;
        return linhas.map(l => cols.map(c => valorDaColuna(l, c)).join('\t')).join('\n');
    }

    /** Matriz de textos (linhas × colunas) -> TSV. Usado na cópia de um intervalo selecionado. */
    function tsvDeMatriz(matriz) { return matriz.map(r => r.join('\t')).join('\n'); }

    /** Chave da marca "já copiado": tabela + item (MDO|código) + coluna. Não depende do id da linha (que muda ao refazer a tela). */
    function chaveMarca(tipo, l, coluna) {
        return tipo + '|' + rotuloMdo(l.mdo) + '|' + String(l.codigo == null ? '' : l.codigo).trim().toUpperCase() + '|' + coluna;
    }
    /** A marca só vale enquanto o valor da célula for o que foi copiado. */
    function marcaValida(marcas, chave, valorAtual) {
        return marcas.has(chave) && marcas.get(chave) === String(valorAtual == null ? '' : valorAtual).trim();
    }

    /** Linhas -> corpo do POST /api/orcamento/exportar-xlsx (quantidade numérica). */
    function paraExcel(linhas) {
        return linhas.map(l => ({ mdo: rotuloMdo(l.mdo), codigo: String(l.codigo || ''), centro: centroDe(l.mdo), desc: String(l.desc || ''), qtd: numeroPtBr(l.qtd) }));
    }

    function valoresUnicos(linhas, coluna) {
        return Array.from(new Set(linhas.map(l => textoDaColuna(l, coluna).trim()).filter(Boolean))).sort();
    }

    const api = { COLUNAS_COPIA, tsvDeMatriz, chaveMarca, marcaValida, compararColuna, centroDe, rotuloMdo, mdoDeRotulo, separar, numeroPtBr, ordenar, passaFiltro, tsv, paraExcel, valoresUnicos };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    raiz.OrcamentoSap = api;
})(typeof window !== 'undefined' ? window : globalThis);
