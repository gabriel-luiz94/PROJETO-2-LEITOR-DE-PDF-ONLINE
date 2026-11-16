/* TASK-165 — parâmetros de obra (OBRA.*): MESMA regra de services/obra_params.py (RULES Regra 4).
   Valores vêm do cabeçalho da obra (Projeto → Informações do relatório SAP) e do contexto selecionado; OBRA.REGIAO é derivado pela tabela regional
   (GET /api/orcamento/regionais, tirada da ICP). */
(function (raiz) {
    'use strict';
    const CAMPOS = { municipio: 'OBRA.MUNICIPIO', bairro: 'OBRA.BAIRRO', utd: 'OBRA.UTD', utep: 'OBRA.UTEP', contexto: 'OBRA.CONTEXTO' };
    /** TASK-176: campos do perfil de exportação do projeto marcados `obra` (padrão: município, bairro, UTD, UTEP) + contexto. */
    function camposDoPerfil(perfil) {
        const PE = raiz.PerfilExportacao;
        if (!perfil || !PE) return CAMPOS;
        return Object.assign(PE.chavesObra(perfil), { contexto: 'OBRA.CONTEXTO' });
    }
    const NOMES = ['OBRA.MUNICIPIO', 'OBRA.BAIRRO', 'OBRA.UTD', 'OBRA.UTEP', 'OBRA.REGIAO', 'OBRA.CONTEXTO'];

    /** Maiúsculas, sem acento, espaços simples e espaço → _ ("Centro Oeste" → CENTRO_OESTE, "34,5kV" → 34,5KV). */
    function normalizar(valor) {
        return String(valor == null ? '' : valor).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim().replace(/\s+/g, '_');
    }
    const chave = valor => normalizar(valor).replace(/_/g, ' ');

    /** OBRA.REGIAO: utep do cabeçalho (se for região da tabela) → utd → município(+bairro) → '' (ausente). */
    function regiaoDe(obra, tabela) {
        obra = obra || {};
        const t = tabela || {}, regioes = t.regioes || [];
        if (!regioes.length) return '';
        const utep = normalizar(obra.utep);
        if (regioes.includes(utep)) return utep;
        const utd = chave(obra.utd);
        if (utd && (t.por_utd || {})[utd]) return t.por_utd[utd];
        const mun = chave(obra.municipio), bairro = chave(obra.bairro);
        if (mun) {
            if (bairro) { const x = (t.por_municipio_bairro || {})[mun + '|' + bairro]; if (x) return x.regiao; }
            return (t.por_municipio || {})[mun] || '';
        }
        return '';
    }

    /** {'OBRA.X': valor normalizado} só do que a obra tem. */
    function montar(obra, tabela, perfil) {
        obra = obra && typeof obra === 'object' ? obra : {};
        const saida = {}, campos = camposDoPerfil(perfil);
        Object.keys(campos).forEach(c => { const v = normalizar(obra[c]); if (v) saida[campos[c]] = v; });
        const r = regiaoDe(obra, tabela);
        if (r) saida['OBRA.REGIAO'] = r;
        return saida;
    }

    /* ── Estado no navegador ── */
    let tabelaAtual = null, tabelaProjeto = null;
    /** Carrega (uma vez por projeto) a tabela regional do projeto; sem tabela devolve {}. */
    async function carregarTabela(projeto) {
        const p = String(projeto || '').trim();
        if (tabelaProjeto === p && tabelaAtual) return tabelaAtual;
        try {
            const r = await raiz.fetch('/api/orcamento/regionais?projeto=' + encodeURIComponent(p));
            tabelaAtual = r.ok ? await r.json() : {};
        } catch (e) { tabelaAtual = {}; }
        tabelaProjeto = p;
        return tabelaAtual;
    }
    const definirTabela = (t, projeto) => { tabelaAtual = t || {}; tabelaProjeto = projeto == null ? null : String(projeto); };

    /** Cabeçalho + contexto da obra aberta: {municipio, bairro, utd, utep, contexto} (só o que tem valor). */
    function obraDoProjeto() {
        const saida = {};
        try {
            const sap = (raiz.GridRProjeto && raiz.GridRProjeto.sap && raiz.GridRProjeto.sap()) || {};
            Object.keys(camposDoPerfil(raiz.PerfilExportacao ? raiz.PerfilExportacao.atual() : null)).filter(k => k !== 'contexto').forEach(k => { const v = String(sap[k] || '').trim(); if (v) saida[k] = v; });
        } catch (e) { /* sem projeto aberto */ }
        try { const c = raiz.contextoSelecionado ? raiz.contextoSelecionado() : null; if (c) saida.contexto = String(c); } catch (e) { /* sem seletor */ }
        return saida;
    }
    /** Valores OBRA.* da obra aberta, para o resolverPedido da tela (a tabela já carregada decide OBRA.REGIAO). */
    const valores = () => montar(obraDoProjeto(), tabelaAtual, raiz.PerfilExportacao ? raiz.PerfilExportacao.atual() : null);

    try { if (raiz.document && raiz.localStorage) carregarTabela(raiz.localStorage.getItem('projeto_selecionado') || ''); } catch (e) { /* sem storage: a tela carrega depois */ }      // já na abertura da tela

    const api = { NOMES, normalizar, regiaoDe, montar, carregarTabela, definirTabela, obraDoProjeto, valores };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.ObraParams = api;
})(typeof window !== 'undefined' ? window : globalThis);
