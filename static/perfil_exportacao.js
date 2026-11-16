/*
 * perfil_exportacao.js (TASK-176, ADR-007) — perfil de exportação por projeto: cabeçalho do relatório, centro por tipo, PDF/XLSX e arquivos oferecidos. Sem DOM.
 * MESMA lógica de services/perfil_exportacao.py (paridade testada em tests/test_perfil_exportacao_176.py). Sem perfil, vale o de Neoenergia (PADRAO).
 */
(function (raiz) {
    'use strict';
    const COLUNAS_XLSX = { mdo: 'TIPO MDO', codigo: 'CÓDIGO', qtd: 'QUANTIDADE', centro: 'CENTRO', desc: 'DESCRIÇÃO' };
    const ARQUIVOS = ['pdf_sap', 'xlsx', 'rec', 'csv'];
    const RE_CHAVE = /^[a-z][a-z0-9_]{0,30}$/, RE_COR = /^#[0-9A-Fa-f]{6}$/;
    const trim = t => String(t == null ? '' : t).trim();
    const clone = o => JSON.parse(JSON.stringify(o));
    const C = (chave, rotulo, tela, coluna, obra, lembrar) => ({ chave, rotulo, tela, coluna, obrigatorio: false, obra: !!obra, lembrar: !!lembrar });
    const PADRAO = {
        cabecalho: [C('titulo', 'TÍTULO:', 'Título', 'esq'), C('nota', 'Nº NOTA:', 'Nº da nota', 'esq'), C('sap', 'Nº SAP:', 'Nº SAP', 'meio'), C('bairro', 'BAIRRO', 'Bairro', 'esq', true),
            C('municipio', 'MUNICIPIO:', 'Município', 'meio', true, true), C('utd', 'UTD:', 'UTD', 'esq', true, true), C('utep', 'UTEP:', 'UTEP', 'meio', true, true),
            C('projetista', 'PROJETISTA:', 'Projetista', 'esq', false, true), C('data', 'DATA:', 'Data', 'meio'), C('id_investimento', 'ID INVESTIMENTO', 'ID Investimento (ODI)', 'id', false, true),
            C('id_investimento_odd', 'ID INVESTIMENTO (ODD)', 'ID Investimento (ODD) — vazio = igual ao ODI', 'id', false, true)],
        centro: { MATERIAL: '1600', 'MAO-DE-OBRA': '1000' },
        pdf: { cor_titulo: '#00A443', rodape: 'Internal Use', rotulo_ordem: 'TIPO DE ORDEM', rotulo_id: 'ID INVESTIMENTO' },
        xlsx: { colunas: ['mdo', 'codigo', 'qtd', 'centro', 'desc'] }, arquivos: ARQUIVOS.slice(),
    };

    function normalizar(perfil) {
        const p = clone(PADRAO);
        if (!perfil || typeof perfil !== 'object' || Array.isArray(perfil)) return p;
        const campos = [];
        (Array.isArray(perfil.cabecalho) ? perfil.cabecalho : []).forEach(x => {
            if (!x || typeof x !== 'object' || !RE_CHAVE.test(trim(x.chave))) return;
            const ch = trim(x.chave), col = trim(x.coluna).toLowerCase();
            campos.push({ chave: ch, rotulo: trim(x.rotulo) || ch.toUpperCase() + ':', tela: trim(x.tela) || trim(x.rotulo).replace(/:+$/, '') || ch, coluna: ['esq', 'meio', 'id'].includes(col) ? col : 'meio',
                obrigatorio: x.obrigatorio === true, obra: x.obra === true, lembrar: x.lembrar === true });
        });
        if (campos.length) p.cabecalho = campos;
        if (perfil.centro && typeof perfil.centro === 'object' && !Array.isArray(perfil.centro) && Object.keys(perfil.centro).length) {
            p.centro = {}; Object.keys(perfil.centro).forEach(k => { if (trim(k)) p.centro[trim(k).toUpperCase()] = trim(perfil.centro[k]); });
        }
        if (perfil.pdf && typeof perfil.pdf === 'object') Object.keys(p.pdf).forEach(k => { if (trim(perfil.pdf[k])) p.pdf[k] = trim(perfil.pdf[k]); });
        const cols = perfil.xlsx && Array.isArray(perfil.xlsx.colunas) ? perfil.xlsx.colunas : null;
        if (cols && cols.length && cols.every(c => c in COLUNAS_XLSX)) p.xlsx.colunas = Array.from(new Set(cols));
        if (Array.isArray(perfil.arquivos)) p.arquivos = perfil.arquivos.filter(a => ARQUIVOS.includes(a));
        return p;
    }

    function validar(perfil) {
        const erros = [];
        if (perfil === undefined || perfil === null) return erros;
        if (typeof perfil !== 'object' || Array.isArray(perfil)) return ["'exportacao' precisa ser um objeto."];
        const cab = perfil.cabecalho;
        if (cab !== undefined && cab !== null) {
            if (!Array.isArray(cab)) erros.push("Exportação: 'cabecalho' precisa ser uma lista de campos.");
            else {
                const vistas = new Set();
                cab.forEach((x, n) => {
                    if (!x || typeof x !== 'object' || !RE_CHAVE.test(trim(x.chave))) { erros.push(`Exportação, campo ${n + 1} do cabeçalho: 'chave' precisa ser minúscula, sem espaço (ex.: contrato).`); return; }
                    if (vistas.has(x.chave)) erros.push(`Exportação: a chave '${x.chave}' do cabeçalho está repetida.`);
                    vistas.add(x.chave);
                    if (trim(x.coluna) && !['esq', 'meio', 'id'].includes(trim(x.coluna).toLowerCase())) erros.push(`Exportação, campo '${x.chave}': 'coluna' precisa ser esq, meio ou id.`);
                });
                const col = x => (x && typeof x === 'object' ? trim(x.coluna).toLowerCase() : '');
                if (cab.filter(x => col(x) === 'esq').length > 5) erros.push('Exportação: a 1ª coluna do cabeçalho do PDF comporta até 5 campos.');
                if (cab.filter(x => col(x) === '' || col(x) === 'meio').length > 4) erros.push('Exportação: a 2ª coluna do cabeçalho do PDF comporta até 4 campos.');
            }
        }
        const c = perfil.centro;
        if (c !== undefined && c !== null && (typeof c !== 'object' || Array.isArray(c) || !Object.values(c).every(v => typeof v === 'string' || Number.isInteger(v)))) erros.push('Exportação: \'centro\' precisa ser {"MATERIAL": "1600", "MAO-DE-OBRA": "1000"}.');
        const pdf = perfil.pdf;
        if (pdf !== undefined && pdf !== null) {
            if (typeof pdf !== 'object' || Array.isArray(pdf)) erros.push("Exportação: 'pdf' precisa ser um objeto.");
            else if (pdf.cor_titulo !== undefined && pdf.cor_titulo !== null && pdf.cor_titulo !== '' && !RE_COR.test(trim(pdf.cor_titulo))) erros.push("Exportação: 'pdf.cor_titulo' precisa ser #RRGGBB.");
        }
        const x = perfil.xlsx;
        if (x !== undefined && x !== null && !(typeof x === 'object' && !Array.isArray(x) && (x.colunas === undefined || x.colunas === null || (Array.isArray(x.colunas) && x.colunas.length && x.colunas.every(k => k in COLUNAS_XLSX))))) erros.push(`Exportação: 'xlsx.colunas' precisa ser uma lista com ${JSON.stringify(Object.keys(COLUNAS_XLSX)).replace(/"/g, "'").replace(/,/g, ', ')}.`);
        const a = perfil.arquivos;
        if (a !== undefined && a !== null && (!Array.isArray(a) || a.some(v => !ARQUIVOS.includes(v)))) erros.push(`Exportação: 'arquivos' precisa ser uma lista com ${JSON.stringify(ARQUIVOS).replace(/"/g, "'").replace(/,/g, ', ')}.`);
        return erros;
    }

    const camposSap = perfil => normalizar(perfil).cabecalho.map(x => x.chave).filter(k => k !== 'titulo' && k !== 'nota');
    function chavesObra(perfil) { const o = {}; normalizar(perfil).cabecalho.forEach(x => { if (x.obra) o[x.chave] = 'OBRA.' + x.chave.toUpperCase(); }); return o; }
    function centroDe(mdo, perfil) { const m = { M: 'MATERIAL', N: 'MAO-DE-OBRA' }[trim(mdo).toUpperCase()]; return m ? (normalizar(perfil).centro[m] || '') : ''; }
    const oferece = (arquivo, perfil) => normalizar(perfil).arquivos.includes(arquivo);

    /* ── Estado no navegador: o perfil do projeto aberto na tela (definido ao carregar a configuração do projeto) ── */
    let atualPerfil = null;
    const definir = perfil => { atualPerfil = perfil && typeof perfil === 'object' && !Array.isArray(perfil) ? perfil : null; };
    const atual = () => atualPerfil;
    /** Busca a configuração do projeto e guarda o perfil (sem perfil/erro = o de Neoenergia). */
    async function carregar(projetoCodigo) {
        try {
            const r = await raiz.fetch('/api/vinculacao-config?projeto_codigo=' + encodeURIComponent(projetoCodigo || 'DEFAULT'));
            definir(r.ok ? ((await r.json()).config || {}).exportacao : null);
        } catch (e) { definir(null); }
        return atualPerfil;
    }

    const api = { PADRAO, COLUNAS_XLSX, ARQUIVOS, normalizar, validar, camposSap, chavesObra, centroDe, oferece, definir, atual, carregar };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.PerfilExportacao = api;
})(typeof window !== 'undefined' ? window : globalThis);
