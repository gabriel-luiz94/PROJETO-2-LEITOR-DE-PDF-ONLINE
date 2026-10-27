/*
 * ler_projeto.js (TASK-062) — "Ler projeto" direto na aba Resumo.
 *
 * Reproduz, no navegador e sem trocar de aba, o caminho Leitor -> "Processar" -> Resumo:
 *   1) passo do Leitor  : RegrasLeitorEngine.processarEClassificar em cada texto extraído;
 *   2) passo do Resumo  : computeRowLogic (confia em entidade/ativo já definidos; senão reprocessa),
 *                         separação em Cabos / Outros / Ramais.
 * É um módulo puro (sem DOM) para poder ser testado em Node contra o oráculo (tests/test_ler_projeto.py).
 * Quem chama (resumo.js) cuida de tela, histórico, localStorage e aprendizado.
 */
(function (raiz) {
    'use strict';

    const EXTENSOES = ['.pdf', '.dxf', '.dwg'];

    function extensaoValida(nome) {
        const n = String(nome || '').toLowerCase();
        return EXTENSOES.some(e => n.endsWith(e));
    }

    const cedeVez = () => new Promise(r => setTimeout(r, 0));

    /** Passo do Leitor: itens crus do /upload -> linhas como o "Processar" grava em processar_dados. */
    async function passoLeitor(itens, proc, cls, motor) {
        const saida = [];
        for (let i = 0; i < itens.length; i++) {
            const item = itens[i];
            const r = motor.processarEClassificar({
                texto: item.texto, cor: item.cor || '#000000', layer: item.layer || '',
                pagina: item.pagina, index: i, allItems: itens
            }, proc, cls);
            const linha = {
                pagina: item.pagina, texto: item.texto, cor: item.cor, layer: item.layer || '',
                entidade: r.entidade, operacao: r.operacao, ativo: r.ativo
            };
            if (item._x !== undefined && item._y !== undefined) { linha._x = item._x; linha._y = item._y; }
            if (item._geo) { linha._geo = item._geo; linha._geoTipo = item._geoTipo; }      // TASK-091
            if (item._poste) linha._poste = item._poste;
            saida.push(linha);
            if (i % 500 === 499) await cedeVez();
        }
        return saida;
    }

    /** Mesma regra do computeRowLogic do resumo.js. */
    function logicaLinha(item, indice, exportados, proc, cls, motor) {
        if (item.entidade && item.entidade !== '0' && item.ativo && item.ativo.trim() !== '') {
            return { entidade: item.entidade, operacao: item.operacao || 'M', ativo: item.ativo, _raw: item };
        }
        const r = motor.processarEClassificar({
            texto: item.texto || '', cor: item.cor || '#000000', layer: item.layer || '',
            pagina: item.pagina, index: indice, allItems: exportados
        }, proc, cls);
        return { entidade: r.entidade, operacao: r.operacao, ativo: r.ativo, _raw: item };
    }

    function separar(processadas) {
        const cabos = processadas.filter(r => r.entidade === 'CABO');
        const outros = processadas.filter(r => r.entidade !== 'CABO' && r.entidade !== '0' && r.entidade !== 'RAMAIS');
        const ramais = processadas
            .filter(r => {
                if (r.entidade === 'RAMAIS' || r.entidade === 'IP') return true;
                if (r.entidade === 'APOIO') {
                    const txt = ((r._raw && r._raw.texto) || r.ativo || '').toUpperCase();
                    return /REC.*CAL[CÇ]ADA/i.test(txt) || /CONC.*BASE/i.test(txt);
                }
                return false;
            })
            .map(r => ({
                entidade: r.entidade,
                texto: (r._raw && r._raw.texto) || r.ativo || '',
                _textoOriginal: (r._raw && r._raw.texto) || '',
                _raw: r._raw || null,
                pagina: r._raw && r._raw.pagina != null ? r._raw.pagina : '-'
            }));
        const padrao = r => { r.entidade = r.entidade || '0'; r.operacao = r.operacao || 'M'; r.ativo = r.ativo || ''; return r; };
        return { cabos: cabos.map(padrao), outros: outros.map(padrao), ramais };
    }

    /** itens (do /upload) + regras -> { exportados, cabos, outros, ramais }. */
    async function processar(itens, proc, cls, motor) {
        const exportados = await passoLeitor(itens, proc || [], cls || [], motor);
        const processadas = [];
        for (let i = 0; i < exportados.length; i++) {
            processadas.push(logicaLinha(exportados[i], i, exportados, proc || [], cls || [], motor));
            if (i % 500 === 499) await cedeVez();
        }
        const { cabos, outros, ramais } = separar(processadas);
        return { exportados, cabos, outros, ramais };
    }

    /** Envia o arquivo ao /upload (mesmo endpoint do Leitor). Devolve a lista de itens ou lança Error com mensagem amigável. */
    async function extrair(arquivo, chaveCloudConvert) {
        if (!extensaoValida(arquivo && arquivo.name)) throw new Error('Formato não suportado. Use PDF, DXF ou DWG.');
        const fd = new FormData();
        fd.append('file', arquivo);
        const resp = await fetch('/upload', { method: 'POST', body: fd, headers: chaveCloudConvert ? { 'X-CloudConvert-Key': chaveCloudConvert } : {} });
        if (!resp.ok) throw new Error(`Falha ao enviar o arquivo (HTTP ${resp.status}).`);
        const json = await resp.json();
        if (json.error) throw new Error(json.error);
        const itens = json.data || [];
        try { if (json.geometria) localStorage.setItem('mapa_geometria_fundo', JSON.stringify(json.geometria)); else localStorage.removeItem('mapa_geometria_fundo'); } catch (e) { /* sem espaço: o mapa só fica sem a rede ao fundo */ }     // TASK-091: rede/postes ao fundo do mapa
        if (!itens.length) throw new Error('Não encontrei texto neste arquivo; nada foi alterado.');
        return itens;
    }

    const api = { EXTENSOES, extensaoValida, processar, extrair, separar };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    raiz.LerProjeto = api;
})(typeof window !== 'undefined' ? window : globalThis);
