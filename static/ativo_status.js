/* TASK-167 — status dos parâmetros de um ativo: quais estão preenchidos (e de onde veio o valor) e quais faltam. Sem DOM.
   Origens: 'pedido' (escrito no ativo), 'vinculo' (cabo vinculado à estrutura), 'obra' (cabeçalho/contexto da obra, OBRA.*), 'padrao' (padrão da declaração na base).
   O padrão CONTA como preenchido (aparece com "· padrão" no balão). Falta = sem valor e sem padrão. */
(function (raiz) {
    'use strict';
    const AP = typeof require !== 'undefined' && typeof module !== 'undefined' ? require('./ativo_parametros.js') : raiz.AtivoParametros;
    const ehObra = n => String(n == null ? '' : n).toUpperCase().startsWith('OBRA.');
    const nomeTela = n => String(n).replace(/#\d+$/, '');

    /** Nomes de parâmetro que o TEXTO liga (por nome ou por posição; OBRA.* nunca é do pedido). */
    function ligadosPeloTexto(texto, nomes) {
        const a = AP.analisarAtivo(texto), ligados = new Set();
        let pos = 0;
        if (a.erro) return ligados;
        a.params.forEach(p => {
            if (ehObra(p.nome)) return;
            if (p.nome === null) {
                while (pos < nomes.length && (ligados.has(nomes[pos]) || ehObra(nomes[pos]))) pos++;
                if (pos < nomes.length) { ligados.add(nomes[pos]); pos++; }
            } else ligados.add(p.nome);
        });
        return ligados;
    }

    const motivoDaFalta = nome => (/^(CABO|TIPOCABO)\d*$/.test(nomeTela(nome)) ? 'nenhum cabo vinculado a esta estrutura' : (ehObra(nome) ? 'a obra não tem esse dado' : 'não informado no pedido'));

    /**
     * @param textoFinal   o ativo como o cálculo o enxerga (com o que o vínculo injetou)
     * @param textoPedido  o ativo como foi escrito (sem o que o vínculo injetou); omitido = igual ao final
     * @returns null (ativo sem parâmetros) | {ativo, texto, params: [{nome, valor, origem, falta, invalido, motivo}], total, faltam, invalidos, completo}
     */
    function statusDoAtivo(textoFinal, textoPedido, assinaturas, obra, extra) {          // extra (TASK-170): { define: {NOME: {valor, origem}}, faltou: {NOME: motivo} } do `define` da Tabela de Estruturas
        const a = AP.analisarAtivo(textoFinal);
        if (a.erro || a.tipo === 'declaracao') return null;
        const sig = (assinaturas || {})[a.base];
        if (!sig || !sig.nomes || !sig.nomes.length) return null;
        const ped = AP.resolverPedido(textoFinal, assinaturas, obra === undefined ? null : obra);
        const final = ligadosPeloTexto(textoFinal, sig.nomes), escrito = ligadosPeloTexto(textoPedido == null ? textoFinal : textoPedido, sig.nomes);
        const params = sig.nomes.map(n => {
            const valor = ped.params[n] === undefined ? null : ped.params[n];
            let origem = null;
            if (ehObra(n)) origem = ped.obra && ped.obra[n] ? (ped.obra[n].origem === 'obra' ? 'obra' : 'padrao') : null;
            else if (escrito.has(n)) origem = 'pedido';
            else if (final.has(n)) origem = (extra && extra.define && extra.define[n] && extra.define[n].valor === valor) ? 'define' : 'vinculo';
            else if (n in (sig.padroes || {})) origem = 'padrao';
            const falta = valor === null, detalhe = origem === 'define' ? extra.define[n].origem : '';
            const invalido = !falta && (ped.erros || []).some(e => e.indexOf(nomeTela(n) + '=') === 0);
            return { nome: nomeTela(n).toLowerCase(), valor, origem: falta ? null : origem, detalhe, falta, invalido, motivo: falta ? ((extra && extra.faltou && extra.faltou[n]) ? 'não definido (' + extra.faltou[n] + ')' : motivoDaFalta(n)) : (invalido ? 'valor não cadastrado na base' : '') };
        });
        const faltam = params.filter(p => p.falta).length, invalidos = params.filter(p => p.invalido).length;
        return { ativo: a.base, texto: ped.texto, params, total: params.length, faltam, invalidos, completo: faltam === 0 && invalidos === 0 };
    }

    /** Resume vários status (uma linha pode ter várias estruturas): 'ok' (todas completas), 'falta' (alguma incompleta) ou null (nenhuma tem parâmetros). */
    function estadoDaLinha(status) {
        const l = (status || []).filter(Boolean);
        if (!l.length) return null;
        return l.every(s => s.completo) ? 'ok' : 'falta';
    }

    const ROTULO_ORIGEM = { pedido: 'manual', vinculo: 'vínculo', obra: 'da obra', padrao: 'padrão', define: 'automático' };
    function rotuloOrigem(o, detalhe) { return (ROTULO_ORIGEM[o] || '') + (o === 'define' && detalhe ? ' ‹' + detalhe + '›' : ''); }

    const api = { statusDoAtivo, estadoDaLinha, rotuloOrigem, ligadosPeloTexto };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.AtivoStatus = api;
})(typeof window !== 'undefined' ? window : globalThis);
