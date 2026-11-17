/*
 * base_sinal.js — aviso entre telas de que a BASE DE ORÇAMENTO mudou (salvou, publicou, sincronizou). Quem altera chama BaseSinal.avisar(); quem usa a base
 * (Resumo/Orçamento, sugestões, validação) chama BaseSinal.ouvir(fn) e recarrega na hora — sem esperar o cache vencer. Funciona entre abas/janelas do mesmo
 * navegador (BroadcastChannel + evento "storage") e dentro da mesma página (ouvintes locais). Só DOM/localStorage; nada vai ao servidor.
 */
(function (raiz) {
    'use strict';
    const CHAVE = 'orcacore_base_versao', CANAL = 'orcacore_base';
    const ouvintes = [];
    let canal = null, ultima = '';
    try { if (typeof raiz.BroadcastChannel === 'function') canal = new raiz.BroadcastChannel(CANAL); } catch (e) { canal = null; }

    function chamar(origem) { ouvintes.slice().forEach(fn => { try { fn(origem); } catch (e) { /* um ouvinte com erro não impede os outros */ } }); }

    /** A base mudou: avisa as outras telas (e os ouvintes desta). */
    function avisar() {
        const v = String(Date.now()) + '.' + Math.random().toString(36).slice(2, 6);
        ultima = v;
        try { raiz.localStorage.setItem(CHAVE, v); } catch (e) { /* sem localStorage: fica o canal */ }
        try { if (canal) canal.postMessage(v); } catch (e) { /* idem */ }
        chamar('local');
    }
    /** `fn(origem)` roda quando a base muda aqui ou em outra tela. Devolve uma função que cancela. */
    function ouvir(fn) { ouvintes.push(fn); return () => { const i = ouvintes.indexOf(fn); if (i >= 0) ouvintes.splice(i, 1); }; }

    // o mesmo aviso chega por DOIS caminhos (canal e storage): a versão evita recarregar duas vezes
    function recebido(v, origem) { if (!v || v === ultima) return; ultima = v; chamar(origem); }
    if (canal) canal.onmessage = e => recebido(e && e.data, 'canal');
    if (raiz.addEventListener) raiz.addEventListener('storage', e => { if (e && e.key === CHAVE) recebido(e.newValue, 'storage'); });

    const api = { avisar, ouvir, CHAVE };
    raiz.BaseSinal = api;
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
