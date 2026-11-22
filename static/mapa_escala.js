/*
 * mapa_escala.js (TASK-224) — contas do mapa de vinculação sem DOM (testado em Node: tests/test_mapa_escala_224.py).
 *
 * Escala: o desenho do CAD está em METROS (vãos de 20–50 m). Na tela, 1 mm de papel = 96/25.4 px CSS; então 1:500 vale 1 m = 2 mm = 7,559 px
 * (aproximado: o monitor real pode ter outro tamanho de pixel). "Ajustar" escolhe o fator para caber tudo na tela e a escala resultante é só informada.
 * Recorte (culling): redes grandes só mantêm visíveis os pontos e vãos perto da janela; o resto fica com display:none (pintar 40 mil elementos SVG a cada
 * movimento era o que travava o arrastar e o zoom).
 */
(function (raiz) {
    'use strict';

    const PX_POR_MM = 96 / 25.4;
    const ESCALAS = [['ajustar', 'Ajustar à tela'], ['500', '1:500']];
    const escalaValida = v => (ESCALAS.some(([k]) => k === String(v)) ? String(v) : 'ajustar');
    /** px por metro do desenho em escala 1:`denominador` (o padrão do mapa é 1:500 → 7,559). */
    const pxPorMetro = denominador => PX_POR_MM * 1000 / Number(denominador);
    /** Denominador (1:D) que a tela mostra agora: `kLayout` = px do mapa por metro do desenho; `pxPorUnidade` = px de tela por unidade do mapa (largura do SVG ÷ largura do viewBox). */
    function denominadorDe(kLayout, pxPorUnidade) {
        const pxMetro = kLayout * pxPorUnidade;
        return pxMetro > 0 && Number.isFinite(pxMetro) ? PX_POR_MM * 1000 / pxMetro : null;
    }
    /** Arredonda para um valor «de régua» (…, 200, 250, 500, 1000, 1500 …) e escreve «1:500». */
    function rotuloDaEscala(d) {
        if (!d || !Number.isFinite(d)) return '';
        const passo = d < 100 ? 5 : d < 1000 ? 25 : d < 10000 ? 100 : 1000;
        return '1:' + (Math.round(d / passo) * passo).toLocaleString('pt-BR');
    }
    const BARRAS = [1, 2, 5, 10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 10000];
    /** Barra de escala: o maior comprimento «redondo» (m) que cabe em `maxPx` na tela; devolve {metros, px, rotulo} (ou null sem escala). */
    function barraDeEscala(pxPorMetroNaTela, maxPx = 140) {
        if (!(pxPorMetroNaTela > 0) || !Number.isFinite(pxPorMetroNaTela)) return null;
        let m = BARRAS[0];
        for (const b of BARRAS) { if (b * pxPorMetroNaTela <= maxPx) m = b; }
        return { metros: m, px: m * pxPorMetroNaTela, rotulo: m >= 1000 ? (m / 1000) + ' km' : m + ' m' };
    }

    /** viewBox [x, y, w, h] centrado em (cx, cy) com largura/altura em unidades do mapa. */
    const vistaCentrada = (cx, cy, w, h) => [cx - w / 2, cy - h / 2, w, h];
    /** Retângulo [x0, y0, x1, y1] de um viewBox, alargado em `folga` (fração do tamanho) para cada lado. */
    function retanguloDaVista(vb, folga = 0) {
        const fx = vb[2] * folga, fy = vb[3] * folga;
        return [vb[0] - fx, vb[1] - fy, vb[0] + vb[2] + fx, vb[1] + vb[3] + fy];
    }
    const cruza = (a, b) => a[0] <= b[2] && a[2] >= b[0] && a[1] <= b[3] && a[3] >= b[1];
    const contem = (grande, pequeno) => grande[0] <= pequeno[0] && grande[1] <= pequeno[1] && grande[2] >= pequeno[2] && grande[3] >= pequeno[3];
    /** Caixa [x0, y0, x1, y1] de uma lista de pontos [[x, y]…], com `folga` em todos os lados. */
    function caixaDe(pontos, folga = 0) {
        let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
        for (const [x, y] of pontos) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
        return [x0 - folga, y0 - folga, x1 + folga, y1 + folga];
    }
    /** Mostra só o que cruza `area` (itens: {b: caixa, v: visível agora}); `mudar(item, visivel)` é chamado só para quem mudou. Devolve quantos mudaram. */
    function recortar(itens, area, mudar) {
        let n = 0;
        for (const it of itens) {
            const v = cruza(it.b, area);
            if (v !== it.v) { it.v = v; mudar(it, v); n++; }
        }
        return n;
    }

    const api = { PX_POR_MM, ESCALAS, escalaValida, pxPorMetro, denominadorDe, rotuloDaEscala, barraDeEscala, vistaCentrada, retanguloDaVista, cruza, contem, caixaDe, recortar };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.MapaEscala = api;
})(typeof window !== 'undefined' ? window : globalThis);
