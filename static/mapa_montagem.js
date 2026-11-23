/*
 * mapa_montagem.js (TASK-233) — contas da montagem rápida da rede no mapa de vinculação, sem DOM (testado em Node: tests/test_mapa_montagem_233.py).
 *
 * Coordenadas em metros do CAD ou em unidades do mapa — tudo aqui só compara distâncias na mesma unidade.
 *   maisPerto        — o candidato mais perto de um ponto, até um raio (ímã da ponta do cabo, ponta solta → ponto)
 *   casar            — pares 1:1 pela menor distância (anotação sem bloco ↔ bloco livre; cada lado usado uma vez)
 *   blocosLivres     — blocos de poste do fundo do CAD que ainda não viraram ponto
 *   recortarFundo    — só o fundo perto da rede (um fundo de outro desenho não aparece)
 *   dentroDoRetangulo, metros, progresso
 */
(function (raiz) {
    'use strict';

    const dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
    const valido = p => Array.isArray(p) && p.length >= 2 && Number.isFinite(p[0]) && Number.isFinite(p[1]);

    /** `candidatos`: [{id, p:[x,y]}]. Devolve {id, d} do mais perto até `raio` (inclusive), ou null. `ignorar`: Set de ids fora da busca. */
    function maisPerto(pt, candidatos, raio, ignorar) {
        if (!valido(pt)) return null;
        let melhor = null;
        for (const c of candidatos || []) {
            if (!c || !valido(c.p) || (ignorar && ignorar.has(c.id))) continue;
            const d = dist(pt, c.p);
            if (d <= raio && (!melhor || d < melhor.d)) melhor = { id: c.id, d };
        }
        return melhor;
    }

    /**
     * Casamento 1:1 guloso pela menor distância: `as` e `bs` = [{id, p}]. Cada id entra em no máximo um par; pares acima de `raio` não contam.
     * Empate (mesma distância) desfaz pela ordem das listas — o resultado não depende de sorteio. Devolve [{a, b, d}] na ordem das distâncias.
     */
    function casar(as, bs, raio) {
        const pares = [];
        (as || []).forEach((a, ia) => (bs || []).forEach((b, ib) => {
            if (!valido(a.p) || !valido(b.p)) return;
            const d = dist(a.p, b.p);
            if (d <= raio) pares.push({ a: a.id, b: b.id, d, ia, ib });
        }));
        pares.sort((x, y) => x.d - y.d || x.ia - y.ia || x.ib - y.ib);
        const usadosA = new Set(), usadosB = new Set(), saida = [];
        for (const p of pares) {
            if (usadosA.has(p.a) || usadosB.has(p.b)) continue;
            usadosA.add(p.a); usadosB.add(p.b);
            saida.push({ a: p.a, b: p.b, d: p.d });
        }
        return saida;
    }

    /** Blocos (`[[x,y]]`) a mais de `tol` de qualquer ponto ocupado (`[[x,y]]`). Grade de busca: rápido para milhares. Devolve [{id:'b<i>', p}]. */
    function blocosLivres(blocos, ocupados, tol = 1) {
        const celula = Math.max(tol, 1e-6) * 2, grade = new Map(), chave = (i, j) => i + ',' + j;
        for (const p of ocupados || []) {
            if (!valido(p)) continue;
            const k = chave(Math.floor(p[0] / celula), Math.floor(p[1] / celula));
            if (!grade.has(k)) grade.set(k, []);
            grade.get(k).push(p);
        }
        const livres = [];
        (blocos || []).forEach((p, i) => {
            if (!valido(p)) return;
            const ci = Math.floor(p[0] / celula), cj = Math.floor(p[1] / celula);
            for (let a = ci - 1; a <= ci + 1; a++) for (let b = cj - 1; b <= cj + 1; b++)
                for (const q of grade.get(chave(a, b)) || []) if (dist(p, q) <= tol) return;
            livres.push({ id: 'b' + i, p: [p[0], p[1]] });
        });
        return livres;
    }

    /** Caixa [x0, y0, x1, y1] de pontos `[[x,y]]`, ou null. */
    function caixa(ps) {
        let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
        for (const p of ps || []) { if (!valido(p)) continue; if (p[0] < x0) x0 = p[0]; if (p[0] > x1) x1 = p[0]; if (p[1] < y0) y0 = p[1]; if (p[1] > y1) y1 = p[1]; }
        return Number.isFinite(x0) ? [x0, y0, x1, y1] : null;
    }

    /**
     * Só o que do fundo fica perto da rede: `area` = caixa da rede (metros); `folga` = quanto além dela ainda aparece (padrão: 30% do lado maior, mín. 60 m).
     * Fundo de outro desenho (que não encosta na rede) some inteiro. Devolve {redes, postes} no mesmo formato da entrada.
     */
    function recortarFundo(fundo, area, folga) {
        if (!fundo || !area) return { redes: [], postes: [] };
        const f = Number.isFinite(folga) ? folga : Math.max(60, Math.max(area[2] - area[0], area[3] - area[1]) * .3);
        const x0 = area[0] - f, y0 = area[1] - f, x1 = area[2] + f, y1 = area[3] + f;
        const dentro = p => valido(p) && p[0] >= x0 && p[0] <= x1 && p[1] >= y0 && p[1] <= y1;
        const redes = (fundo.redes || []).filter(r => r && Array.isArray(r.p) && r.p.length >= 2 && r.p.some(dentro));
        const postes = (fundo.postes || []).filter(dentro);
        return { redes, postes };
    }

    /** ids (de um Map id → [x, y]) dentro do retângulo dado por dois cantos quaisquer. */
    function dentroDoRetangulo(posicoes, a, b) {
        const x0 = Math.min(a[0], b[0]), x1 = Math.max(a[0], b[0]), y0 = Math.min(a[1], b[1]), y1 = Math.max(a[1], b[1]);
        const ids = [];
        for (const [id, p] of posicoes) if (valido(p) && p[0] >= x0 && p[0] <= x1 && p[1] >= y0 && p[1] <= y1) ids.push(id);
        return ids;
    }

    /** Comprimento de vão em metros inteiros (mínimo 1) entre dois pontos do CAD; null sem os dois. */
    function metros(a, b) {
        if (!valido(a) || !valido(b)) return null;
        const d = dist(a, b);
        return d > 0 ? Math.max(1, Math.round(d)) : null;
    }

    /** {feitos, total, pct, texto} — «47/52 (90%)». */
    function progresso(feitos, total) {
        const pct = total > 0 ? Math.round(100 * feitos / total) : 100;
        return { feitos, total, pct, texto: feitos + '/' + total + (total ? ' (' + pct + '%)' : '') };
    }

    const api = { maisPerto, casar, blocosLivres, caixa, recortarFundo, dentroDoRetangulo, metros, progresso, dist };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.MapaMontagem = api;
})(typeof window !== 'undefined' ? window : globalThis);
