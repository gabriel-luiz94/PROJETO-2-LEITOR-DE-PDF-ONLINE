/*
 * mapa_layout.js (TASK-225) — «Rearranjar» do mapa de vinculação, sem DOM (testado em Node: tests/test_mapa_layout_225.py).
 *
 * Organiza o grafo de forma uniforme e limpa: cada parte conectada é desenhada por forças (molas nos vãos com o mesmo comprimento, repulsão entre pontos),
 * partindo das posições atuais (a forma geral da rede é preservada), depois os pontos vão para uma grade (alinhado e sem sobreposição) e as partes são
 * dispostas lado a lado. Tudo é determinístico (nada de sorteio): o mesmo mapa sempre gera o mesmo resultado. Pontos «fixos» (grupo bloqueado) não saem do lugar.
 */
(function (raiz) {
    'use strict';

    /** Caixa [x0, y0, x1, y1] de um conjunto de pontos. */
    function caixa(ps) {
        let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
        for (const [x, y] of ps) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
        return [x0, y0, x1, y1];
    }

    /** Partes conectadas: lista de listas de ids, na ordem dos ids dados (pontos soltos viram partes de 1). */
    function componentes(ids, arestas) {
        const vizinhos = new Map(ids.map(id => [id, []]));
        for (const [a, b] of arestas) { if (vizinhos.has(a) && vizinhos.has(b) && a !== b) { vizinhos.get(a).push(b); vizinhos.get(b).push(a); } }
        const visto = new Set(), saida = [];
        for (const id of ids) {
            if (visto.has(id)) continue;
            const parte = [], fila = [id]; visto.add(id);
            while (fila.length) { const x = fila.shift(); parte.push(x); for (const v of vizinhos.get(x)) if (!visto.has(v)) { visto.add(v); fila.push(v); } }
            saida.push(parte);
        }
        return saida;
    }

    /** Forças (Fruchterman–Reingold) numa parte conectada. `pos`: Map id → [x, y] (alterado); `fixos`: Set de ids que não mexem. */
    function forcas(parte, arestas, pos, L, fixos, iteracoes) {
        const n = parte.length; if (n < 2) return;
        const idx = new Map(parte.map((id, i) => [id, i]));
        const X = parte.map(id => pos.get(id)[0]), Y = parte.map(id => pos.get(id)[1]);
        const par = arestas.filter(([a, b]) => idx.has(a) && idx.has(b) && a !== b).map(([a, b]) => [idx.get(a), idx.get(b)]);
        const fixo = parte.map(id => fixos.has(id));
        // grade de repulsão: só pontos até 3 células de distância repelem (n grande não vira n²)
        const R = L * 0.62, celula = L * 2.5;
        let temp = L * 1.2;
        const resfriar = temp / (iteracoes + 1);
        for (let it = 0; it < iteracoes; it++) {
            const dx = new Float64Array(n), dy = new Float64Array(n);
            const grade = new Map();
            for (let i = 0; i < n; i++) { const k = Math.floor(X[i] / celula) + ',' + Math.floor(Y[i] / celula); if (!grade.has(k)) grade.set(k, []); grade.get(k).push(i); }
            for (let i = 0; i < n; i++) {
                const cx = Math.floor(X[i] / celula), cy = Math.floor(Y[i] / celula);
                for (let gx = cx - 1; gx <= cx + 1; gx++) for (let gy = cy - 1; gy <= cy + 1; gy++) {
                    const viz = grade.get(gx + ',' + gy); if (!viz) continue;
                    for (const j of viz) {
                        if (j <= i) continue;
                        let ex = X[i] - X[j], ey = Y[i] - Y[j], d = Math.hypot(ex, ey);
                        if (d < 1e-6) { ex = ((i * 7 + j * 3) % 5) - 2 || 1; ey = ((i * 3 + j * 5) % 5) - 2 || -1; d = Math.hypot(ex, ey); }
                        const f = (R * R) / d;                                   // repulsão (um pouco menor que o vão: a mola manda no comprimento)
                        dx[i] += ex / d * f; dy[i] += ey / d * f; dx[j] -= ex / d * f; dy[j] -= ey / d * f;
                    }
                }
            }
            for (const [a, b] of par) {
                const ex = X[a] - X[b], ey = Y[a] - Y[b], d = Math.hypot(ex, ey) || 1e-6;
                const f = (d * d) / L;                                           // mola: o vão tende a medir L
                dx[a] -= ex / d * f; dy[a] -= ey / d * f; dx[b] += ex / d * f; dy[b] += ey / d * f;
            }
            for (let i = 0; i < n; i++) {
                if (fixo[i]) continue;
                const d = Math.hypot(dx[i], dy[i]) || 1e-6, passo = Math.min(d, temp);
                X[i] += dx[i] / d * passo; Y[i] += dy[i] / d * passo;
            }
            temp -= resfriar;
        }
        parte.forEach((id, i) => pos.set(id, [X[i], Y[i]]));
    }

    /** Leva cada ponto para a célula livre mais próxima de uma grade de passo `g` (alinhado, sem dois pontos no mesmo lugar). Fixos ocupam a célula onde estão. */
    function grade(ids, pos, g, fixos) {
        const ocupado = new Set(), chave = (i, j) => i + ',' + j;
        for (const id of ids) if (fixos.has(id)) { const [x, y] = pos.get(id); ocupado.add(chave(Math.round(x / g), Math.round(y / g))); }
        for (const id of ids) {
            if (fixos.has(id)) continue;
            const [x, y] = pos.get(id), ci = Math.round(x / g), cj = Math.round(y / g);
            let melhor = null;
            for (let r = 0; r < 60 && !melhor; r++) {                               // anéis crescentes ao redor da célula ideal
                let md = Infinity;
                for (let i = ci - r; i <= ci + r; i++) for (let j = cj - r; j <= cj + r; j++) {
                    if (Math.max(Math.abs(i - ci), Math.abs(j - cj)) !== r || ocupado.has(chave(i, j))) continue;
                    const d = Math.hypot(i * g - x, j * g - y);
                    if (d < md) { md = d; melhor = [i, j]; }
                }
            }
            if (!melhor) melhor = [ci, cj];
            ocupado.add(chave(melhor[0], melhor[1]));
            pos.set(id, [melhor[0] * g, melhor[1] * g]);
        }
    }

    /**
     * Rearranja. `nos`: [{id, x, y}] (posições atuais, qualquer unidade); `arestas`: [[idA, idB]]; opções:
     *   comprimento — vão desejado (mesma unidade); fixos — ids que não se movem; espaco — folga entre as partes (padrão 2×comprimento).
     * Devolve Map id → [x, y]. As partes conectadas ficam lado a lado, da maior para a menor, em linhas de largura parecida.
     */
    function rearranjar(nos, arestas, opcoes = {}) {
        const L = opcoes.comprimento > 0 ? opcoes.comprimento : 60, fixos = new Set(opcoes.fixos || []), espaco = opcoes.espaco > 0 ? opcoes.espaco : L * 2;
        const pos = new Map(nos.map(n => [n.id, [n.x, n.y]]));
        const ids = nos.map(n => n.id);
        const partes = componentes(ids, arestas).sort((a, b) => b.length - a.length);
        const iteracoes = ids.length > 800 ? 40 : ids.length > 300 ? 70 : 110;
        const g = L;                                                          // passo da grade: dois pontos nunca ficam a menos disto
        const grandes = partes.filter(p => p.length > 1), soltos = partes.filter(p => p.length === 1).map(p => p[0]);
        const blocos = [];                                                           // cada parte desenhada na própria origem, depois posicionada
        for (const parte of grandes) {
            const local = new Map(parte.map(id => [id, pos.get(id).slice()]));
            const todosFixos = parte.every(id => fixos.has(id));
            if (!todosFixos) {
                forcas(parte, arestas, local, L, fixos, iteracoes);
                grade(parte, local, g, fixos);
            }
            blocos.push({ ids: parte, pos: local, fixo: parte.some(id => fixos.has(id)) });
        }
        // posiciona: partes com ponto fixo ficam onde estão; as demais vão para a direita/abaixo em linhas
        const saida = new Map();
        const ocupadas = [];
        for (const b of blocos) if (b.fixo) { for (const id of b.ids) saida.set(id, b.pos.get(id)); ocupadas.push(caixa(b.ids.map(id => b.pos.get(id)))); }
        const base = ocupadas.length ? caixa(ocupadas.flatMap(c => [[c[0], c[1]], [c[2], c[3]]])) : null;
        const todas = blocos.filter(b => !b.fixo).map(b => ({ b, c: caixa(b.ids.map(id => b.pos.get(id))) }));
        const areaTotal = todas.reduce((s, { c }) => s + (c[2] - c[0] + espaco) * (c[3] - c[1] + espaco), 0) + soltos.length * (L * L);
        const larguraLinha = Math.max(L * 6, Math.sqrt(areaTotal) * 1.4);
        let ox = base ? base[2] + espaco : 0, oy = base ? base[1] : 0, alturaLinha = 0, x0 = ox;
        for (const { b, c } of todas) {
            const w = c[2] - c[0], h = c[3] - c[1];
            if (ox - x0 > 0 && ox - x0 + w > larguraLinha) { ox = x0; oy += alturaLinha + espaco; alturaLinha = 0; }
            for (const id of b.ids) { const [x, y] = b.pos.get(id); saida.set(id, [x - c[0] + ox, y - c[1] + oy]); }
            ox += w + espaco; alturaLinha = Math.max(alturaLinha, h);
        }
        // pontos sem vão: fila de células na grade, depois das partes
        if (soltos.length) {
            if (ox - x0 > 0) { ox = x0; oy += alturaLinha + espaco; }
            const colunas = Math.max(1, Math.round(Math.sqrt(soltos.length)));
            soltos.forEach((id, i) => {
                if (fixos.has(id)) { saida.set(id, pos.get(id).slice()); return; }
                saida.set(id, [ox + (i % colunas) * L, oy + Math.floor(i / colunas) * L]);
            });
        }
        return saida;
    }

    /** Comprimento de vão típico (mediana) das arestas nas posições dadas; `padrao` se não houver. */
    function vaoTipico(pos, arestas, padrao) {
        const ds = arestas.map(([a, b]) => (pos.has(a) && pos.has(b) ? Math.hypot(pos.get(a)[0] - pos.get(b)[0], pos.get(a)[1] - pos.get(b)[1]) : 0)).filter(d => d > 0).sort((a, b) => a - b);
        return ds.length ? ds[Math.floor(ds.length / 2)] : padrao;
    }

    const api = { componentes, rearranjar, vaoTipico, caixa };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.MapaLayout = api;
})(typeof window !== 'undefined' ? window : globalThis);
