/*
 * ativo_itens.js (TASK-186) — o texto do ativo de Outros como LISTA DE ITENS ("12600D, N1(XLPE,D,A), N4" → 12600D · N1 com parâmetros · N4).
 * Serve ao modo «Ocultar parâmetros»: o campo mostra e edita só os nomes ("12600D, N1, N4") e os parâmetros ficam PRESOS a cada item enquanto ele existir.
 * Apagou ou renomeou o item = é um item novo (sem parâmetros). Sem DOM; testado em Node (tests/test_ativo_itens_186.py).
 */
(function (raiz) {
    'use strict';

    /** Divide no 1º nível (vírgula fora de ( ) e [ ]): [{ini, fim, bruto}] com as posições do trecho SEM espaços nas pontas; trechos vazios saem. */
    function tokens(texto) {
        const t = String(texto == null ? '' : texto), saida = [];
        let prof = 0, colchete = false, ini = 0;
        const fecha = fim => {
            let a = ini, b = fim;
            while (a < b && /\s/.test(t[a])) a++;
            while (b > a && /\s/.test(t[b - 1])) b--;
            if (b > a) saida.push({ ini: a, fim: b, bruto: t.slice(a, b) });
        };
        for (let i = 0; i < t.length; i++) {
            const c = t[i];
            if (c === '[') colchete = true;
            else if (c === ']') colchete = false;
            else if (!colchete && c === '(') prof++;
            else if (!colchete && c === ')') prof = Math.max(0, prof - 1);
            else if (!colchete && prof === 0 && c === ',') { fecha(i); ini = i + 1; }
        }
        fecha(t.length);
        return saida;
    }

    /** "N1(XLPE,D,A)" → {nome: "N1", params: "(XLPE,D,A)"}. O "(" tem de vir colado ao nome; um "(" no começo ou sem ")" fica tudo no nome/params como está. */
    function separar(bruto) {
        let colchete = false;
        for (let i = 0; i < bruto.length; i++) {
            const c = bruto[i];
            if (c === '[') colchete = true;
            else if (c === ']') colchete = false;
            else if (!colchete && c === '(' && i > 0) return { nome: bruto.slice(0, i), params: bruto.slice(i) };
        }
        return { nome: bruto, params: '' };
    }

    /** O nome do ativo SEM a quantidade: "6 CON" → "CON", "6-CON" → "CON", "*1,5-IT_2" → "IT_2". Postes e códigos que só começam com dígitos ("12600D", "3XC16") ficam como estão. */
    const semQuantidade = nome => String(nome == null ? '' : nome).trim().replace(/^\*?\d+(?:[.,]\d+)?(?:-(?=\S)|\s+(?=\S))/, '');
    const mk = (nome, params) => ({ nome, base: semQuantidade(nome), params: params || '', raw: nome + (params || '') });

    /** Itens do texto cheio: [{nome, base, params, raw}] — `nome` é como está escrito (com a quantidade), `base` é o ativo procurado na base. */
    function itens(texto) {
        return tokens(texto).map(k => { const s = separar(k.bruto); return mk(s.nome, s.params); });
    }
    const juntar = lista => lista.map(i => i.nome + (i.params || '')).join(', ');
    /** Texto só com os nomes: "12600D, N1(XLPE,D,A), N4" → "12600D, N1, N4". */
    const compacto = texto => itens(texto).map(i => i.nome).join(', ');

    /** Dentro dos parênteses de um item ("(XLPE,D,A)" → "XLPE,D,A"). */
    function conteudo(params) {
        const p = String(params || '');
        if (!p) return '';
        return p.slice(1, p.endsWith(')') ? -1 : undefined);
    }
    /** Põe (ou tira, com texto vazio) os parâmetros de um item; "()" vazio vira nenhum. */
    function comParams(lista, i, params) {
        return lista.map((it, k) => (k === i ? mk(it.nome, params === '()' ? '' : (params || '')) : it));
    }

    /** Onde o texto novo começa e termina de diferir do antigo (o cursor decide quando há trechos repetidos: "N1, N1" apagando o 1º ou o 2º). */
    function regiao(antigo, novo, cursor) {
        let p = 0;
        const max = Math.min(antigo.length, novo.length);
        while (p < max && antigo[p] === novo[p]) p++;
        if (cursor != null && cursor >= 0) p = Math.min(p, cursor);
        let s = 0;
        const maxS = Math.min(antigo.length - p, novo.length - p);
        while (s < maxS && antigo[antigo.length - 1 - s] === novo[novo.length - 1 - s]) s++;
        return { p, s };
    }

    /**
     * O usuário editou o texto compacto. `antigos` = itens de antes (alinhados com `compactoAntigo`), `compactoNovo` = o que está no campo agora, `cursor` = posição do cursor.
     * Item que ficou inteiro antes ou depois do trecho alterado mantém os parâmetros; o que foi tocado (renomeado, digitado, colado) é um item novo.
     * Devolve {itens, texto} — `texto` é o texto cheio (nomes + parâmetros) para guardar na linha.
     */
    function aplicarEdicao(antigos, compactoAntigo, compactoNovo, cursor) {
        const velho = tokens(compactoAntigo), novoT = tokens(compactoNovo), r = regiao(String(compactoAntigo), String(compactoNovo), cursor);
        const lenV = String(compactoAntigo).length, lenN = String(compactoNovo).length;
        const lista = novoT.map(k => {
            const s = separar(k.bruto);
            if (s.params) return mk(s.nome, s.params);                      // colou "N1(XLPE,D,A)": os parâmetros vêm junto
            let origem = -1;
            if (k.fim <= r.p) origem = novoT.indexOf(k);                                               // inteiro antes do trecho alterado: mesma posição
            else if (k.ini >= lenN - r.s) origem = velho.findIndex(v => v.ini === k.ini + (lenV - lenN));   // inteiro depois: mesma distância do fim
            const it = origem >= 0 ? antigos[origem] : null;
            return it && it.nome === s.nome ? mk(it.nome, it.params) : mk(s.nome, '');
        });
        return { itens: lista, texto: juntar(lista) };
    }

    const api = { semQuantidade, tokens, separar, itens, juntar, compacto, conteudo, comParams, regiao, aplicarEdicao };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.AtivoItens = api;
})(typeof window !== 'undefined' ? window : globalThis);
