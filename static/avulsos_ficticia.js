/* TASK-171 — a estrutura FICTÍCIA do cabo sugere os componentes avulsos (alças, conectores…) para a tabela Outros. Sem DOM.
   MESMA lógica de services/avulsos_ficticia.py (paridade testada em tests/test_avulsos_ficticia_171.py).
   O cabo novo ancora numa estrutura existente que não está na obra; o material dela (os grupos da base que CHAMAM ativos:
   `N4(cabo=%,cabo2=%):=3*ALCA(cabo) + 3*ALCA(cabo2)`) não saía. Aqui só se calcula o que sugerir (lado do cabo NOVO); a tela adiciona com um clique. */
(function (raiz) {
    'use strict';
    const noNode = typeof require !== 'undefined' && typeof module !== 'undefined';
    // resolvidos na hora do uso: no navegador este arquivo pode carregar antes dos outros dois
    const AP = new Proxy({}, { get: (_, k) => (noNode ? require('./ativo_parametros.js') : raiz.AtivoParametros)[k] });
    const VR = new Proxy({}, { get: (_, k) => (noNode ? require('./vinculacao_regras.js') : raiz.VinculacaoRegras)[k] });

    /** {BASE: [{restricoes, itens: [{ativo, qtd}]}]} — os grupos da base que chamam ativos (TASK-134), para o projeto (o do projeto vence o genérico). */
    function gruposDeChamada(linhas, projeto) {
        const p = String(projeto || '').trim().toUpperCase(), analises = [], declaradas = new Set();
        (linhas || []).forEach(l => {
            const cel = String(l.ativo || '').trim();
            if (!cel) return;
            const lp = String(l.projeto || '').trim().toUpperCase();
            if (lp && p && !lp.split('/').map(x => x.trim()).includes(p)) return;
            const a = AP.analisarCelulaBase(cel);
            if (a.erro || !a.base) return;
            analises.push([!!lp, a]);
            if (a.tipo !== 'composto') declaradas.add(a.base);
        });
        const gGen = {}, gProj = {};
        analises.forEach(([doProj, a]) => {
            if (a.tipo !== 'composto' || a.cabecalho === null || a.cabecalho === undefined) return;
            if (!declaradas.has(a.base) || !(a.condicoes && a.condicoes.length)) return;
            const restr = {}; a.condicoes.forEach(q => { restr[q.nome] = { op: q.op, alts: q.alts.slice() }; });
            const alvo = doProj ? gProj : gGen;
            alvo[a.base] = alvo[a.base] || {};
            if (!(a.cabecalho in alvo[a.base])) alvo[a.base][a.cabecalho] = { restricoes: restr, itens: a.itens };
        });
        const grupos = {};
        new Set(Object.keys(gGen).concat(Object.keys(gProj))).forEach(base => { grupos[base] = Object.values(Object.assign({}, gGen[base] || {}, gProj[base] || {})); });
        return grupos;
    }

    const nomeTela = n => AP.nomeTela(n);
    function nomeDoAlvo(a, k, assinaturas) {
        const sig = (assinaturas || {})[a.base], nomes = (sig && sig.nomes) || [];
        return k < nomes.length ? nomeTela(nomes[k]).toLowerCase() : null;
    }

    /** Itens avulsos sugeridos para o CABO ligado à estrutura fictícia: [{ativo, qtd, operacao, completar}]. */
    function avulsosDaFicticia(cabo, ficticia, cfg, prefixoFn, assinaturas, grupos) {
        const base = String(ficticia || '').trim().toUpperCase(), sig = (assinaturas || {})[base];
        if (!base || !sig || !(grupos || {})[base]) return [];
        const paramsCabo = VR.parametrosDoCaboNaEstrutura(base, cabo, cfg, prefixoFn);
        if (!Object.keys(paramsCabo).length) return [];
        const pedido = {};
        sig.nomes.forEach(n => {
            if (paramsCabo[n] && paramsCabo[n].length) pedido[n] = paramsCabo[n][0];
            else if (n in sig.padroes) pedido[n] = sig.padroes[n];
        });
        const doCabo = new Set(sig.nomes.filter(n => paramsCabo[n] && paramsCabo[n].length).map(nomeTela));
        const e = VR.infoEstrutura(base, VR.normalizarConfig(cfg)), deCabos = new Set(e.parametros ? e.parametros.map(p => nomeTela(p.nome)) : (e.parametro ? [e.parametro] : []));          // nomes que a Tabela de Estruturas dá a CABOS
        const outroLado = new Set(Array.from(deCabos).filter(n => !doCabo.has(n)));                  // os cabos que já existem (cabo2…): o que a fictícia NÃO sabe
        const valores = {}; Object.keys(pedido).forEach(n => { valores[nomeTela(n)] = pedido[n]; });
        const conhecidos = new Set(sig.nomes.map(nomeTela)), op = cabo.operacao || 'I', soma = new Map();
        grupos[base].forEach(g => {
            if (!AP.linhaVale(g.restricoes, pedido)) return;
            g.itens.forEach(it => {
                const a = AP.analisarAtivo(it.ativo);
                if (a.erro || a.tipo !== 'grupo') return;
                let usaCabo = false, usaOutro = false;
                const partes = a.params.map((q, k) => {
                    const v = q.valor.toUpperCase();
                    let textoV;
                    if (conhecidos.has(v)) {                          // o argumento é o NOME de um parâmetro da fictícia
                        let val = valores[v];
                        if (doCabo.has(v)) usaCabo = true;
                        else if (outroLado.has(v) || val === undefined) { usaOutro = true; val = undefined; }       // o cabo existente (ou parâmetro sem valor): em branco
                        textoV = val !== undefined ? val : '?';
                    } else textoV = q.valor;
                    const nome = q.nome ? nomeTela(q.nome).toLowerCase() : nomeDoAlvo(a, k, assinaturas);
                    return nome ? `${nome}=${textoV}` : textoV;
                });
                if (usaOutro && !usaCabo) return;                    // só o outro lado (o cabo existente)
                const chave = `${a.base}(${partes.join(', ')})` + '\u0000' + (usaOutro ? '1' : '0');
                soma.set(chave, (soma.get(chave) || 0) + Number(it.qtd));
            });
        });
        return Array.from(soma.entries()).map(([k, qtd]) => { const [ativo, bl] = k.split('\u0000'); return { ativo, qtd, operacao: op, completar: bl === '1' }; });
    }

    /** Uma lista por estrutura fictícia do cabo: [{ficticia, itens}] (só as que têm sugestão). */
    function avulsosDoCabo(cabo, cfg, prefixoFn, assinaturas, grupos) {
        const saida = [];
        VR.estruturasFicticiasDe(cabo).forEach(cod => { const itens = avulsosDaFicticia(cabo, cod, cfg, prefixoFn, assinaturas, grupos); if (itens.length) saida.push({ ficticia: cod, itens }); });
        return saida;
    }
    /** O texto do `ativo` da linha de Outros: a quantidade vai na notação `3-ALCA(cabo=…)` (é a que o cálculo lê; o `qtdAtivos` de Outros não entra). */
    const textoDaLinha = it => `${Number(it.qtd)}-${it.ativo}`;
    const chaveOrigem = (idxCabo, ficticia) => `${idxCabo}:${String(ficticia).toUpperCase()}`;

    /** Estado da sugestão de UM par (cabo, fictícia) frente às linhas de Outros já adicionadas: 'pendente' | 'adicionado' | 'desatualizado' (+ os índices das linhas antigas). */
    function estadoDosAvulsos(idxCabo, ficticia, itens, outros) {
        const chave = chaveOrigem(idxCabo, ficticia), existentes = [];
        const velhas = [];
        (outros || []).forEach((r, i) => { if (r && r.origemAvulso === chave) existentes.push(i); else if (r && r.origemAvulso === chave + '!velho') velhas.push(i); });          // '!velho' = linha de uma sugestão anterior (já substituída)
        if (!existentes.length) return { estado: 'pendente', linhas: [], velhas };
        const conta = l => l.slice().sort().join(';');
        const esperado = itens.filter(x => !x.completar).map(textoDaLinha), atuais = existentes.map(i => String(outros[i].ativo || '').trim());
        return { estado: conta(atuais) === conta(esperado) ? 'adicionado' : 'desatualizado', linhas: existentes, velhas };
    }

    const api = { gruposDeChamada, avulsosDaFicticia, avulsosDoCabo, chaveOrigem, estadoDosAvulsos, textoDaLinha };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.AvulsosFicticia = api;
})(typeof window !== 'undefined' ? window : globalThis);
