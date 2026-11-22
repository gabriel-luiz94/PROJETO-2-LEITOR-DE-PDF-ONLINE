/*
 * grupos_linhas.js (TASK-195) — grupos de linhas nas tabelas Outros e Cabos, sem DOM (testado em Node: tests/test_grupos_linhas_195.py).
 *
 * Uma linha cujo Ativo começa com «@» (ex.: «@P1 Rua das Flores») DECLARA um grupo; o «#» continua sendo comentário. A declaração é marcada como linha
 * CONGELADA (mesmo mecanismo do comentário: fica fora da Totalizadora, do orçamento, da validação, do Resumo da rede e dos algoritmos) e guarda um id estável (`gid`).
 * Os MEMBROS ficam sempre juntos, logo abaixo da declaração, e apontam para ela (`grupoDe`). Recolher o grupo (`grupoRecolhido` na declaração) só esconde as linhas.
 * Sem nenhuma linha «@» nada aqui faz efeito. As operações recebem a lista de linhas e devolvem uma lista NOVA (as linhas são os mesmos objetos) mais a linha que deve
 * receber o foco; quem chama usa `mapa(antiga, nova)` para acertar os vínculos por índice (como a ordenação por coluna faz).
 */
(function (raiz) {
    'use strict';

    const RE_DECL = /^\s*@\S/;
    const ehDeclaracao = t => RE_DECL.test(String(t == null ? '' : t));
    let _seq = 0;
    /** Id estável e (na prática) único: sobrevive a salvar/abrir a obra e ao Ctrl+Z. */
    function novoId() { return 'g' + Date.now().toString(36) + (++_seq).toString(36) + Math.random().toString(36).slice(2, 6); }
    /** «@P1 Rua das Flores» → «P1»; «@Trecho A» → «Trecho». */
    const nomeDe = t => String(t == null ? '' : t).trim().replace(/^@+\s*/, '').split(/\s+/)[0] || '';
    /** O que a lista de grupos mostra: o texto da declaração sem o «@» («P1 RUA DAS FLORES»). */
    const rotuloDe = row => String((row && row.ativo) || '').trim().replace(/^@+\s*/, '');

    const eDecl = r => !!(r && r.grupoDecl);
    function indiceDaDeclaracao(lista, gid) { return lista.findIndex(r => eDecl(r) && r.gid === gid); }

    /** Marca a linha como declaração quando o texto passa a começar com «@» e tira a marca quando deixa de começar. Devolve null (nada mudou) ou {tipo: 'virou'|'deixou', gid}.
     *  Linha bloqueada não muda. Quem chama solta os membros quando devolve 'deixou' (`liberar`). */
    function sincronizar(row, anterior) {
        if (!row || row.bloqueada) return null;
        if (ehDeclaracao(row.ativo)) {
            if (row.grupoDecl) return null;
            row.grupoDecl = true; row.congelada = true; row.gid = row.gid || novoId();
            delete row.grupoDe; delete row.comentario;
            return { tipo: 'virou', gid: row.gid };
        }
        if (row.grupoDecl) {
            const gid = row.gid;
            delete row.grupoDecl; delete row.congelada; delete row.gid; delete row.grupoRecolhido;
            return { tipo: 'deixou', gid };
        }
        return null;
    }

    /** Os membros do grupo deixam de pertencer a ele (continuam na tabela). */
    function liberar(lista, gid) { lista.forEach(r => { if (r && r.grupoDe === gid) delete r.grupoDe; }); }

    /** {h, fim}: índice da declaração e da última linha do bloco do grupo a que a linha `i` pertence (null se não está em grupo). */
    function bloco(lista, i) {
        const r = lista[i]; if (!r) return null;
        const gid = eDecl(r) ? r.gid : r.grupoDe;
        if (!gid) return null;
        const h = indiceDaDeclaracao(lista, gid); if (h < 0) return null;
        let fim = h; while (lista[fim + 1] && lista[fim + 1].grupoDe === gid) fim++;
        return { h, fim, gid };
    }
    const membrosDe = (lista, gid) => lista.filter(r => r.grupoDe === gid);
    const grupos = lista => lista.filter(eDecl);
    const temGrupos = lista => lista.some(eDecl);

    /** Membros sempre logo abaixo da declaração; marca órfã (grupo que não existe mais) some. Devolve uma lista nova. */
    function normalizar(lista) {
        const vistos = new Set();
        lista.forEach(r => { if (eDecl(r)) { if (vistos.has(r.gid)) r.gid = novoId(); vistos.add(r.gid); } });
        lista.forEach(r => { if (r.grupoDe && (!vistos.has(r.grupoDe) || eDecl(r))) delete r.grupoDe; });
        const porGrupo = new Map();
        lista.forEach(r => { if (r.grupoDe) { if (!porGrupo.has(r.grupoDe)) porGrupo.set(r.grupoDe, []); porGrupo.get(r.grupoDe).push(r); } });
        const saida = [];
        lista.forEach(r => { if (r.grupoDe) return; saida.push(r); if (eDecl(r) && porGrupo.has(r.gid)) saida.push(...porGrupo.get(r.gid)); });
        return saida;
    }
    /** antiga[i] foi parar em mapa[i] na nova lista (linhas que saíram ficam undefined). */
    function mapa(antiga, nova) { const pos = new Map(); nova.forEach((r, k) => pos.set(r, k)); return antiga.map(r => pos.get(r)); }

    /* ── Enter / Shift+Enter ── */
    /** Enter numa linha de grupo (declaração ou membro): a linha nova é MEMBRO e entra no fim do grupo. */
    function novaLinhaNoGrupo(lista, i, nova) {
        const b = bloco(lista, i); if (!b) return null;
        nova.grupoDe = b.gid;
        const res = lista.slice(); res.splice(b.fim + 1, 0, nova);
        return { lista: res, foco: nova };
    }
    /** Shift+Enter numa linha de grupo: a linha nova fica SOLTA, imediatamente abaixo do grupo (mesmo lugar do Enter, sem pertencer a ele). */
    function novaLinhaSolta(lista, i, nova) {
        const b = bloco(lista, i); if (!b) return null;
        delete nova.grupoDe;
        const res = lista.slice(); res.splice(b.fim + 1, 0, nova);
        return { lista: res, foco: nova };
    }

    /* ── menu de contexto ── */
    function selecionadas(lista, idxs) { return Array.from(new Set(idxs)).filter(i => lista[i] && !eDecl(lista[i])).sort((a, b) => a - b).map(i => lista[i]); }
    /** Adicionar ao grupo: as linhas (que não são declarações) vão para o fim do grupo `gid`. */
    function adicionarAoGrupo(lista, idxs, gid) {
        const linhas = selecionadas(lista, idxs); if (!linhas.length) return null;
        const fora = new Set(linhas), res = lista.filter(r => !fora.has(r));
        const h = indiceDaDeclaracao(res, gid); if (h < 0) return null;
        linhas.forEach(r => { r.grupoDe = gid; });
        let fim = h; while (res[fim + 1] && res[fim + 1].grupoDe === gid) fim++;
        res.splice(fim + 1, 0, ...linhas);
        return { lista: normalizar(res), foco: linhas[0], n: linhas.length };
    }
    /** Criar grupo: a declaração entra na posição da primeira linha selecionada e as linhas passam a ser dela. `fabricar(texto)` cria a linha da declaração (com os campos da tabela). */
    function criarGrupo(lista, idxs, nome, fabricar) {
        const linhas = selecionadas(lista, idxs); if (!linhas.length) return null;
        const nomeLimpo = String(nome == null ? '' : nome).trim().replace(/^@+\s*/, '').trim(); if (!nomeLimpo) return null;
        const fora = new Set(linhas), primeiro = lista.findIndex(r => fora.has(r));
        const antes = lista.slice(0, primeiro).filter(r => !fora.has(r)).length;
        const res = lista.filter(r => !fora.has(r));
        const decl = fabricar('@' + nomeLimpo); decl.grupoDecl = true; decl.congelada = true; decl.gid = novoId(); delete decl.grupoDe; delete decl.comentario;
        linhas.forEach(r => { r.grupoDe = decl.gid; });
        res.splice(antes, 0, decl, ...linhas);
        return { lista: normalizar(res), foco: decl, n: linhas.length };
    }
    /** Tirar do grupo: as linhas deixam de ser membros e ficam logo abaixo do bloco. */
    function tirarDoGrupo(lista, idxs) {
        const saem = Array.from(new Set(idxs)).filter(i => lista[i] && lista[i].grupoDe).sort((a, b) => a - b).map(i => lista[i]); if (!saem.length) return null;
        const por = new Map(); saem.forEach(r => { if (!por.has(r.grupoDe)) por.set(r.grupoDe, []); por.get(r.grupoDe).push(r); });
        const fora = new Set(saem), res = lista.filter(r => !fora.has(r));
        por.forEach((rs, gid) => {
            rs.forEach(r => { delete r.grupoDe; });
            const h = indiceDaDeclaracao(res, gid); let fim = h; while (res[fim + 1] && res[fim + 1].grupoDe === gid) fim++;
            res.splice(fim + 1, 0, ...rs);
        });
        return { lista: normalizar(res), foco: saem[0], n: saem.length };
    }
    /** Desfazer o grupo: a declaração sai e os membros ficam como linhas comuns. */
    function desfazerGrupo(lista, gid) {
        const h = indiceDaDeclaracao(lista, gid); if (h < 0) return null;
        const membros = membrosDe(lista, gid); liberar(lista, gid);
        return { lista: lista.filter((r, k) => k !== h), foco: membros[0] || null, n: membros.length };
    }

    /* ── Alt+↑ / Alt+↓ ── */
    /** Move a linha (ou, numa declaração, o grupo todo) uma posição. Sair do grupo por cima/por baixo tira a linha dele; descer sobre uma declaração aberta a faz o 1º membro;
     *  grupos recolhidos são atravessados de uma vez. Devolve null quando não há para onde ir. */
    function mover(lista, i, dir) {
        const r = lista[i], n = lista.length; if (!r) return null;
        const recolhido = gid => lista.some(x => eDecl(x) && x.gid === gid && x.grupoRecolhido);
        if (eDecl(r)) {
            const b = bloco(lista, i), len = b.fim - b.h + 1, bloc = lista.slice(b.h, b.fim + 1), res = lista.slice(0, b.h).concat(lista.slice(b.fim + 1));
            if (dir < 0) {
                if (b.h === 0) return null;
                const alvo = lista[b.h - 1], ini = alvo.grupoDe ? bloco(lista, b.h - 1).h : b.h - 1;
                res.splice(ini, 0, ...bloc);
            } else {
                const k = b.fim + 1; if (k >= n) return null;
                const ue = eDecl(lista[k]) ? bloco(lista, k).fim : k;
                res.splice(ue - len + 1, 0, ...bloc);
            }
            return { lista: normalizar(res), foco: r };
        }
        let j = i + dir;
        while (j >= 0 && j < n && lista[j].grupoDe && recolhido(lista[j].grupoDe)) j += dir;
        if (j < 0 || j >= n) return null;
        const t = lista[j], res = lista.slice(); res.splice(i, 1);
        if (dir < 0) {
            if (r.grupoDe && eDecl(t) && t.gid === r.grupoDe) { delete r.grupoDe; res.splice(j, 0, r); }
            else if (!r.grupoDe && eDecl(t)) res.splice(j, 0, r);
            else { if (!r.grupoDe && t.grupoDe) r.grupoDe = t.grupoDe; res.splice(j, 0, r); }
        } else {
            const jj = j - 1;
            if (r.grupoDe && t.grupoDe !== r.grupoDe) { delete r.grupoDe; res.splice(jj, 0, r); }                         // sai do grupo pela borda de baixo e fica onde está
            else if (!r.grupoDe && eDecl(t) && !t.grupoRecolhido) { r.grupoDe = t.gid; res.splice(jj + 1, 0, r); }
            else if (!r.grupoDe && eDecl(t)) { const f = bloco(res, jj).fim; res.splice(f + 1, 0, r); }
            else res.splice(jj + 1, 0, r);
        }
        return { lista: normalizar(res), foco: r };
    }

    /** Ao copiar linhas de outra obra para esta: os grupos ganham ids novos (a mesma obra pode entrar duas vezes sem misturar os grupos). Devolve as mesmas linhas, já com os ids trocados. */
    function reidentificar(linhas) {
        const novo = new Map();
        linhas.forEach(r => { if (eDecl(r)) { const g = novoId(); novo.set(r.gid, g); r.gid = g; } });
        linhas.forEach(r => { if (r.grupoDe) { if (novo.has(r.grupoDe)) r.grupoDe = novo.get(r.grupoDe); else delete r.grupoDe; } });
        return linhas;
    }
    /** Número (0, 1, 2…) da declaração entre as declarações da tabela: dá a cor do grupo. */
    function ordemDoGrupo(lista, gid) { return grupos(lista).findIndex(r => r.gid === gid); }
    /** Quantas linhas de dados (sem comentário nem linha em branco) o grupo tem. */
    function contar(lista, gid) { return membrosDe(lista, gid).filter(r => String(r.ativo || '').trim() !== '' && !r.comentario).length; }

    /* ── Pontos do mapa (TASK-225): cada ponto é um grupo «@nome» + linha de comentário «#» + as linhas do ativo ── */
    const TEXTO_COMENTARIO_VAZIO = '#';
    /** Próximo nome automático «P7»: maior número entre os grupos «P<n>» da tabela + 1. */
    function proximoNome(lista, prefixo = 'P', piso = 0) {
        const re = new RegExp('^' + prefixo + '(\\d+)$', 'i');
        let max = piso | 0;
        for (const r of lista) if (eDecl(r)) { const m = re.exec(nomeDe(r.ativo)); if (m) max = Math.max(max, Number(m[1])); }
        return prefixo + (max + 1);
    }
    /** Texto do nome informado → título do grupo («@» + nome em maiúsculas, sem «@»/«#» soltos). Vazio → null (nome automático). */
    function nomeLimpo(texto) {
        const t = String(texto == null ? '' : texto).replace(/^[\s@#]+/, '').replace(/\s+/g, ' ').trim().toUpperCase();
        return t ? t.slice(0, 60) : null;
    }
    /** As duas linhas que abrem um ponto: o título do grupo e a linha de comentário (já membro). `operacao` = a do ativo (a herança de operação não muda). */
    function abrirPonto(nome, operacao, comentario) {
        const gid = novoId(), op = operacao || 'I';
        const texto = String(comentario == null ? '' : comentario).trim().replace(/^#+\s*/, '');
        return {
            gid,
            titulo: { entidade: '0', operacao: op, ativo: '@' + nome, grupoDecl: true, congelada: true, gid },
            comentario: { entidade: '0', operacao: op, ativo: texto ? '# ' + texto : TEXTO_COMENTARIO_VAZIO, comentario: true, congelada: true, grupoDe: gid }
        };
    }
    /** Dados do ponto a partir de uma linha da tabela (índice `i`): título, comentário, membros — ou null se a linha não está em grupo. */
    function pontoDaLinha(lista, i) {
        const b = bloco(lista, i); if (!b) return null;
        const titulo = lista[b.h], membros = []; let comentario = null;
        for (let k = b.h + 1; k <= b.fim; k++) { if (lista[k].comentario && !comentario) comentario = k; else membros.push(k); }
        return { gid: b.gid, h: b.h, fim: b.fim, nome: rotuloDe(titulo), comentarioIdx: comentario,
            comentario: comentario === null ? '' : String(lista[comentario].ativo || '').replace(/^\s*#\s?/, '').trim(), membros };
    }
    /** Agrupa vários pontos de uma vez: `pontos` = [{linhas: [objetos de linha], nome?}]. Cada um vira «@nome», «#» e as suas linhas (nome automático quando falta). Lista nova + quantos grupos. */
    function agruparPontos(lista, pontos, piso = 0) {
        let atual = lista.slice(), n = 0;
        for (const p of pontos) {
            const linhas = p.linhas.filter(r => r && !eDecl(r) && !r.grupoDe);
            if (!linhas.length) continue;
            const idxs = linhas.map(r => atual.indexOf(r)).filter(i => i >= 0);
            if (!idxs.length) continue;
            const existentes = new Set(atual.filter(eDecl).map(r => rotuloDe(r).toUpperCase()));
            let nome = nomeLimpo(p.nome);
            if (!nome || existentes.has(nome)) nome = proximoNome(atual, 'P', piso);                      // nome pedido já existe: número automático
            const abertura = abrirPonto(nome, linhas[0].operacao, p.comentario);
            const r = criarGrupo(atual, idxs, nome, () => abertura.titulo);
            if (!r) continue;
            abertura.comentario.grupoDe = abertura.titulo.gid;                                                   // criarGrupo dá um id novo ao título
            r.lista.splice(r.lista.indexOf(abertura.titulo) + 1, 0, abertura.comentario);
            atual = r.lista; n++;
        }
        return { lista: atual, n };
    }

    const api = { proximoNome, nomeLimpo, abrirPonto, pontoDaLinha, agruparPontos, ehDeclaracao, novoId, nomeDe, rotuloDe, sincronizar, liberar, bloco, membrosDe, grupos, temGrupos, normalizar, mapa, novaLinhaNoGrupo, novaLinhaSolta,
        adicionarAoGrupo, criarGrupo, tirarDoGrupo, desfazerGrupo, mover, reidentificar, ordemDoGrupo, contar, selecionadas };
    if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.GruposLinhas = api;
})(typeof window !== 'undefined' ? window : globalThis);
