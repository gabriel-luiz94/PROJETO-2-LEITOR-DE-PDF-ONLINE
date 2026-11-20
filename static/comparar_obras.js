/* TASK-211: comparação do preenchimento de Cabos e Outros no Resumo. */
(function (raiz) {
    'use strict';

    const texto = valor => String(valor ?? '').trim().replace(/\s+/g, ' ');
    const normal = valor => texto(valor).toUpperCase();
    const linhas = (obra, tipo) => {
        const tabela = obra?.[tipo];
        return Array.isArray(tabela) ? tabela : Array.isArray(tabela?.data) ? tabela.data : [];
    };
    const campo = (linha, chave) => linha[chave] ?? linha._raw?.[chave];

    function ponto(valor) {
        if (!Array.isArray(valor) || valor.length < 2 || !valor.slice(0, 2).every(Number.isFinite)) return '';
        return valor.slice(0, 2).map(n => Number(n.toFixed(2))).join(', ');
    }

    function placa(linha) {
        const direta = normal(linha.placa || linha._placa);
        if (direta) return direta;
        const placaNoTexto = String(linha.textoOriginal || linha._raw?.texto || '').split(/[\r\n]+/)
            .map(normal).filter(t => /^[A-Z]{1,2}\d{5,7}$/.test(t));
        return placaNoTexto.length === 1 ? placaNoTexto[0] : '';
    }

    const juntar = itens => itens.filter(Boolean).join(' · ');
    const chave = itens => JSON.stringify(itens);

    function preparar(obra, opcoes, placasComuns) {
        const outros = linhas(obra, 'outros');
        const locais = outros.map(linha => linha ? { placa: placa(linha), pos: ponto(campo(linha, '_poste')) } : {});
        const placasPorPonto = new Map();
        locais.forEach(l => {
            if (!l.pos || !placasComuns.has(l.placa)) return;
            if (placasPorPonto.has(l.pos) && placasPorPonto.get(l.pos) !== l.placa) placasPorPonto.set(l.pos, '');
            else placasPorPonto.set(l.pos, l.placa);
        });
        const etiquetaPonto = pos => placasPorPonto.get(pos) || pos;

        function vinculo(id, estrutura) {
            const indice = Math.floor(id + 1e-9);
            if (!Number.isInteger(indice) || !outros[indice]) return 'Referência ausente (' + id + ')';
            const linha = outros[indice], loc = locais[indice];
            const destino = loc.placa && placasComuns.has(loc.placa) ? loc.placa :
                loc.pos ? etiquetaPonto(loc.pos) : loc.placa || normal(linha.ativo);
            if (!estrutura) return destino;
            const slot = Math.round((id - indice) * 100);
            const codigos = opcoes.estruturas?.(linha) || [];
            const codigo = normal(codigos[slot]) || (slot ? 'Estrutura ' + (slot + 1) : normal(linha.ativo));
            return juntar([destino, codigo]);
        }

        function registros(tipo) {
            return linhas(obra, tipo).flatMap((linha, i) => {
                if (!linha || linha.comentario || linha.grupoDecl || !normal(linha.ativo)) return [];
                const cabos = tipo === 'cabos';
                const geo = campo(linha, '_geo');
                const extremos = cabos && Array.isArray(geo) && geo.length >= 2
                    ? [ponto(geo[0]), ponto(geo[geo.length - 1])] : [];
                const campoLista = k => Array.isArray(linha[k]) ? linha[k] : [];
                const atribuicoes = cabos ? campoLista('vinculoPostes').map(id => vinculo(id, false)).sort() : [];
                const pontas = extremos.length && extremos.every(Boolean) ? extremos.map(etiquetaPonto).sort()
                    : atribuicoes.length === 2 && campoLista('vinculoPostes').every(id => locais[id]?.pos
                        || placasComuns.has(locais[id]?.placa)) ? atribuicoes : [];
                const local = cabos ? { pos: pontas.length ? chave(pontas) : '', placa: '' } : locais[i];
                const info = cabos ? opcoes.cabo?.(linha.ativo) : null;
                const valores = {
                    'Ativo': normal(linha.ativo),
                    'Operação': normal(linha.operacao || 'M'),
                    'Entidade': normal(linha.entidade || '0')
                };
                if (cabos) {
                    const qtd = linha.qtdAtivos;
                    Object.assign(valores, {
                        'Metragem': Number.isFinite(info?.comprimento) ? info.comprimento + ' m' : '',
                        'Quantidade': qtd == null || qtd === '' ? '' : String(qtd),
                        'Fases': texto(linha.fases), 'Tensão': normal(linha.tensao), 'Nível': normal(linha.nivel || campo(linha, '_geoTipo')),
                        'Vínculos': campoLista('vinculoEstruturas').map(id => vinculo(id, true)).sort().join(' | '),
                        'Postes atribuídos': atribuicoes.join(' | '),
                        'Estruturas fictícias': campoLista('estruturasFicticias').map(normal).sort().join(' | ')
                    });
                } else {
                    valores.Placa = local.placa || '';
                    valores['Cabos fictícios'] = normal(campoLista('cabosFicticios').join(', '));
                    valores['Cabos fictícios por estrutura'] = Object.entries(linha.cabosFicticiosSlots || {})
                        .filter(([, lista]) => Array.isArray(lista) && lista.length)
                        .map(([slot, lista]) => slot + ': ' + lista.map(normal).join(', ')).join(' | ');
                }
                valores.Estado = linha.congelada ? 'Congelada' : '';
                return [{
                    i, local, valores, metros: info?.comprimento, assinatura: chave(Object.values(valores)),
                    rotulo: cabos ? pontas.join(' ↔ ') : juntar([local.placa, local.pos]),
                    codigo: cabos ? normal(info?.prefixo || linha.ativo) : normal(opcoes.codigoOutro?.(linha) || linha.ativo)
                }];
            });
        }
        return { cabos: registros('cabos'), outros: registros('outros') };
    }

    function agrupar(lista, por) {
        const mapa = new Map();
        lista.forEach(reg => {
            const id = por(reg);
            if (!id) return;
            if (!mapa.has(id)) mapa.set(id, []);
            mapa.get(id).push(reg);
        });
        return mapa;
    }

    function compararTabela(referencia, comparada) {
        const usados = new Set(), resultado = [];
        const pendentes = lista => lista.filter(r => !usados.has(r));
        function emitir(a, b, conferir = false) {
            if (a) usados.add(a);
            if (b) usados.add(b);
            const campos = a && b ? Object.keys(a.valores).filter(c => a.valores[c] !== b.valores[c]) : [];
            const estado = conferir ? 'conferir' : !a ? 'comparada' : !b ? 'referencia' : campos.length ? 'diferente' : 'igual';
            const deltaMetros = Number.isFinite(a?.metros) && Number.isFinite(b?.metros) ? b.metros - a.metros : null;
            resultado.push({ referencia: a, comparada: b, estado, campos, deltaMetros });
        }
        function parearUnicos(a, b, por, aceitar = () => true, conferir = false) {
            const mb = agrupar(b, por);
            agrupar(a, por).forEach((as, id) => {
                const bs = mb.get(id);
                if (as.length === 1 && bs?.length === 1 && aceitar(as[0], bs[0])) emitir(as[0], bs[0], conferir);
            });
        }
        parearUnicos(referencia, comparada, r => r.local.placa);
        const porPosB = agrupar(pendentes(comparada), r => r.local.pos);
        agrupar(pendentes(referencia), r => r.local.pos).forEach((as, pos) => {
            const bs = porPosB.get(pos);
            if (!bs) return;
            // Linhas idênticas são contadas individualmente, inclusive em vãos com cabos sobrepostos.
            const iguais = agrupar(bs, r => r.assinatura);
            as.forEach(a => { const b = iguais.get(a.assinatura)?.pop(); if (b) emitir(a, b); });
            parearUnicos(pendentes(as), pendentes(bs), r => r.codigo);
            const restantesA = pendentes(as), restantesB = pendentes(bs);
            if (restantesA.length === 1 && restantesB.length === 1) emitir(restantesA[0], restantesB[0]);
            else if (restantesA.length && restantesB.length) {
                restantesA.forEach(a => emitir(a, null, true));
                restantesB.forEach(b => emitir(null, b, true));
            }
        });
        const localizado = r => r.local.pos || r.local.placa;
        const semLocalB = agrupar(pendentes(comparada).filter(r => !localizado(r)), r => r.assinatura);
        pendentes(referencia).filter(r => !localizado(r)).forEach(a => {
            const b = semLocalB.get(a.assinatura)?.pop();
            if (b) emitir(a, b);
        });
        const candidato = (a, b) => !localizado(a) || !localizado(b);
        parearUnicos(pendentes(referencia), pendentes(comparada), r => r.assinatura, candidato, true);
        parearUnicos(pendentes(referencia), pendentes(comparada), r => r.codigo, candidato, true);
        const as = pendentes(referencia), bs = pendentes(comparada);
        const codigos = lista => ({
            todos: new Set(lista.map(r => r.codigo)),
            semLocal: new Set(lista.filter(r => !localizado(r)).map(r => r.codigo))
        });
        const ca = codigos(as), cb = codigos(bs);
        const incerto = (r, c) => (localizado(r) ? c.semLocal : c.todos).has(r.codigo);
        as.forEach(a => emitir(a, null, incerto(a, cb)));
        bs.forEach(b => emitir(null, b, incerto(b, ca)));
        return resultado.sort((a, b) => (a.referencia?.i ?? Infinity) - (b.referencia?.i ?? Infinity)
            || (a.comparada?.i ?? Infinity) - (b.comparada?.i ?? Infinity));
    }

    function comparar(referencia, comparada, opcoes = {}) {
        const obterPlacas = obra => linhas(obra, 'outros').filter(Boolean).map(placa).filter(Boolean);
        const somenteUnicas = lista => {
            const unicas = new Set(), repetidas = new Set();
            lista.forEach(id => { if (unicas.has(id)) repetidas.add(id); else unicas.add(id); });
            repetidas.forEach(id => unicas.delete(id));
            return unicas;
        };
        const aPlacas = somenteUnicas(obterPlacas(referencia)), bPlacas = somenteUnicas(obterPlacas(comparada));
        const comuns = new Set([...aPlacas].filter(id => bPlacas.has(id)));
        const a = preparar(referencia, opcoes, comuns), b = preparar(comparada, opcoes, comuns);
        return { cabos: compararTabela(a.cabos, b.cabos), outros: compararTabela(a.outros, b.outros) };
    }

    const rotulos = {
        igual: 'Igual', diferente: 'Alterado', referencia: 'Só na obra aberta',
        comparada: 'Só na obra comparada', conferir: 'Conferir correspondência'
    };

    function abrir(resultado, opcoes) {
        const dialogo = document.getElementById('modal-comparar-obras');
        const corpo = dialogo.querySelector('tbody'), filtro = dialogo.querySelector('select');
        const botoes = Array.from(dialogo.querySelectorAll('[data-tabela]'));
        const resumo = dialogo.querySelector('[role="status"]');
        const anterior = dialogo.querySelector('[data-pagina="anterior"]'), proxima = dialogo.querySelector('[data-pagina="proxima"]');
        const fechamento = dialogo.querySelector('[data-fechar]');
        const paginaTam = 100;
        const formatarDelta = new Intl.NumberFormat('pt-BR', { signDisplay: 'always', maximumFractionDigits: 4 });
        let tipo = 'outros', pagina = 0;
        const criar = (tag, txt, pai) => {
            const el = document.createElement(tag);
            if (txt != null) el.textContent = txt;
            if (pai) pai.appendChild(el);
            return el;
        };
        dialogo.querySelector('[data-nome]').textContent = opcoes.nome;
        dialogo.querySelector('[data-referencia]').textContent = opcoes.referencia || '';
        filtro.value = 'divergencias';

        function coluna(tr, registro, atual, campos) {
            const td = criar('td', null, tr);
            if (!registro) { td.textContent = '—'; return; }
            const v = registro.valores;
            criar('strong', v.Ativo, td);
            criar('small', juntar([v.Entidade, 'Operação ' + v.Operação, v.Metragem, v.Quantidade && 'Qtd. ' + v.Quantidade]), td);
            Object.entries(v).filter(([k, valor]) => !['Ativo', 'Operação', 'Entidade', 'Metragem', 'Quantidade'].includes(k)
                && (valor || campos.includes(k))).forEach(([k, valor]) => criar('small', k + ': ' + (valor || '—'), td));
            if (atual) {
                const bt = criar('button', 'Ir à linha ' + (registro.i + 1), td);
                bt.className = 'btn-secondary';
                bt.onclick = () => { dialogo.close(); opcoes.localizar(tipo, registro.i); };
            } else criar('small', 'Linha ' + (registro.i + 1), td);
        }

        function render() {
            const todos = resultado[tipo];
            const filtros = filtro.value;
            const visiveis = todos.filter(item => filtros === 'todos' || (filtros === 'divergencias' ? item.estado !== 'igual' : item.estado === filtros));
            botoes.forEach(b => {
                b.setAttribute('aria-pressed', String(b.dataset.tabela === tipo));
                b.querySelector('small').textContent = resultado[b.dataset.tabela].filter(r => r.estado !== 'igual').length;
            });
            const inicio = pagina * paginaTam;
            resumo.textContent = visiveis.length
                ? (inicio + 1) + '–' + Math.min(inicio + paginaTam, visiveis.length) + ' de ' + visiveis.length + ' registros'
                : filtros === 'divergencias' && todos.length ? 'Nenhuma divergência nesta tabela.' : 'Nenhum registro neste filtro.';
            corpo.replaceChildren();
            visiveis.slice(inicio, inicio + paginaTam).forEach(item => {
                const tr = criar('tr', null, corpo);
                const r = item.referencia, b = item.comparada;
                const locais = [...new Set([r?.rotulo, b?.rotulo].filter(Boolean))];
                criar('td', locais.join(' → ') || 'Sem localização', tr);
                coluna(tr, r, true, item.campos);
                coluna(tr, b, false, item.campos);
                const td = criar('td', null, tr);
                const status = criar('strong', rotulos[item.estado], td);
                status.dataset.estado = item.estado;
                if (item.campos.length) criar('small', item.campos.join(', '), td);
                if (item.deltaMetros) criar('small', formatarDelta.format(item.deltaMetros) + ' m', td);
            });
            anterior.disabled = !pagina;
            proxima.disabled = inicio + paginaTam >= visiveis.length;
        }
        botoes.forEach(bt => bt.onclick = () => { tipo = bt.dataset.tabela; pagina = 0; render(); });
        filtro.onchange = () => { pagina = 0; render(); };
        anterior.onclick = () => { pagina--; render(); };
        proxima.onclick = () => { pagina++; render(); };
        fechamento.onclick = () => dialogo.close();
        dialogo.onclose = () => {
            corpo.replaceChildren();
            dialogo.onclose = filtro.onchange = anterior.onclick = proxima.onclick = fechamento.onclick = null;
            botoes.forEach(bt => bt.onclick = null);
        };
        render();
        if (!dialogo.open) dialogo.showModal();
    }

    const api = { comparar, abrir };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else raiz.CompararObras = api;
})(typeof window !== 'undefined' ? window : globalThis);
