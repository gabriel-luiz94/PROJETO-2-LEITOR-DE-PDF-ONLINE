/*
 * static/regras_leitor_engine.js — Motor de regras do leitor (TASK-006).
 *
 * Duas tabelas encadeadas, avaliadas em ordem de prioridade (`ordem` crescente dentro de cada
 * `fase`):
 *   - Processamento: texto bruto, cor, layer (+ vizinhança) -> operação, ativo
 *   - Classificação: operação, ativo, cor, layer, texto bruto -> operação (ajustada), entidade
 *
 * A tabela de Processamento é dividida em 3 fases (campo `fase` de cada linha):
 *   1 = seleção de ramo (equivalente ao processAtivoFormula original): primeira regra que casar
 *       decide o ativo inicial; as demais da fase 1 são ignoradas.
 *   2 = pós-processamento: todas as regras que casarem são aplicadas em sequência sobre o ativo
 *       já decidido pela fase 1 (substituições incondicionais) — exceto SUBSTITUIR_TOTAL, que
 *       encerra a fase 2 imediatamente ao casar (equivalente ao "return" antecipado do código
 *       original para o caso "combo").
 *   3 = overrides por cor/layer (elo fusível, FLY, BLOCO+TERRA3, reinstalação, 01_ORCAMENTO):
 *       primeira regra que casar E tiver `parar=true` decide e encerra a fase.
 *
 * Modos de regra de Processamento:
 *   - "DEFINIR"          : casa contra o TEXTO BRUTO; se casar, define o ativo a partir de um
 *                           template (placeholders {1}, {2}... = grupos capturados;
 *                           {TEXTO_LIMPO} = texto bruto limpo — hifens normalizados, espaços
 *                           colapsados, sem mudar maiúsc./minúsc.)
 *   - "SUBSTITUIR"        : casa contra o ATIVO já calculado e aplica um `String.replace` (sintaxe
 *                           nativa do JS no template: $1, $2...)
 *   - "SUBSTITUIR_TOTAL"  : casa contra o ATIVO já calculado e SUBSTITUI o ativo inteiro pelo
 *                           template (mesmos placeholders de DEFINIR).
 *
 * Funciona tanto no navegador (window.RegrasLeitorEngine) quanto em Node (module.exports), para
 * permitir testar contra o oráculo em scratchpad/regras_test.
 */
(function (root) {
    function _isGray(hex) {
        if (!hex || hex.length < 7) return false;
        const r = parseInt(hex.substring(1, 3), 16);
        const g = parseInt(hex.substring(3, 5), 16);
        const b = parseInt(hex.substring(5, 7), 16);
        return Math.abs(r - g) < 5 && Math.abs(g - b) < 5 && r > 20 && r < 230;
    }

    function _isBlue(hex) {
        if (!hex || hex.length < 7) return false;
        const r = parseInt(hex.substring(1, 3), 16);
        const g = parseInt(hex.substring(3, 5), 16);
        const b = parseInt(hex.substring(5, 7), 16);
        return b > 80 && b > r * 1.5 && b > g * 1.5;
    }

    function classificarCor(hex) {
        const h = (hex || '#000000').toUpperCase();
        if (h === '#FF0000') return 'VERMELHO';
        if (_isGray(h)) return 'CINZA';
        if (_isBlue(h)) return 'AZUL';
        return 'OUTRA';
    }

    function corCasa(corEm, hex, corNaoEm) {
        const c = classificarCor(hex);
        if (corNaoEm && corNaoEm.indexOf(c) !== -1) return false;
        if (!corEm || corEm.length === 0) return true;
        return corEm.some(function (regraCor) {
            if (regraCor === 'PRETO') return c !== 'VERMELHO' && c !== 'CINZA';
            return c === regraCor;
        });
    }

    function layerCasa(layerEm, layerReal) {
        if (!layerEm || layerEm.length === 0) return true;
        const lr = (layerReal || '').trim().toUpperCase();
        return layerEm.map(function (l) { return l.toUpperCase(); }).includes(lr);
    }

    var HIFENS_UNICODE = new RegExp('[­‐-―−]', 'g');
    var NBSP = new RegExp(' ', 'g');

    function limparTexto(texto) {
        // Espelha script.js: normaliza hifens unicode e colapsa espaços, sem mudar caixa.
        const normalized = (texto || '').replace(HIFENS_UNICODE, '-').replace(NBSP, ' ');
        return normalized.replace(/\s*-\s*/g, '-').trim().replace(/\s+/g, ' ');
    }

    function operacaoPelaCor(hex) {
        const c = classificarCor(hex);
        if (c === 'VERMELHO') return 'I';
        if (c === 'CINZA') return 'R';
        return 'M';
    }

    function _safeRegex(source, flags) {
        try { return new RegExp(source, flags); } catch (e) { return null; }
    }

    function _aplicarTemplateDefinir(template, match, textoLimpo, vizinhancaValor, textoBruto) {
        if (template === null || template === undefined) return null;
        let out = template;
        if (match) {
            for (let i = 1; i < match.length; i++) {
                out = out.split('{' + i + '}').join(match[i] !== undefined ? match[i] : '');
            }
        }
        out = out.split('{TEXTO_LIMPO}').join(textoLimpo || '');
        out = out.split('{vizinhanca}').join(vizinhancaValor !== null && vizinhancaValor !== undefined ? vizinhancaValor : '');
        if (out.indexOf('{TEXTO_SEM_ULTIMOS5}') !== -1) {
            // Espelha script.js: arrumeRaw = col2.trim().replace(/\s+/g,' '); substring(0, len-5)
            const arrumeRaw = (textoBruto || '').trim().replace(/\s+/g, ' ');
            const semUltimos5 = arrumeRaw.slice(0, Math.max(0, arrumeRaw.length - 5));
            out = out.split('{TEXTO_SEM_ULTIMOS5}').join(semUltimos5);
        }
        return out;
    }

    // Aplica uma lista (já ordenada) de regras de Processamento sobre o estado {operacao, ativo}.
    function _rodarFase(item, regrasFase, estado, textoLimpo) {
        for (const r of regrasFase) {
            if (!corCasa(r.cor_em, item.cor, r.cor_nao_em)) continue;
            if (!layerCasa(r.layer_em, item.layer)) continue;

            const modo = r.modo || 'DEFINIR';
            let casou = false;
            let ativoResultante = null;

            if (modo === 'DEFINIR') {
                let match = null;
                if (r.texto_regex) {
                    const re = _safeRegex(r.texto_regex, 'i');
                    if (!re) continue;
                    match = item.texto ? item.texto.match(re) : null;
                    if (!match) continue;
                }
                casou = true;

                let vizValor = null;
                if (r.vizinhanca && r.vizinhanca.regex) {
                    const janela = r.vizinhanca.janela || 10;
                    const start = Math.max(0, item.index - janela);
                    const end = Math.min((item.allItems || []).length - 1, item.index + janela);
                    const re = _safeRegex(r.vizinhanca.regex, 'i');
                    if (re) {
                        for (let i = start; i <= end; i++) {
                            if (i === item.index) continue;
                            const viz = item.allItems[i];
                            if (!viz || !viz.texto) continue;
                            if (r.vizinhanca.mesma_pagina !== false && viz.pagina !== item.pagina) continue;
                            const m = viz.texto.match(re);
                            if (m) { vizValor = m[1] !== undefined ? m[1] : m[0]; break; }
                        }
                    }
                    if (vizValor === null) {
                        ativoResultante = null; // vizinhança exigida mas não encontrada: não altera
                    } else {
                        ativoResultante = _aplicarTemplateDefinir(r.ativo_template, match, textoLimpo, vizValor, item.texto);
                    }
                } else {
                    ativoResultante = _aplicarTemplateDefinir(r.ativo_template, match, textoLimpo, null, item.texto);
                }
            } else if (modo === 'SUBSTITUIR' || modo === 'SUBSTITUIR_TOTAL') {
                if (!r.ativo_regex) continue;
                const re = _safeRegex(r.ativo_regex, 'i');
                if (!re) continue;
                const match = estado.ativo.match(re);
                if (!match) continue;
                casou = true;
                if (modo === 'SUBSTITUIR_TOTAL') {
                    ativoResultante = _aplicarTemplateDefinir(r.ativo_template, match, textoLimpo, null, item.texto);
                } else {
                    ativoResultante = estado.ativo.replace(re, r.ativo_template || '');
                }
            }

            if (!casou) continue;

            if (r.operacao === 'PELA_COR') {
                estado.operacao = operacaoPelaCor(item.cor);
            } else if (r.operacao === '0') {
                estado.operacao = '0';
            } else if (r.operacao !== undefined && r.operacao !== null && r.operacao !== '') {
                estado.operacao = r.operacao;
            }
            // r.operacao === '' ou ausente: não altera a operação já decidida.

            if (ativoResultante !== null) estado.ativo = ativoResultante;

            if (r.parar || modo === 'SUBSTITUIR_TOTAL') break;
        }
    }

    // item: {texto, cor, layer, pagina, index, allItems}
    function processar(item, regras) {
        const textoLimpo = limparTexto(item.texto);
        const estado = { operacao: null, ativo: '' };

        const todas = (regras || []).filter(function (r) { return r.ativa !== false; });
        [1, 2, 3].forEach(function (fase) {
            const doFase = todas
                .filter(function (r) { return (r.fase || 3) === fase; })
                .sort(function (a, b) { return (a.ordem || 0) - (b.ordem || 0); });
            _rodarFase(item, doFase, estado, textoLimpo);
        });

        let operacao = estado.operacao;
        const ativo = estado.ativo;

        // Baseline: se nenhuma regra definiu operação, cai para PELA_COR (igual ao original).
        if (operacao === null) operacao = operacaoPelaCor(item.cor);

        // Passos fixos de motor (não são linhas de tabela — ver TASK-006.md "Regras de motor"):
        const layerUpper = (item.layer || '').trim().toUpperCase();
        if (layerUpper === '01_LV' && operacao && operacao.charAt(0) !== '*') {
            operacao = '*' + operacao;
        }
        const blueExemptLayers = ['01_LV', '01_RETENS', '01_RETENS_LV'];
        if (_isBlue(item.cor) && blueExemptLayers.indexOf(layerUpper) === -1) {
            operacao = '0';
        }
        if (layerUpper === '01_PARTICULAR') {
            operacao = '0';
        }

        return { operacao: operacao, ativo: (ativo || '').trim() };
    }

    function classificar(operacao, ativo, cor, layer, texto, regras) {
        let entidade = '0';
        let operacaoFinal = operacao;

        const ordenadas = (regras || [])
            .filter(function (r) { return r.ativa !== false; })
            .slice()
            .sort(function (a, b) { return (a.ordem || 0) - (b.ordem || 0); });

        for (const r of ordenadas) {
            if (r.operacao_em && r.operacao_em.length > 0 && r.operacao_em.indexOf(operacaoFinal) === -1) continue;
            if (!corCasa(r.cor_em, cor, r.cor_nao_em)) continue;
            if (!layerCasa(r.layer_em, layer)) continue;
            if (r.ativo_regex) {
                const re = _safeRegex(r.ativo_regex, 'i');
                if (!re || !re.test(ativo || '')) continue;
            }
            if (r.texto_regex) {
                const re = _safeRegex(r.texto_regex, 'i');
                if (!re || !re.test(texto || '')) continue;
            }

            if (r.entidade) entidade = r.entidade;
            if (r.operacao_ajustada) operacaoFinal = r.operacao_ajustada;

            if (r.parar) break;
        }

        return { operacao: operacaoFinal, entidade: entidade };
    }

    // Combina as duas tabelas e aplica as regras de motor operação=0 (TASK-006.md).
    function processarEClassificar(item, regrasProcessamento, regrasClassificacao) {
        const p = processar(item, regrasProcessamento);
        if (p.operacao === '0') {
            return { operacao: '0', ativo: p.ativo, entidade: '0' };
        }
        const cls = classificar(p.operacao, p.ativo, item.cor, item.layer, item.texto, regrasClassificacao);
        return { operacao: cls.operacao, ativo: p.ativo, entidade: cls.entidade };
    }

    const api = {
        classificarCor: classificarCor,
        corCasa: corCasa,
        layerCasa: layerCasa,
        limparTexto: limparTexto,
        operacaoPelaCor: operacaoPelaCor,
        processar: processar,
        classificar: classificar,
        processarEClassificar: processarEClassificar
    };

    if (typeof module !== 'undefined' && module.exports) {
        module.exports = api;
    }
    if (typeof root !== 'undefined') {
        root.RegrasLeitorEngine = api;
    }
})(typeof window !== 'undefined' ? window : undefined);
