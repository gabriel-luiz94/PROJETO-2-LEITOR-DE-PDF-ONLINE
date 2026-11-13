/*
 * projeto.js (TASK-154) — PROJETO DE TRABALHO: uma pasta com um arquivo `.obra` que recebe tudo o que é feito enquanto se trabalha.
 *   • Menu "Projeto": Novo projeto… / Abrir projeto… / Propriedades… / Voltar ao projeto default.
 *   • Projeto default = o funcionamento de sempre (nada é gravado em arquivo). Só existe no aplicativo desktop (a ponte `pywebview.api.projeto_*`);
 *     na versão online/navegador este módulo não faz nada e o menu não aparece.
 *   • Com um projeto aberto, a gravação é AUTOMÁTICA e em segundo plano: depois de uma pausa de 2 s sem mexer, ao trocar de tela
 *     (Resultado, Tabela de orçamento, Totalizadora), ao voltar ao default e ao fechar. Só grava o que mudou e nunca trava a tela.
 *   • Resumo (Cabos/Outros/Totalizadora) é gravado pela página do Resumo; o resultado do orçamento e as informações do relatório SAP, pelas suas telas.
 *   • Se o arquivo for alterado por fora (outro computador), avisa antes de sobrescrever.
 * O estado do projeto atual fica em localStorage (`gridr_projeto`), então Resumo, Resultado (moldura) e Tabela de orçamento enxergam o mesmo projeto.
 * Funções puras expostas em `GridRProjeto` e testadas em Node: tests/test_projeto_153.py.
 */
(function (raiz) {
    'use strict';

    const CHAVE = 'gridr_projeto';
    const ESPERA_MS = 2000;
    const CAMPOS_SAP = [['sap', 'Nº SAP'], ['bairro', 'Bairro'], ['municipio', 'Município'], ['utd', 'UTD'], ['utep', 'UTEP'], ['projetista', 'Projetista'],
        ['data', 'Data'], ['id_investimento', 'ID Investimento (ODI)'], ['id_investimento_odd', 'ID Investimento (ODD) — vazio = igual ao ODI']];

    /* ── Parte pura ── */
    /** Hash curto (FNV-1a 32 bits + tamanho) de um texto: identifica o cálculo do orçamento que gerou as linhas do Resultado. */
    function hashTexto(t) {
        t = String(t == null ? '' : t);
        let h = 0x811c9dc5;
        for (let i = 0; i < t.length; i++) { h ^= t.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
        return h.toString(16) + '-' + t.length.toString(16);
    }
    /** Pasta de um caminho do Windows ou Linux ("D:\\Obras\\A\\x.obra" → "D:\\Obras\\A"). */
    function pastaDoCaminho(c) {
        c = String(c || '');
        const i = Math.max(c.lastIndexOf('\\'), c.lastIndexOf('/'));
        return i > 0 ? c.slice(0, i) : '';
    }
    /** O Resumo "já está preenchido"? (alguma linha com ativo/texto em Cabos, Outros ou Totalizadora). */
    function temDados(dados) {
        const lista = t => (Array.isArray(t) ? t : (t && Array.isArray(t.data) ? t.data : []));
        return ['cabos', 'outros', 'totalizadora'].some(k => lista(dados && dados[k]).some(r => r && (String(r.ativo || '').trim() || String(r.texto || '').trim())));
    }
    /** Lê o JSON copiado do Atlas ("atlas-projeto-rede", versão 1). Devolve {ok, campos} (só campos preenchidos; vazio nunca apaga) ou {ok: false, erro}.
     *  Mapa: titulo, nota, sap, bairro, municipio, utd, utep, projetista, data → mesmos nomes; odiCodigo → id_investimento; oddCodigo → id_investimento_odd. */
    function interpretarSap(texto) {
        let t = String(texto == null ? '' : texto).replace(/^\uFEFF/, '').trim();
        const cerca = t.match(/^```[a-z]*\s*([\s\S]*?)\s*```$/i);              // veio dentro de um bloco de código markdown
        if (cerca) t = cerca[1].trim();
        if (!t) return { ok: false, erro: 'Não há nada copiado.' };
        let o;
        try { o = JSON.parse(t); } catch (e) { return { ok: false, erro: 'O texto copiado não é um JSON do Atlas.' }; }
        if (!o || typeof o !== 'object' || Array.isArray(o)) return { ok: false, erro: 'O texto copiado não é um JSON do Atlas.' };
        if (o.source !== 'atlas-projeto-rede') return { ok: false, erro: 'O texto copiado não é do Atlas (campo "source" diferente de "atlas-projeto-rede").' };
        if (o.version !== 1) return { ok: false, erro: 'Versão ' + String(o.version).slice(0, 10) + ' do Atlas não é suportada (esperado 1). Atualize o programa.' };
        const limpo = v => (typeof v === 'string' || (typeof v === 'number' && isFinite(v)) ? String(v).replace(/[\u0000-\u001f\u007f]+/g, ' ').trim().slice(0, 120) : '');
        const campos = {};
        ['titulo', 'nota', 'sap', 'bairro', 'municipio', 'utd', 'utep', 'projetista', 'data'].forEach(k => { const v = limpo(o[k]); if (v) campos[k] = v; });
        if (campos.data) { const m = campos.data.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) campos.data = m[3] + '/' + m[2] + '/' + m[1]; }
        let idi = limpo(o.odiCodigo);
        if (!idi) { const m = limpo(o.odi).match(/^([0-9][0-9.]*)/); if (m) idi = m[1]; }           // sem o código: pega o número do começo de "8.2 Melhoramento"
        if (idi) campos.id_investimento = idi;
        let idd = limpo(o.oddCodigo);
        if (!idd) { const m = limpo(o.odd).match(/^([0-9][0-9.]*)/); if (m) idd = m[1]; }
        if (idd) campos.id_investimento_odd = idd;
        if (!Object.keys(campos).length) return { ok: false, erro: 'O JSON do Atlas não traz nenhum dado para preencher.' };
        return { ok: true, campos };
    }
    const api = { hashTexto, pastaDoCaminho, temDados, interpretarSap, CAMPOS_SAP, CHAVE };
    if (typeof document === 'undefined') { if (typeof module !== 'undefined' && module.exports) module.exports = api; return; }
    if (raiz.GridRProjeto) return;

    /* ── Estado compartilhado ── */
    const doc = raiz.document;
    function estado() { try { const e = JSON.parse(raiz.localStorage.getItem(CHAVE) || 'null'); return e && e.caminho ? e : null; } catch (x) { return null; } }
    function gravarEstado(e) { try { if (e) raiz.localStorage.setItem(CHAVE, JSON.stringify(e)); else raiz.localStorage.removeItem(CHAVE); } catch (x) { /* sem storage: vale só nesta tela */ } }
    function ponte() {
        for (const w of [raiz, raiz.parent, raiz.top]) {
            try { if (w && w.pywebview && w.pywebview.api && w.pywebview.api.projeto_criar) return w.pywebview.api; } catch (x) { /* outra origem */ }
        }
        return null;
    }
    const ativo = () => !!(ponte() && estado());
    const codigoRegras = () => { try { return raiz.localStorage.getItem('projeto_selecionado_codigo') || '229'; } catch (x) { return '229'; } };
    const avisar = m => { try { if (raiz.showToast) raiz.showToast(m); else if (raiz.parent && raiz.parent !== raiz && raiz.parent.showToast) raiz.parent.showToast(m); } catch (x) { /* sem toast */ } };
    /** Quem grava de fato é a instância da janela principal (uma fila só, mesmo com o Resultado numa moldura). */
    function dono() { try { if (raiz.top !== raiz && raiz.top.GridRProjeto && raiz.top.GridRProjeto.__enviar) return raiz.top.GridRProjeto; } catch (x) { /* outra origem */ } return null; }

    /* ── Interface mínima (diálogos) ── */
    function el(tag, estilo, texto, filhos) {
        const e = doc.createElement(tag);
        if (estilo) e.style.cssText = estilo;
        if (texto != null) e.textContent = texto;
        (filhos || []).forEach(f => e.appendChild(f));
        return e;
    }
    const ESTILO_BTN = 'padding:8px 14px;border-radius:6px;border:1px solid #30363d;background:#21262d;color:#e6edf3;font:inherit;font-size:13px;cursor:pointer;';
    const ESTILO_BTN_OK = 'padding:8px 14px;border-radius:6px;border:1px solid #2ea043;background:#238636;color:#fff;font:inherit;font-size:13px;cursor:pointer;';
    const ESTILO_CAMPO = 'width:100%;box-sizing:border-box;background:#0d1117;border:1px solid #30363d;border-radius:6px;color:#e6edf3;padding:8px;font:inherit;font-size:14px;';
    function botao(texto, ok, aoClicar) { const b = el('button', ok ? ESTILO_BTN_OK : ESTILO_BTN, texto); b.type = 'button'; b.addEventListener('click', aoClicar); return b; }
    function habilitar(b, sim) { b.disabled = !sim; b.style.opacity = sim ? '1' : '0.4'; b.style.cursor = sim ? 'pointer' : 'not-allowed'; }
    /** Janela modal: devolve {fundo, corpo, rodape, fechar}. Fica acima da moldura do Resultado (z-index maior que o de abrir_na_app.js). */
    function janela(titulo, largura) {
        const fundo = el('div', 'position:fixed;inset:0;z-index:200000;background:rgba(1,4,9,.7);display:flex;align-items:center;justify-content:center;font-family:system-ui,"Segoe UI",sans-serif;');
        const caixa = el('div', `width:${largura || 460}px;max-width:94vw;max-height:92vh;overflow:auto;background:#161b22;border:1px solid #30363d;border-radius:10px;box-shadow:0 8px 30px #000a;color:#e6edf3;`);
        const corpo = el('div', 'padding:16px;display:flex;flex-direction:column;gap:12px;');
        const rodape = el('div', 'padding:12px 16px;border-top:1px solid #30363d;display:flex;gap:8px;justify-content:flex-end;');
        caixa.append(el('h3', 'margin:0;padding:12px 16px;border-bottom:1px solid #30363d;font-size:15px;', titulo), corpo, rodape);
        fundo.setAttribute('data-gridr-projeto', '1');
        fundo.appendChild(caixa);
        doc.body.appendChild(fundo);
        const fechar = () => { fundo.remove(); doc.removeEventListener('keydown', tecla, true); };
        const tecla = e => { if (e.key === 'Escape') { e.stopPropagation(); fechar(); } };
        doc.addEventListener('keydown', tecla, true);
        return { fundo, corpo, rodape, fechar };
    }
    function campo(rotulo, valor, extra) {
        const i = el('input', ESTILO_CAMPO); i.type = 'text'; i.value = valor || ''; i.autocomplete = 'off';
        const bloco = el('div', '', null, [el('label', 'display:block;font-size:12px;color:#8b949e;margin-bottom:4px;', rotulo), i]);
        if (extra) bloco.appendChild(extra);
        return { bloco, input: i };
    }
    /** Pergunta com botões; resolve com o `valor` do botão escolhido (ou null ao cancelar/Esc). `opcoes` = [{texto, valor, ok, descricao}] */
    function perguntar(titulo, texto, opcoes) {
        return new Promise(resolve => {
            const j = janela(titulo, 460);
            let respondido = false;
            const responder = v => { if (respondido) return; respondido = true; j.fechar(); resolve(v); };
            j.corpo.appendChild(el('div', 'font-size:13px;color:#8b949e;', texto));
            opcoes.forEach(o => {
                const b = el('button', 'text-align:left;padding:10px;border:1px solid #30363d;border-radius:8px;background:' + (o.ok ? '#1f6feb22' : '#0d1117') + ';color:#e6edf3;font:inherit;cursor:pointer;', null, [el('b', 'display:block;font-size:14px;', o.texto)]);
                if (o.descricao) b.appendChild(el('span', 'font-size:12px;color:#8b949e;', o.descricao));
                b.type = 'button'; b.addEventListener('click', () => responder(o.valor)); j.corpo.appendChild(b);
            });
            j.rodape.appendChild(botao('Cancelar', false, () => responder(null)));
            j.fundo.addEventListener('mousedown', e => { if (e.target === j.fundo) responder(null); });
            const aoFechar = new MutationObserver(() => { if (!doc.body.contains(j.fundo)) { aoFechar.disconnect(); responder(null); } });
            aoFechar.observe(doc.body, { childList: true });
        });
    }

    /** Botão "Colar SAP": lê a área de transferência (se o sistema deixar) ou abre uma caixinha para colar com Ctrl+V.
     *  Resolve com {ok: true, campos} ou null (cancelou). Erro de formato é mostrado na própria caixinha. */
    async function colarSap() {
        let texto = '', negado = false;
        try { texto = await raiz.navigator.clipboard.readText(); } catch (x) { negado = true; }
        const r = interpretarSap(texto);
        if (r.ok) return r;
        return new Promise(resolve => {
            const j = janela('Colar dados do SAP (Atlas)', 520);
            const caixa = el('textarea', ESTILO_CAMPO + 'height:170px;resize:vertical;font-family:Consolas,monospace;font-size:12px;');
            caixa.placeholder = 'Cole aqui (Ctrl+V) o JSON copiado do Atlas';
            const aviso = el('div', 'font-size:12px;color:#f85149;min-height:16px;', negado ? 'O sistema não deixou ler a área de transferência: cole o texto aqui.' : r.erro);
            j.corpo.append(caixa, aviso);
            let feito = false;
            const sair = v => { if (feito) return; feito = true; j.fechar(); resolve(v); };
            const usar = () => { const x = interpretarSap(caixa.value); if (x.ok) sair(x); else aviso.textContent = x.erro; };
            caixa.addEventListener('input', () => { if (interpretarSap(caixa.value).ok) usar(); });
            j.rodape.append(botao('Cancelar', false, () => sair(null)), botao('Usar', true, usar));
            j.fundo.addEventListener('mousedown', e => { if (e.target === j.fundo) sair(null); });
            caixa.focus();
        });
    }
    const ROTULOS = { titulo: 'título', nota: 'nº da nota', sap: 'Nº SAP', bairro: 'bairro', municipio: 'município', utd: 'UTD', utep: 'UTEP', projetista: 'projetista', data: 'data', id_investimento: 'ID investimento (ODI)', id_investimento_odd: 'ID investimento (ODD)' };

    /* ── Gravação ──
       Projeto aberto: o salvamento automático pode estar LIGADO (grava 2 s depois de cada alteração e ao trocar de tela) ou DESLIGADO (só o botão Salvar grava).
       Novo projeto nasce ligado; projeto aberto vem desligado. Projeto default: o automático (ligado por padrão) guarda um RASCUNHO fixo para recuperar depois. */
    let resumo = null;                 // ganchos do Resumo (registrarResumo)
    let dadosProntos = false;          // só grava Resumo depois de carregar o projeto (ou de a tela ter dados): nunca sobrescreve com a tela vazia
    let sujo = false, temporizador = null, pendente = null, gravando = null, suspenso = false, ultimaGravacao = null, erroAvisado = false;
    let coletorResultado = null;       // o Resultado (na moldura) registra como entregar as linhas dele quando o usuário clica em Salvar
    let rascunhoSessao = false, temporizadorRascunho = null;
    let ui = null;                     // elementos do menu (montarMenu)

    const autoDefaultLigado = () => { try { return raiz.localStorage.getItem('gridr_auto_default') !== '0'; } catch (x) { return true; } };
    const autoLigado = () => { const e = estado(); return e ? e.auto === true : autoDefaultLigado(); };
    /** Lê o estado ATUAL, mescla `patch` e grava (leitura+escrita sem pausa: nunca guarda assinatura velha nem apaga o que outra parte mudou). */
    function atualizarEstado(patch) { const e = estado(); if (!e) return null; const n = Object.assign({}, e, patch); gravarEstado(n); return n; }
    const temPendencias = () => !!(sujo || (pendente && Object.keys(pendente).length));
    const hora = d => d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    function definirProjeto(r, auto) {
        const pr = r.projeto;
        gravarEstado({ caminho: r.caminho, nome: pr.nome, numero: pr.numero_obra, sap: pr.sap, assinatura: r.assinatura, projeto: pr.projeto, auto: !!auto });
    }

    /** Texto e cor da etiqueta de estado: SEMPRE diz se o salvamento automático está ligado ou desligado. O aviso de arquivo alterado por fora é uma etiqueta à parte. */
    function rotuloEstado() {
        const e = estado();
        const auto = autoLigado() ? 'Automático ligado' : 'Automático desligado';
        if (!e) return autoDefaultLigado() ? { t: (ultimaGravacao ? 'Rascunho gravado às ' + hora(ultimaGravacao) : 'Rascunho') + ' · ' + auto, cor: '#d97706' } : { t: 'Sem rascunho · ' + auto, cor: '#8b949e' };
        if (suspenso) return { t: 'Não gravado · ' + auto, cor: '#d29922' };
        if (gravando) return { t: 'Gravando… · ' + auto, cor: '#d97706' };
        if (temPendencias()) return autoLigado() ? { t: 'Gravando em instantes… · ' + auto, cor: '#d97706' } : { t: 'Não salvo · ' + auto, cor: '#d29922' };
        return { t: (ultimaGravacao ? 'Salvo às ' + hora(ultimaGravacao) : 'Salvo') + ' · ' + auto, cor: '#3fb950' };
    }
    function atualizarMenu() {
        if (!ui) return;
        const e = estado(), r = rotuloEstado();
        ui.chip.textContent = e ? '📁 ' + e.nome : '📁 Projeto default';
        ui.chip.title = e ? e.caminho : 'Projeto default: o programa guarda um rascunho do que você faz. Salve para criar um projeto.';
        ui.dot.style.background = r.cor; ui.txt.textContent = r.t;
        ui.pill.title = 'Salvamento automático: ' + (autoLigado() ? 'ligado' : 'desligado') + (e ? ' · ' + e.caminho : ' · projeto default (rascunho)');
        ui.alerta.style.display = e && suspenso ? 'inline-flex' : 'none';
        const destaque = !!e && temPendencias() && !autoLigado();
        ui.salvar.style.color = destaque ? '#d97706' : '#e6edf3'; ui.split.style.borderColor = destaque ? '#d97706' : '#30363d';
        const lig = autoLigado();
        ui.sw.style.background = lig ? '#3fb950' : '#484f58'; ui.knob.style.left = lig ? '18px' : '2px';
        ui.hint.textContent = e ? (lig ? 'Ligado: grava 2 s depois de cada alteração.' : 'Desligado: você decide quando gravar (Salvar / Ctrl+S).')
                                : (lig ? 'Ligado: guarda um rascunho do que você faz, para recuperar depois.' : 'Desligado: nada é guardado até você salvar.');
    }

    async function conflito(modificadoEm) {
        const fmt = d => d.toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });
        const quando = modificadoEm && !isNaN(new Date(modificadoEm)) ? ' O arquivo foi modificado em ' + fmt(new Date(modificadoEm)) + (ultimaGravacao ? '; a sua última gravação nesta sessão foi às ' + hora(ultimaGravacao) : '') + '.' : '';
        const r = await perguntar('Projeto alterado por fora',
            'O arquivo do projeto foi modificado por outro programa ou computador depois da sua última gravação.' + quando + ' O que fazer?',
            [{ texto: 'Sobrescrever com o que está na tela', valor: 'sobrescrever', ok: true, descricao: 'Grava o que está na tela por cima do arquivo (o que mudou por fora nos mesmos dados é perdido).' },
             { texto: 'Recarregar do arquivo', valor: 'recarregar', descricao: 'Descarta o que está na tela e usa o arquivo atual.' },
             { texto: 'Decidir depois', valor: 'depois', descricao: 'Nada é gravado até você decidir (clique no aviso vermelho ao lado do Salvar).' }]);
        return r || 'depois';
    }
    async function escrever(parcial, forcar) {
        const e = estado(), p = ponte();
        if (!e || !p) return;
        const r = await p.projeto_gravar(e.caminho, JSON.stringify(parcial), forcar ? '' : (e.assinatura || ''), !!forcar, e.projeto || '');
        if (r && r.ok) {
            atualizarEstado({ assinatura: r.assinatura }); ultimaGravacao = new Date(); suspenso = false; erroAvisado = false;
        } else if (r && r.conflito) {
            const escolha = await conflito(r.modificado_em);
            if (escolha === 'sobrescrever') return escrever(parcial, true);
            if (escolha === 'recarregar') { await recarregarDoArquivo(); return; }
            suspenso = true; pendente = Object.assign({}, parcial, pendente || {});
        } else {
            if (!erroAvisado) { erroAvisado = true; avisar('Projeto não gravado: ' + ((r && r.erro) || 'erro desconhecido')); }
            pendente = Object.assign({}, parcial, pendente || {});
            suspenso = true;
        }
        atualizarMenu();
    }
    /** Grava o que estiver pendente, uma gravação por vez (junta o que chegar enquanto grava). */
    function drenar() {
        if (!gravando) {
            gravando = (async () => {
                try {
                    while (pendente && !suspenso) { const p = pendente; pendente = null; atualizarMenu(); await escrever(p, false); }
                } catch (x) { avisar('Projeto não gravado: ' + x.message); }
                finally { gravando = null; atualizarMenu(); }
            })();
            atualizarMenu();
        }
        return gravando;
    }
    function enviar(parcial) {                                 // com o automático ligado grava; desligado só guarda como pendente ("Não salvo")
        const d = dono();
        return d ? d.__enviar(parcial) : enviarLocal(parcial);
    }
    function enviarLocal(parcial) {
        if (!ativo()) return Promise.resolve();
        pendente = Object.assign(pendente || {}, parcial);
        if (!autoLigado() || suspenso) { atualizarMenu(); return Promise.resolve(); }
        return drenar();
    }
    function enviarAgora(parcial) {                            // ação explícita (Aplicar/OK nas propriedades): grava mesmo com o automático desligado
        const d = dono();
        if (d) return d.__enviarAgora(parcial);
        return enviarAgoraLocal(parcial);
    }
    function enviarAgoraLocal(parcial) {
        if (!ativo()) return Promise.resolve();
        pendente = Object.assign(pendente || {}, parcial); suspenso = false;
        return drenar();
    }
    function marcarSujo(imediato) {
        if (!resumo) return;
        if (estado()) {
            if (!ativo()) return;
            sujo = true; atualizarMenu();
            if (autoLigado()) { clearTimeout(temporizador); temporizador = setTimeout(salvarAutomatico, imediato ? 0 : ESPERA_MS); }
        } else if (ponte() && autoDefaultLigado()) {
            clearTimeout(temporizadorRascunho); temporizadorRascunho = setTimeout(gravarRascunho, imediato ? 0 : ESPERA_MS);
        }
    }
    /** Salvamento automático (só com ele ligado): grava o Resumo alterado. Troca de tela, voltar ao default e fechar também chamam. */
    async function salvarAutomatico() {
        clearTimeout(temporizador);
        if (!ativo() || !autoLigado()) return;
        if (resumo && dadosProntos && sujo) { sujo = false; await enviar({ dados: resumo.coletar() }); }
        else if (gravando) await gravando;
    }
    /** Botão Salvar do projeto aberto: grava tudo agora (Resumo, Resultado, SAP, nome/número), com o automático ligado ou não. Resolve true se gravou. */
    async function salvar(manual) {
        if (!ativo()) return false;
        clearTimeout(temporizador);
        const parcial = {};
        if (resumo && (sujo || manual)) {
            if (dadosProntos) parcial.dados = resumo.coletar();
            else if (manual) avisar('O Resumo não foi gravado: o projeto não foi carregado. Abra o projeto de novo.');
        }
        if (coletorResultado) { try { const r = coletorResultado(); if (r) parcial.resultado = r; } catch (x) { /* a tela do Resultado segue valendo */ } }
        sujo = false; suspenso = false;
        pendente = Object.assign(pendente || {}, parcial);
        await drenar();
        const ok = !suspenso;
        if (manual && ok) avisar('Projeto salvo.');
        return ok;
    }
    async function definirAuto(ligar) {
        if (estado()) {
            atualizarEstado({ auto: !!ligar }); atualizarMenu();
            if (ligar) { await salvar(false); avisar('Salvamento automático ligado.'); }
            else avisar('Salvamento automático desligado: use Salvar (Ctrl+S) para gravar.');
        } else {
            try { raiz.localStorage.setItem('gridr_auto_default', ligar ? '1' : '0'); } catch (x) { /* sem storage */ }
            atualizarMenu();
            if (ligar) gravarRascunho();
        }
    }
    /** Antes de sair do projeto aberto (abrir outro, criar novo, voltar ao default): com o automático desligado e pendências, pergunta. true = pode seguir. */
    async function confirmarSaida() {
        const e = estado();
        if (!e) return true;
        if (autoLigado()) { await salvarAutomatico(); return true; }
        if (!temPendencias()) return true;
        const r = await perguntar('Salvar alterações em “' + e.nome + '”?', 'Há alterações que ainda não foram gravadas no arquivo do projeto.',
            [{ texto: 'Salvar', valor: 'salvar', ok: true }, { texto: 'Não salvar', valor: 'nao', descricao: 'As alterações não gravadas serão perdidas.' }]);
        if (!r) return false;
        return r === 'salvar' ? await salvar(true) : true;
    }
    function salvarClick() {
        if (!ponte()) return Promise.resolve();
        if (!resumo) { const d = dono(); if (d) return d.__salvarClick(); }
        if (!estado()) { abrirJanela('novo'); return Promise.resolve(); }       // no default, Salvar cria o projeto
        return salvar(true);
    }
    function salvarComoClick() {
        if (!ponte()) return;
        if (!resumo) { const d = dono(); if (d) { d.__salvarComoClick(); return; } }
        abrirJanela(estado() ? 'salvarcomo' : 'novo');
    }
    async function resolverConflito() { suspenso = false; atualizarMenu(); await salvar(true); }

    /* ── Rascunho do projeto default ── */
    async function gravarRascunho() {
        const p = ponte();
        if (!p || estado() || !resumo || !autoDefaultLigado() || !resumo.temDados()) return;      // nunca grava a tela vazia por cima do rascunho
        const r = await p.projeto_rascunho_gravar(JSON.stringify({ dados: resumo.coletar() }), codigoRegras(), !rascunhoSessao);
        if (r && r.ok) { rascunhoSessao = true; ultimaGravacao = new Date(); }
        atualizarMenu();
    }
    async function recuperarRascunho() {
        const p = ponte();
        if (!p || !resumo) return;
        const info = await p.projeto_rascunho_info();
        const lista = (info && info.ok && info.rascunhos) || [];
        if (!lista.length) { await perguntar('Recuperar rascunho', 'Ainda não há rascunho do projeto default. Ele é gravado quando o salvamento automático está ligado e há dados no Resumo.', [{ texto: 'Entendi', valor: 'ok', ok: true }]); return; }
        const fmt = x => (x.salvo_em ? new Date(x.salvo_em).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' }) : '') + ' · ' + x.itens + ' linha(s)';
        const rotulos = { atual: 'Rascunho mais recente', anterior: 'Rascunho da sessão anterior' };
        const id = await perguntar('Recuperar rascunho do projeto default', 'Escolha o rascunho a recuperar para o Resumo.', lista.map((x, i) => ({ texto: rotulos[x.id] || x.id, valor: x.id, ok: i === 0, descricao: fmt(x) })));
        if (!id) return;
        const r = await p.projeto_rascunho_ler(id, codigoRegras());
        if (!r || !r.ok) { avisar('Não foi possível recuperar o rascunho: ' + ((r && r.erro) || '')); return; }
        let modo = 'abrir';
        if (resumo.temDados()) {
            modo = await perguntar('Recuperar rascunho', 'O Resumo já tem dados. O que deseja fazer?',
                [{ texto: 'Substituir', valor: 'abrir', ok: true, descricao: 'Apaga o que está na tela e usa o rascunho.' }, { texto: 'Mesclar', valor: 'mesclar', descricao: 'Junta o que está na tela com o rascunho.' }]);
            if (!modo) return;
        }
        if (modo === 'abrir') resumo.carregar(r.projeto.dados); else resumo.mesclar(r.projeto.dados);
        dadosProntos = true;
        if (estado()) marcarSujo(true); else gravarRascunho();
        avisar('Rascunho recuperado.');
    }

    /* ── Aplicar o projeto na tela ── */
    function aplicarNumero() {
        const e = estado();
        const inp = doc.getElementById('input-num-obra');
        try { if (e && e.numero) raiz.localStorage.setItem('numero_obra', e.numero); } catch (x) { /* sem storage */ }
        if (!inp) return;
        if (e) { inp.value = e.numero || ''; inp.readOnly = true; inp.title = 'Número da obra do projeto (altere em Projeto → Propriedades)'; }
        else { inp.readOnly = false; inp.title = 'Número da Obra'; }
    }
    async function recarregarDoArquivo() {
        const e = estado(), p = ponte();
        if (!e || !p) return;
        const r = await p.projeto_abrir(e.caminho, codigoRegras());
        if (!r || !r.ok) { avisar('Não foi possível recarregar o projeto: ' + ((r && r.erro) || '')); return; }
        definirProjeto(r, e.auto === true);
        if (resumo) { resumo.carregar(r.projeto.dados); dadosProntos = true; sujo = false; }
        suspenso = false; pendente = null; aplicarNumero(); atualizarMenu();
        avisar('Projeto recarregado do arquivo.');
    }

    /** Vincula o arquivo `r` ({caminho, assinatura, projeto}) como projeto atual (com o salvamento automático DESLIGADO); pergunta Abrir/Mesclar/Sobrescrever se o Resumo já tem dados. */
    async function ativarAberto(r) {
        const pr = r.projeto;
        if (!(await confirmarSaida())) return false;                          // havia outro projeto com alterações não salvas
        let modo = 'abrir';
        if (resumo && resumo.temDados()) {
            modo = await perguntar('Abrir “' + pr.nome + '”', 'O Resumo já tem dados. O que deseja fazer?',
                [{ texto: 'Abrir', valor: 'abrir', ok: true, descricao: 'Apaga o que está na tela e usa o que está no projeto.' },
                 { texto: 'Mesclar', valor: 'mesclar', descricao: 'Junta o que está na tela com o que vem do projeto.' },
                 { texto: 'Sobrescrever', valor: 'sobrescrever', descricao: 'Mantém o que está na tela e grava por cima do projeto (na próxima vez que você salvar).' }]);
            if (!modo) return false;
        }
        definirProjeto(r, false);
        suspenso = false; pendente = null; ultimaGravacao = null; sujo = false;
        if (resumo) {
            if (modo === 'abrir') resumo.carregar(pr.dados);
            else if (modo === 'mesclar') { resumo.mesclar(pr.dados); sujo = true; }
            else sujo = true;                                   // sobrescrever: a tela vale e vai por cima do arquivo quando salvar
            dadosProntos = true;
        }
        aplicarNumero(); atualizarMenu();
        avisar(pr.legado ? 'Projeto antigo aberto. Informe o número da obra em Projeto → Propriedades.' : 'Projeto “' + pr.nome + '” aberto. Salvamento automático desligado: use Salvar (Ctrl+S).');
        return true;
    }

    /* ── TASK-162: projeto entregue pelo plugin do nanoCAD (<desenho>.obra, gravado por ele na pasta do desenho) ── */
    const mesmoArquivo = (a, b) => String(a || '').replace(/\//g, '\\').toLowerCase() === String(b || '').replace(/\//g, '\\').toLowerCase();
    /** Abre o projeto que o plugin gravou. Se ele já é o projeto aberto, relê do arquivo (o desenho atualizou as linhas); senão, o mesmo
     *  caminho do menu Abrir (pergunta Abrir/Mesclar/Sobrescrever se o Resumo já tem dados). Devolve true se abriu. */
    async function abrirDoPlugin(caminho, projetoCodigo) {
        const p = ponte();
        if (!p) return false;
        const r = await p.projeto_abrir(caminho, projetoCodigo || codigoRegras());
        if (!r || !r.ok) {
            await perguntar('Projeto do nanoCAD', 'Não foi possível abrir o projeto do desenho: ' + ((r && r.erro) || 'erro ao ler o arquivo.'), [{ texto: 'Entendi', valor: 'ok', ok: true }]);
            return false;
        }
        const e = estado();
        if (e && mesmoArquivo(e.caminho, r.caminho)) {
            if (temPendencias()) {
                const v = await perguntar('Projeto atualizado pelo nanoCAD', 'O desenho gravou linhas novas em “' + r.projeto.nome + '”, mas a tela tem alterações não salvas.',
                    [{ texto: 'Usar o do desenho', valor: 'desenho', ok: true, descricao: 'Recarrega o projeto do arquivo (as alterações da tela se perdem).' },
                     { texto: 'Manter a tela', valor: 'tela', descricao: 'Não recarrega agora; ao salvar, o programa avisa que o arquivo mudou.' }]);
                if (v !== 'desenho') return false;
            }
            definirProjeto(r, e.auto === true);
            if (resumo) { resumo.carregar(r.projeto.dados); dadosProntos = true; sujo = false; }
            suspenso = false; pendente = null; aplicarNumero(); atualizarMenu();
            avisar('Projeto “' + r.projeto.nome + '” atualizado pelo nanoCAD.');
            return true;
        }
        return ativarAberto(r);
    }

    /* ── Janela do projeto (Novo / Abrir / Propriedades / Salvar como) ── */
    function abrirJanelaSap(valores, aoAplicar) {
        const j = janela('Informações do relatório SAP', 520);
        const grade = el('div', 'display:grid;grid-template-columns:1fr 1fr;gap:10px;');
        const campos = {};
        CAMPOS_SAP.forEach(([k, rot]) => { const c = campo(rot, (valores || {})[k] || ''); if (k === 'id_investimento_odd') c.bloco.style.gridColumn = '1 / -1'; campos[k] = c.input; grade.appendChild(c.bloco); });
        j.corpo.appendChild(el('div', 'font-size:12px;color:#8b949e;', 'O título do relatório é o nome do projeto e o nº da nota é o número da obra. Estes dados vão preenchidos ao gerar os relatórios ODI e ODD.'));
        j.corpo.appendChild(el('div', '', null, [botao('📋 Colar SAP', false, async () => {
            const r = await colarSap();
            if (!r) return;
            CAMPOS_SAP.forEach(([k]) => { if (r.campos[k]) campos[k].value = r.campos[k]; });
            avisar('Preenchido: ' + CAMPOS_SAP.filter(([k]) => r.campos[k]).map(([k]) => ROTULOS[k]).join(', ') + '.');
        })]));
        j.corpo.appendChild(grade);
        j.rodape.append(botao('Cancelar', false, j.fechar), botao('OK', true, () => {
            const v = {}; CAMPOS_SAP.forEach(([k]) => { v[k] = campos[k].value.trim(); });
            j.fechar(); aoAplicar(v);
        }));
    }
    /** Resultado gravado no projeto atual (pendente em memória ou o do arquivo), para o "Salvar como" levar junto. */
    async function resultadoAtual() {
        if (coletorResultado) { try { const r = coletorResultado(); if (r && r.linhas && r.linhas.length) return r; } catch (x) { /* segue */ } }
        if (pendente && pendente.resultado) return pendente.resultado;
        const e = estado(), p = ponte();
        if (!e || !p) return null;
        const r = await p.projeto_abrir(e.caminho, e.projeto || codigoRegras());
        return (r && r.ok && r.projeto && r.projeto.resultado) || null;
    }
    /** modo: 'novo' | 'abrir' (com `aberto` = retorno do projeto_abrir) | 'propriedades' | 'salvarcomo' */
    function abrirJanela(modo, aberto) {
        const p = ponte();
        if (!p) return;
        const atual = estado();
        const copia = modo === 'salvarcomo';
        const origem = modo === 'abrir' ? aberto.projeto
            : ((modo === 'propriedades' || copia) && atual ? { nome: copia ? atual.nome + ' (cópia)' : atual.nome, numero_obra: atual.numero, sap: atual.sap || {} } : { nome: '', numero_obra: '', sap: {} });
        const titulo = modo === 'novo' ? 'Novo projeto' : (copia ? 'Salvar como…' : (modo === 'abrir' ? 'Abrir projeto' : 'Projeto: ' + (atual ? atual.nome : '')));
        const j = janela(titulo, 460);
        const criaArquivo = modo === 'novo' || copia;
        let criado = !criaArquivo, sap = Object.assign({}, origem.sap || {}), alterado = false, pasta = copia && atual ? pastaDoCaminho(atual.caminho) : '', ativado = modo === 'propriedades';
        const cNome = campo('Nome do projeto', origem.nome), cNum = campo('Número de obra (usado na exportação do REC)', origem.numero_obra);
        const dica = el('span', 'font-size:12px;color:#8b949e;display:block;margin-top:4px;', '');
        const cPasta = campo('Pasta do projeto', pasta, dica);
        const btnEscolher = botao('Escolher…', false, async () => {
            const r = await p.projeto_escolher_pasta(pasta);
            if (r && r.ok) { pasta = r.pasta; cPasta.input.value = pasta; mudou(); }
        });
        cPasta.input.readOnly = true;
        const linhaPasta = el('div', 'display:flex;gap:8px;'); cPasta.bloco.insertBefore(linhaPasta, cPasta.input); linhaPasta.appendChild(cPasta.input); linhaPasta.appendChild(btnEscolher);
        const btnSap = botao('📋 Informações do relatório SAP…', false, () => abrirJanelaSap(sap, v => { sap = v; mudou(); }));
        j.corpo.append(cNome.bloco, cNum.bloco);
        if (criaArquivo) j.corpo.appendChild(cPasta.bloco);
        else j.corpo.appendChild(el('div', 'font-size:12px;color:#8b949e;word-break:break-all;', 'Arquivo: ' + (modo === 'abrir' ? aberto.caminho : (atual ? atual.caminho : ''))));
        const btnColar = botao('📋 Colar SAP', false, async () => {
            const r = await colarSap();
            if (!r) return;
            const c = r.campos;
            if (c.titulo) cNome.input.value = c.titulo;                                   // o título do SAP vira o nome do projeto
            if (c.nota) cNum.input.value = c.nota;                                        // e a nota, o número de obra
            sap = Object.assign({}, sap); CAMPOS_SAP.forEach(([k]) => { if (c[k]) sap[k] = c[k]; });
            mudou();
            avisar('Preenchido: ' + Object.keys(c).map(k => ROTULOS[k]).join(', ') + '.');
        });
        j.corpo.appendChild(el('div', 'display:flex;gap:8px;flex-wrap:wrap;', null, [btnColar, btnSap]));
        if (modo === 'novo' && resumo && resumo.temDados()) j.corpo.appendChild(el('div', 'font-size:12px;color:#8b949e;', 'O que está na tela será salvo no projeto novo. O salvamento automático começa ligado.'));
        if (copia) j.corpo.appendChild(el('div', 'font-size:12px;color:#8b949e;', 'O projeto atual continua como está no disco; a tela passa a ser este novo arquivo.'));
        const bAplicar = botao('Aplicar', false, () => aplicar(false)), bCriar = botao(copia ? 'Salvar como' : 'Criar', true, criar), bOk = botao('OK', true, () => aplicar(true));
        j.rodape.append(...(modo === 'novo' ? [bAplicar, bCriar, bOk] : (copia ? [botao('Cancelar', false, j.fechar), bCriar] : [botao('Cancelar', false, j.fechar), bAplicar, bOk])));
        const nomeOk = () => cNome.input.value.trim().length > 0;
        function mudou() {
            alterado = true;
            if (criaArquivo && !criado) dica.textContent = pasta && nomeOk() ? 'Será criado: ' + pasta + (pasta.indexOf('/') >= 0 && pasta.indexOf('\\') < 0 ? '/' : '\\') + cNome.input.value.trim() + '.obra' : '';
            estados();
        }
        function estados() {
            habilitar(bCriar, criaArquivo && !criado && nomeOk() && !!pasta);
            habilitar(bAplicar, criado && ativado && alterado && nomeOk());
            habilitar(bOk, criado && nomeOk() && (modo === 'abrir' || ativado || modo === 'propriedades'));
            if (modo === 'abrir') habilitar(bAplicar, nomeOk());
        }
        [cNome.input, cNum.input].forEach(i => i.addEventListener('input', () => mudou()));
        async function criar() {
            habilitar(bCriar, false);
            if (!copia && !(await confirmarSaida())) { estados(); return; }          // criar um projeto novo deixa o aberto: pergunta se há pendências
            const meta = { nome: cNome.input.value.trim(), numero_obra: cNum.input.value.trim(), sap };
            if (resumo && (copia ? dadosProntos : resumo.temDados())) meta.dados = resumo.coletar();            // o que já está na tela vai para o projeto novo
            if (copia) { const res = await resultadoAtual(); if (res) meta.resultado = res; }
            let r = await p.projeto_criar(pasta, JSON.stringify(meta), codigoRegras());
            if (r && r.existe) {
                await perguntar('Projeto já existe', 'Já existe “' + meta.nome + '.obra” nessa pasta. Escolha outro nome ou outra pasta.', [{ texto: 'Entendi', valor: 'ok', ok: true }]);
                estados(); return;
            }
            if (!r || !r.ok) { avisar('Não foi possível criar o projeto: ' + ((r && r.erro) || '')); estados(); return; }
            gravarEstado({ caminho: r.caminho, nome: meta.nome, numero: meta.numero_obra, sap: meta.sap, assinatura: r.assinatura, projeto: codigoRegras(),
                           auto: copia ? !!(atual && atual.auto === true) : true });          // novo: automático ligado; salvar como: mantém o que estava
            criado = true; ativado = true; alterado = false; suspenso = false; pendente = null; sujo = false; ultimaGravacao = new Date();
            if (resumo) dadosProntos = true;
            aplicarNumero(); atualizarMenu(); estados();
            avisar((copia ? 'Projeto salvo como: ' : 'Projeto criado: ') + r.caminho);
            if (copia) j.fechar();
        }
        async function aplicar(fechar) {
            if (modo === 'abrir' && !ativado) {
                const r = JSON.parse(JSON.stringify(aberto));
                r.projeto.nome = cNome.input.value.trim(); r.projeto.numero_obra = cNum.input.value.trim(); r.projeto.sap = sap;
                if (!(await ativarAberto(r))) return;
                ativado = true; alterado = true;                        // nome/número/SAP editados na janela vão para o arquivo logo abaixo
            }
            const e = estado();
            if (e && alterado) {
                atualizarEstado({ nome: cNome.input.value.trim(), numero: cNum.input.value.trim(), sap });
                await enviarAgora({ nome: cNome.input.value.trim(), numero_obra: cNum.input.value.trim(), sap });
                aplicarNumero(); atualizarMenu();
            } else if (suspenso) { await resolverConflito(); }
            alterado = false; estados();
            if (fechar) j.fechar();
        }
        mudou(); alterado = false; estados();
        cNome.input.focus();
    }

    async function abrirArquivo() {
        const p = ponte();
        if (!p) return;
        const r = await p.projeto_escolher_arquivo(codigoRegras());
        if (!r || r.cancelado) return;
        if (!r.ok) { await perguntar('Não foi possível abrir', r.erro || 'Erro ao ler o arquivo.', [{ texto: 'Entendi', valor: 'ok', ok: true }]); return; }
        abrirJanela('abrir', r);
    }
    async function voltarAoDefault() {
        if (!(await confirmarSaida())) return;
        gravarEstado(null); pendente = null; sujo = false; suspenso = false; ultimaGravacao = null;
        aplicarNumero(); atualizarMenu();
        avisar('Voltou ao projeto default.');
        gravarRascunho();                                                      // o que está na tela passa a ser guardado como rascunho
    }

    /* ── Barra do Resumo (host: #projeto-menu-host): 📁 projeto · [💾 Salvar | ▾] · etiqueta de estado ── */
    const DISCO = '<svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true"><path d="M5 3h11l4 4v13a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/><path d="M7.5 3v5h7.5V3M7.5 21v-7h9v7" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/></svg>';
    const SETA = '<svg viewBox="0 0 24 24" width="14" height="14" aria-hidden="true"><path d="M6 9l6 6 6-6" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/></svg>';
    const SINCRONIA = '<svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true"><path d="M20 12a8 8 0 0 0-13.7-5.6M4 12a8 8 0 0 0 13.7 5.6" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/><path d="M6.8 2.8 6 7l4.2-.8M17.2 21.2 18 17l-4.2.8" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>';
    function montarMenu() {
        const host = doc.getElementById('projeto-menu-host');
        if (!host || host.dataset.pronto || !ponte()) return;
        host.dataset.pronto = '1';
        host.style.cssText = 'position:relative;display:inline-flex;align-items:center;gap:8px;';
        const chip = el('button', 'padding:6px 12px;border-radius:6px;border:1px solid #1f6feb;background:#1f6feb22;color:#e6edf3;font:inherit;font-size:0.82rem;cursor:pointer;max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;', '📁 Projeto default');
        chip.id = 'btn-projeto'; chip.type = 'button';
        const split = el('span', 'display:inline-flex;border:1px solid #30363d;border-radius:6px;overflow:hidden;background:#21262d;');
        const salvarBtn = el('button', 'display:inline-flex;align-items:center;gap:6px;padding:5px 10px;border:0;background:transparent;color:#e6edf3;font:inherit;font-size:0.82rem;cursor:pointer;');
        salvarBtn.id = 'btn-projeto-salvar'; salvarBtn.type = 'button'; salvarBtn.title = 'Salvar (Ctrl+S)'; salvarBtn.innerHTML = DISCO; salvarBtn.appendChild(doc.createTextNode('Salvar'));
        const seta = el('button', 'display:inline-flex;align-items:center;padding:5px 6px;border:0;border-left:1px solid #30363d;background:transparent;color:#8b949e;font:inherit;cursor:pointer;');
        seta.id = 'btn-projeto-menu'; seta.type = 'button'; seta.title = 'Salvar como, salvamento automático e projetos'; seta.innerHTML = SETA;
        split.append(salvarBtn, seta);
        const pill = el('span', 'display:inline-flex;align-items:center;gap:6px;font-size:0.72rem;padding:3px 10px;border-radius:12px;border:1px solid #30363d;background:#21262d;white-space:nowrap;');
        pill.id = 'projeto-status';
        const dot = el('span', 'width:8px;height:8px;border-radius:50%;display:inline-block;background:#8b949e;'), txt = el('span', 'color:#e6edf3;');
        pill.append(dot, txt);
        const alerta = el('button', 'display:none;align-items:center;gap:5px;font:inherit;font-size:0.72rem;padding:3px 10px;border-radius:12px;border:1px solid #f85149;background:#f8514922;color:#f85149;cursor:pointer;white-space:nowrap;', '⚠ Arquivo alterado por fora — resolver');
        alerta.type = 'button'; alerta.title = 'O arquivo do projeto mudou no disco depois da sua última gravação. Clique para escolher o que fazer.';
        alerta.addEventListener('click', () => { if (suspenso) resolverConflito(); });
        const menu = el('div', 'position:absolute;top:100%;left:0;margin-top:4px;min-width:300px;background:#21262d;border:1px solid #30363d;border-radius:6px;padding:4px 0;z-index:9000;display:none;box-shadow:0 8px 24px #0008;');
        const fechaMenu = () => { menu.style.display = 'none'; };
        const item = (rotulo, atalho, aoClicar, visivel) => {
            const d = el('div', 'padding:7px 14px;cursor:pointer;font-size:0.85rem;color:#e6edf3;display:flex;justify-content:space-between;gap:14px;', null, [el('span', '', rotulo), el('small', 'color:#8b949e;', atalho || '')]);
            d.addEventListener('mouseenter', () => { d.style.background = '#30363d'; }); d.addEventListener('mouseleave', () => { d.style.background = ''; });
            d.addEventListener('click', () => { fechaMenu(); aoClicar(); }); d.__visivel = visivel; return d;
        };
        const linha = () => el('div', 'border-top:1px solid #30363d;margin:4px 0;');
        // Salvamento automático: chave liga/desliga (não fecha o menu)
        const swBox = el('div', 'padding:7px 14px 2px;cursor:pointer;font-size:0.85rem;color:#e6edf3;display:flex;justify-content:space-between;align-items:center;gap:14px;');
        const rotAuto = el('span', 'display:inline-flex;align-items:center;gap:8px;'); const icoAuto = el('span', 'color:#3fb950;display:inline-flex;'); icoAuto.innerHTML = SINCRONIA;
        rotAuto.append(icoAuto, el('span', '', 'Salvamento automático'));
        const sw = el('span', 'width:34px;height:18px;border-radius:9px;background:#484f58;position:relative;display:inline-block;flex:none;transition:background .15s;');
        const knob = el('span', 'position:absolute;top:2px;left:2px;width:14px;height:14px;border-radius:50%;background:#fff;transition:left .15s;'); sw.appendChild(knob);
        const hint = el('div', 'padding:0 14px 6px 40px;font-size:0.72rem;color:#8b949e;');
        swBox.append(rotAuto, sw); swBox.addEventListener('click', ev => { ev.stopPropagation(); definirAuto(!autoLigado()); });
        const itens = [item('💾 Salvar', 'Ctrl+S', () => salvarClick()), item('💾 Salvar como…', 'Ctrl+Shift+S', () => salvarComoClick()), linha(), swBox, hint, linha(),
            item('📁 Novo projeto…', '', () => abrirJanela('novo')), item('📂 Abrir projeto…', '', abrirArquivo), item('⚙ Propriedades do projeto…', '', () => abrirJanela('propriedades'), () => !!estado()),
            item('🕘 Recuperar rascunho do default…', '', recuperarRascunho), item('↩ Voltar ao projeto default', '', voltarAoDefault, () => !!estado())];
        itens.forEach(i => menu.appendChild(i));
        const alterna = ev => { ev.stopPropagation(); itens.forEach(i => { if (i.__visivel) i.style.display = i.__visivel() ? 'flex' : 'none'; }); atualizarMenu(); menu.style.display = menu.style.display === 'none' ? 'block' : 'none'; };
        chip.addEventListener('click', alterna); seta.addEventListener('click', alterna);
        salvarBtn.addEventListener('click', ev => { ev.stopPropagation(); fechaMenu(); salvarClick(); });
        doc.addEventListener('click', fechaMenu);
        host.append(chip, split, pill, alerta, menu);
        ui = { chip, split, salvar: salvarBtn, pill, dot, txt, alerta, sw, knob, hint, menu };
        atualizarMenu();
    }

    /* ── Resumo: ganchos e início ── */
    async function iniciar() {
        montarMenu();
        const p = ponte(), e = estado();
        if (!p) return;
        if (raiz.AbrirNaApp && !raiz.AbrirNaApp.__projeto) {                // grava antes de abrir Resultado / Tabela de orçamento (só com o automático ligado)
            const original = raiz.AbrirNaApp.abrir;
            raiz.AbrirNaApp.abrir = function () { salvarAutomatico(); return original.apply(this, arguments); };
            raiz.AbrirNaApp.__projeto = true;
        }
        if (!e) { aplicarNumero(); atualizarMenu(); return; }
        const existe = await p.projeto_assinatura(e.caminho);
        if (!existe || !existe.ok || !existe.existe) {
            gravarEstado(null); aplicarNumero(); atualizarMenu();
            avisar('O projeto “' + e.nome + '” não foi encontrado. Voltando ao projeto default.');
            return;
        }
        atualizarEstado({ assinatura: existe.assinatura });                  // o que está no disco ao abrir a tela é o ponto de partida: a assinatura guardada pode ser de uma gravação cuja resposta se perdeu ao fechar
        if (resumo && !resumo.temDados()) {                                  // continua de onde parou
            const r = await p.projeto_abrir(e.caminho, codigoRegras());
            if (r && r.ok) {
                definirProjeto(r, e.auto === true);
                resumo.carregar(r.projeto.dados); dadosProntos = true; sujo = false;
            } else {
                avisar('Projeto “' + e.nome + '” não foi carregado: ' + ((r && r.erro) || '') + ' A gravação do Resumo fica suspensa até você abrir o projeto de novo.');
            }
        } else dadosProntos = true;
        aplicarNumero(); atualizarMenu();
    }
    function registrarResumo(ganchos) {
        resumo = ganchos;
        iniciar();
    }

    /* ── Resultado do orçamento ── */
    let observador = null, termoResultado = null;
    /** As linhas do Resultado do projeto (pendentes em memória ou gravadas no arquivo), se foram calculadas com o mesmo payload (`assinatura`); senão null. */
    async function carregarResultado(assinatura) {
        const e = estado(), p = ponte();
        if (!e || !p || !assinatura) return null;
        const d = dono(), pend = d ? d.__pendenteResultado() : ((pendente && pendente.resultado) || null);
        if (pend && pend.assinatura === assinatura && Array.isArray(pend.linhas) && pend.linhas.length) return pend;      // com o automático desligado, as edições ainda estão só em memória
        const r = await p.projeto_abrir(e.caminho, e.projeto || codigoRegras());
        const res = r && r.ok && r.projeto && r.projeto.resultado;
        return res && res.assinatura === assinatura && Array.isArray(res.linhas) && res.linhas.length ? res : null;
    }
    /** Acompanha a tabela do Resultado: as linhas vão para o projeto (grava com o automático ligado; senão ficam pendentes até o Salvar). */
    function acompanharResultado(tbody, coletar, assinatura) {
        if (!ativo() || !tbody) return;
        if (observador) observador.disconnect();
        const entregar = () => ({ assinatura, linhas: coletar() });
        const gravar = () => { try { if (ativo()) enviar({ resultado: entregar() }); } catch (x) { /* a tela segue valendo */ } };
        const agendar = () => { clearTimeout(termoResultado); termoResultado = setTimeout(gravar, ESPERA_MS); };
        observador = new MutationObserver(agendar);
        observador.observe(tbody, { childList: true, subtree: true, characterData: true });
        const dn = dono() || api;
        dn.__registrarColetor(entregar);                                       // o botão Salvar (na janela principal) pede as linhas atuais
        agendar();
        raiz.addEventListener('pagehide', () => { clearTimeout(termoResultado); gravar(); dn.__registrarColetor(null); });
    }
    /** Pop-up de confirmação do número da obra na exportação do REC; resolve com o número ou null (cancelou). Não altera o projeto. */
    function confirmarNumeroObra(atual) {
        return new Promise(resolve => {
            const j = janela('Confirmar número de obra (REC)', 420);
            const c = campo('Número de obra', atual || '');
            j.corpo.append(c.bloco, el('div', 'font-size:12px;color:#8b949e;', 'Vale só para esta exportação. O número do projeto muda em Projeto → Propriedades.'));
            let feito = false;
            const sair = v => { if (feito) return; feito = true; j.fechar(); resolve(v); };
            const ok = () => { const v = c.input.value.trim(); if (v) sair(v); };
            c.input.addEventListener('keydown', e => { if (e.key === 'Enter') ok(); });
            j.rodape.append(botao('Cancelar', false, () => sair(null)), botao('Exportar', true, ok));
            j.fundo.addEventListener('mousedown', e => { if (e.target === j.fundo) sair(null); });
            c.input.focus(); c.input.select();
        });
    }
    function salvarSap(cab) {
        if (!estado() || !cab) return Promise.resolve();
        const sap = {}; CAMPOS_SAP.forEach(([k]) => { if (typeof cab[k] === 'string') sap[k] = cab[k].trim(); });
        atualizarEstado({ sap });
        return enviar({ sap });
    }

    // Fechar o programa / trocar de página: grava o que estiver pendente (melhor esforço; a gravação automática já cobre quase tudo)
    raiz.addEventListener('pagehide', () => { if (resumo) { salvarAutomatico(); if (!estado()) gravarRascunho(); } });
    doc.addEventListener('visibilitychange', () => { if (doc.visibilityState === 'hidden' && resumo) { salvarAutomatico(); if (!estado()) gravarRascunho(); } });
    // Com projeto aberto e o automático DESLIGADO, avisa ao fechar a janela com alterações não salvas
    raiz.addEventListener('beforeunload', e => { if (resumo && ponte() && estado() && !autoLigado() && temPendencias()) { e.preventDefault(); e.returnValue = ''; } });
    // Ctrl+S salva, Ctrl+Shift+S salva como (só no aplicativo desktop)
    doc.addEventListener('keydown', e => {
        if (!(e.ctrlKey || e.metaKey) || e.altKey || String(e.key).toLowerCase() !== 's' || !ponte()) return;
        if (!resumo && !dono()) return;
        e.preventDefault();
        if (e.shiftKey) salvarComoClick(); else salvarClick();
    }, true);
    // Qualquer edição na página marca o projeto como alterado; a gravação espera 2 s de pausa
    ['input', 'change', 'paste', 'cut'].forEach(ev => doc.addEventListener(ev, e => { if (resumo && !(e.target && e.target.closest && e.target.closest('[data-gridr-projeto]'))) marcarSujo(false); }, true));
    raiz.addEventListener('pywebviewready', () => { montarMenu(); });
    doc.addEventListener('DOMContentLoaded', () => { montarMenu(); let n = 0; const t = setInterval(() => { montarMenu(); if (++n > 40 || (doc.getElementById('btn-projeto'))) clearInterval(t); }, 250); });

    Object.assign(api, {
        ativo, estado, pasta: () => { const e = estado(); return e ? pastaDoCaminho(e.caminho) : ''; }, numeroObra: () => { const e = estado(); return e ? (e.numero || '') : ''; },
        nome: () => { const e = estado(); return e ? e.nome : ''; }, sap: () => { const e = estado(); return e && e.sap ? e.sap : {}; },
        registrarResumo, marcarSujo, colarSap, ROTULOS, abrirDoPlugin, temPonte: () => !!ponte(), salvarAgora: salvarAutomatico, abrirJanela, voltarAoDefault, carregarResultado, acompanharResultado, confirmarNumeroObra, salvarSap,
        salvar: salvarClick, salvarComo: salvarComoClick, definirAuto, recuperarRascunho, autoLigado, temPendencias,
        __enviar: enviarLocal, __enviarAgora: enviarAgoraLocal, __salvarClick: salvarClick, __salvarComoClick: salvarComoClick,
        __registrarColetor: fn => { coletorResultado = fn; }, __pendenteResultado: () => (pendente && pendente.resultado) || null
    });
    raiz.GridRProjeto = api;
})(typeof window !== 'undefined' ? window : globalThis);
