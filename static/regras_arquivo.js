/* ═══════════════════════════════════════════════════════════════════════════
   regras_arquivo.js — Exportar/importar regras do leitor e ajustes em arquivo JSON (TASK-061).
   Lógica pura (interpretar, resumir, aplicar) + um pouco de DOM (baixar, ler arquivo, janela "Importar").
   A importação NUNCA grava: devolve a lista nova, que o chamador coloca no RASCUNHO da tela; quem grava é o Salvar de sempre (validação do
   servidor + histórico com reversão). Conteúdo do arquivo entra na página só por textContent.
   Funciona no navegador (window.RegrasArquivo) e em Node (module.exports), para os testes.
═══════════════════════════════════════════════════════════════════════════ */
(function (root) {
    'use strict';
    const LIMITE_BYTES = 1024 * 1024;
    const LIMITE_ITENS = 1000;
    const LIMITE_REGEX = 500;
    const VERSAO = 1;
    const FORMATO = { regras: 'leitor-regras', ajustes: 'leitor-ajustes' };

    class ErroArquivo extends Error {
        constructor(mensagens) { super(mensagens.join(' ')); this.mensagens = mensagens; }
    }

    function estavel(v) {
        if (Array.isArray(v)) return '[' + v.map(estavel).join(',') + ']';
        if (v && typeof v === 'object') return '{' + Object.keys(v).sort().map(k => JSON.stringify(k) + ':' + estavel(v[k])).join(',') + '}';
        return JSON.stringify(v === undefined ? null : v);
    }
    const clone = x => JSON.parse(JSON.stringify(x));
    /** Regra do leitor sem a `ordem` (que muda ao acrescentar) para comparar "é a mesma regra?". */
    function semOrdem(r) { const c = clone(r); delete c.ordem; return c; }

    function nomeDeArquivo(tipo, projeto, tabela) {
        const seguro = s => String(s || 'projeto').replace(/[^\w\-.]+/g, '_').replace(/^[_.]+|[_.]+$/g, '').slice(0, 40) || 'projeto';
        return tipo === 'ajustes' ? `ajustes-${seguro(projeto)}.json` : `regras-leitor-${seguro(tabela)}-${seguro(projeto)}.json`;
    }

    function montarEnvelope(tipo, projeto, tabela, itens) {
        const env = { formato: FORMATO[tipo], versao: VERSAO, exportado_em: new Date().toISOString(), projeto: String(projeto || '') };
        if (tipo === 'regras') env.tabela = tabela;
        env[tipo === 'ajustes' ? 'ajustes' : 'regras'] = clone(itens);
        return env;
    }

    /** Texto do arquivo -> {itens, projeto, envelope}. Aceita o envelope exportado ou a lista pura (como as sementes). Levanta ErroArquivo. */
    function interpretar(texto, tipo, tabelaEsperada) {
        const erros = [];
        if (typeof texto !== 'string' || !texto.length) throw new ErroArquivo(['O arquivo está vazio.']);
        if (texto.length > LIMITE_BYTES) throw new ErroArquivo([`Arquivo grande demais (máx. ${Math.round(LIMITE_BYTES / 1024 / 1024)} MB).`]);
        let dados;
        try { dados = JSON.parse(texto); } catch (e) { throw new ErroArquivo(['O arquivo não é um JSON válido.']); }
        let itens, projeto = null, envelope = false;
        if (Array.isArray(dados)) {
            itens = dados;
        } else if (dados && typeof dados === 'object') {
            envelope = true;
            if (dados.formato !== FORMATO[tipo]) {
                const outro = Object.values(FORMATO).includes(dados.formato);
                throw new ErroArquivo([outro ? `Este arquivo é de outro cadastro (${dados.formato}); use "Importar" no cadastro certo.` : 'O arquivo não foi exportado por este programa (formato desconhecido).']);
            }
            if (!Number.isInteger(dados.versao) || dados.versao < 1) throw new ErroArquivo(['Versão do arquivo inválida.']);
            if (dados.versao > VERSAO) throw new ErroArquivo([`Este arquivo é de uma versão mais nova (${dados.versao}); atualize o programa.`]);
            if (tipo === 'regras' && dados.tabela && tabelaEsperada && dados.tabela !== tabelaEsperada) {
                throw new ErroArquivo([`Este arquivo é da tabela de ${dados.tabela}, mas você está importando em ${tabelaEsperada}.`]);
            }
            itens = tipo === 'ajustes' ? dados.ajustes : dados.regras;
            projeto = typeof dados.projeto === 'string' && dados.projeto ? dados.projeto : null;
        } else {
            throw new ErroArquivo(['O arquivo precisa ser uma lista de regras ou um arquivo exportado por este programa.']);
        }
        if (!Array.isArray(itens)) throw new ErroArquivo(['O arquivo não tem a lista de regras.']);
        if (!itens.length) throw new ErroArquivo(['O arquivo não tem nenhuma regra.']);
        if (itens.length > LIMITE_ITENS) throw new ErroArquivo([`Itens demais no arquivo (${itens.length}; máx. ${LIMITE_ITENS}).`]);
        itens.forEach((r, i) => {
            if (!r || typeof r !== 'object' || Array.isArray(r)) { erros.push(`Item #${i + 1}: precisa ser um objeto.`); return; }
            [['texto_regex', r.texto_regex], ['ativo_regex', r.ativo_regex], ['vizinhanca.regex', r.vizinhanca && r.vizinhanca.regex]].forEach(([campo, v]) => {
                if (typeof v === 'string' && v.length > LIMITE_REGEX) erros.push(`Item #${i + 1}: '${campo}' passa de ${LIMITE_REGEX} caracteres.`);
            });
        });
        if (erros.length) throw new ErroArquivo(erros.slice(0, 20));
        return { itens, projeto, envelope };
    }

    /** O que a importação faria na lista `atual`. `opcoes.chave` = 'id' (ajustes) ou null (regras do leitor, comparadas sem a ordem);
     *  `opcoes.limpar` normaliza um item antes de comparar. Devolve {novas, atualizadas, iguais, removidas} (contagens). */
    function resumir(atual, importadas, modo, opcoes) {
        const chave = (opcoes && opcoes.chave) || null;
        const limpar = (opcoes && opcoes.limpar) || (x => x);
        const norm = r => estavel(chave ? limpar(r) : semOrdem(r));
        const sai = { novas: 0, atualizadas: 0, iguais: 0, removidas: 0 };
        if (chave) {
            const porId = new Map(atual.map(r => [r[chave], r]));
            importadas.forEach(r => {
                const ex = porId.get(r[chave]);
                if (!ex) sai.novas++; else if (norm(ex) === norm(r)) sai.iguais++; else sai.atualizadas++;
            });
            if (modo === 'substituir') { const ids = new Set(importadas.map(r => r[chave])); sai.removidas = atual.filter(r => !ids.has(r[chave])).length; }
        } else {
            const existentes = new Set(atual.map(norm));
            const vistos = new Set();
            importadas.forEach(r => {
                const k = norm(r);
                if (existentes.has(k) || vistos.has(k)) sai.iguais++; else { sai.novas++; vistos.add(k); }
            });
            if (modo === 'substituir') { const imp = new Set(importadas.map(norm)); sai.removidas = atual.filter(r => !imp.has(norm(r))).length; }
        }
        return sai;
    }

    /** Lista resultante. Acrescentar: regras do leitor novas entram no FIM (a `ordem` é renumerada depois da última, por fase no Processamento,
     *  mantendo a ordem relativa do arquivo); ajustes: novos no fim e os de mesmo id são atualizados no lugar. Substituir: a lista do arquivo. */
    function aplicar(atual, importadas, modo, opcoes) {
        const chave = (opcoes && opcoes.chave) || null;
        const limpar = (opcoes && opcoes.limpar) || (x => x);
        if (modo === 'substituir') return clone(importadas);
        if (chave) {
            const porId = new Map(importadas.map(r => [r[chave], r]));
            const usados = new Set();
            const saida = atual.map(r => { if (porId.has(r[chave])) { usados.add(r[chave]); return clone(porId.get(r[chave])); } return r; });
            importadas.forEach(r => { if (!usados.has(r[chave]) && !atual.some(a => a[chave] === r[chave])) { usados.add(r[chave]); saida.push(clone(r)); } });
            return saida;
        }
        const norm = r => estavel(semOrdem(r));
        const existentes = new Set(atual.map(norm));
        const porFase = !!(opcoes && opcoes.porFase);                       // Processamento: a ordem é por fase (padrão 3); Classificação: uma só
        const grupo = r => (porFase ? (r.fase || 3) : 0);
        const ultima = {};
        atual.forEach(r => { const g = grupo(r); ultima[g] = Math.max(ultima[g] === undefined ? 0 : ultima[g], Number(r.ordem) || 0); });
        const novas = [];
        importadas.slice().sort((a, b) => (Number(a.ordem) || 0) - (Number(b.ordem) || 0)).forEach(r => {
            const k = norm(r);
            if (existentes.has(k)) return;
            existentes.add(k);
            const c = clone(r);
            const g = grupo(c);
            ultima[g] = (ultima[g] === undefined ? 0 : ultima[g]) + 10;
            c.ordem = ultima[g];
            novas.push(c);
        });
        return atual.concat(novas);
    }

    /* ── DOM (só no navegador) ─────────────────────────────────────────────── */
    function baixar(nome, envelope) {
        const blob = new Blob([JSON.stringify(envelope, null, 2)], { type: 'application/json' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = nome;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    }

    function escolherArquivo() {
        return new Promise(resolve => {
            const inp = document.createElement('input');
            inp.type = 'file'; inp.accept = '.json,application/json'; inp.style.display = 'none';
            inp.addEventListener('change', () => { const f = inp.files && inp.files[0]; inp.remove(); resolve(f || null); });
            inp.addEventListener('cancel', () => { inp.remove(); resolve(null); });
            document.body.appendChild(inp);
            inp.click();
        });
    }

    function lerTexto(arquivo) {
        return new Promise((resolve, reject) => {
            if (arquivo.size > LIMITE_BYTES) { reject(new ErroArquivo([`Arquivo grande demais (máx. ${Math.round(LIMITE_BYTES / 1024 / 1024)} MB).`])); return; }
            const fr = new FileReader();
            fr.onload = () => resolve(String(fr.result || ''));
            fr.onerror = () => reject(new ErroArquivo(['Não foi possível ler o arquivo.']));
            fr.readAsText(arquivo, 'utf-8');
        });
    }

    /** Janela de confirmação com o resumo. `calcular(modo)` -> {novas, atualizadas, iguais, removidas}. Resolve 'acrescentar' | 'substituir' | null. */
    function janelaImportacao({ titulo, aviso, calcular, rotuloItem }) {
        return new Promise(resolve => {
            const fundo = document.createElement('div');
            fundo.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:10050;display:flex;align-items:center;justify-content:center;';
            const caixa = document.createElement('div');
            caixa.style.cssText = 'background:#161b22;color:#e6edf3;border:1px solid #30363d;border-radius:10px;padding:18px;width:460px;max-width:94%;font:14px/1.5 system-ui,sans-serif;';
            const h = document.createElement('h3'); h.style.cssText = 'margin:0 0 8px;font-size:1rem;'; h.textContent = titulo;
            caixa.appendChild(h);
            if (aviso) {
                const av = document.createElement('div');
                av.style.cssText = 'background:#3d2e00;border:1px solid #d29922;color:#f2cc60;border-radius:6px;padding:6px 8px;margin-bottom:8px;font-size:.85rem;';
                av.textContent = aviso; caixa.appendChild(av);
            }
            const mk = (valor, rotulo, marcado) => {
                const l = document.createElement('label'); l.style.cssText = 'display:block;margin:4px 0;cursor:pointer;';
                const r = document.createElement('input'); r.type = 'radio'; r.name = 'modoImportacao'; r.value = valor; r.checked = marcado; r.style.marginRight = '6px';
                l.append(r, document.createTextNode(rotulo)); return [l, r];
            };
            const [l1, r1] = mk('acrescentar', `Acrescentar: mantém o que já existe e junta o que for novo`, true);
            const [l2, r2] = mk('substituir', `Substituir: a lista fica igual à do arquivo`, false);
            const resumo = document.createElement('div'); resumo.style.cssText = 'margin:10px 0;padding:8px;background:#0d1117;border-radius:6px;font-size:.88rem;';
            const nota = document.createElement('div'); nota.style.cssText = 'color:#8b949e;font-size:.78rem;margin-bottom:10px;';
            nota.textContent = 'Nada é gravado agora: o resultado vai para o rascunho da tela. Só vale ao clicar em Salvar (a versão anterior fica no histórico).';
            const atualizar = () => {
                const modo = r2.checked ? 'substituir' : 'acrescentar';
                const r = calcular(modo);
                const partes = [`${r.novas} ${rotuloItem}(s) nova(s)`];
                if (r.atualizadas) partes.push(`${r.atualizadas} atualizada(s)`);
                partes.push(`${r.iguais} igual(is) (sem mudança)`);
                if (modo === 'substituir') partes.push(`${r.removidas} a remover`);
                resumo.textContent = partes.join(' · ');
            };
            r1.addEventListener('change', atualizar); r2.addEventListener('change', atualizar); atualizar();
            const bt = document.createElement('div'); bt.style.cssText = 'display:flex;justify-content:flex-end;gap:8px;';
            const fecha = v => { fundo.remove(); resolve(v); };
            const cancelar = document.createElement('button'); cancelar.type = 'button'; cancelar.textContent = 'Cancelar'; cancelar.onclick = () => fecha(null);
            const ok = document.createElement('button'); ok.type = 'button'; ok.textContent = 'Carregar no rascunho'; ok.id = 'btn-importar-aplicar';
            ok.style.cssText = 'background:#238636;color:#fff;border:none;border-radius:6px;padding:5px 12px;cursor:pointer;';
            ok.onclick = () => fecha(r2.checked ? 'substituir' : 'acrescentar');
            bt.append(cancelar, ok);
            caixa.append(l1, l2, resumo, nota, bt);
            fundo.appendChild(caixa);
            fundo.addEventListener('click', e => { if (e.target === fundo) fecha(null); });
            document.body.appendChild(fundo);
        });
    }

    const api = { LIMITE_BYTES, LIMITE_ITENS, LIMITE_REGEX, FORMATO, ErroArquivo, estavel, semOrdem, nomeDeArquivo, montarEnvelope, interpretar, resumir, aplicar,
                  baixar, escolherArquivo, lerTexto, janelaImportacao };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    if (root) root.RegrasArquivo = api;
})(typeof window !== 'undefined' ? window : undefined);
