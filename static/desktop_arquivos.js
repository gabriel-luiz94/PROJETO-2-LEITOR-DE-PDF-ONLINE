/*
 * desktop_arquivos.js (TASK-142) — no aplicativo desktop (janela nativa/pywebview), baixar e importar arquivos pelas caixas NATIVAS do Windows.
 *   • download (`<a download>` com blob:/data:) e `showSaveFilePicker` → "Salvar como" nativo (api.salvar_arquivo)
 *   • `<input type="file">` (clique do usuário ou `.click()` do código) → "Abrir" nativo (api.abrir_arquivos); os bytes viram `File`s no input e o `change` dispara
 * No navegador comum (versão online) a ponte não existe e NADA aqui muda o comportamento. Parte pura (base64) exposta em `DesktopArquivos` e testada em Node.
 */
(function (raiz) {
    'use strict';

    function paraBase64(buf) {
        const bytes = new Uint8Array(buf), bloco = 0x8000;
        let s = '';
        for (let i = 0; i < bytes.length; i += bloco) s += String.fromCharCode.apply(null, bytes.subarray(i, i + bloco));
        return raiz.btoa(s);
    }
    function deBase64(b64) {
        const s = raiz.atob(b64), out = new Uint8Array(s.length);
        for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
        return out;
    }
    const api = { paraBase64, deBase64 };
    if (typeof document === 'undefined') { if (typeof module !== 'undefined' && module.exports) module.exports = api; return; }
    if (raiz.__desktopArquivosLigado) return;
    raiz.__desktopArquivosLigado = true;
    raiz.DesktopArquivos = api;

    const doc = raiz.document;
    /** A ponte do pywebview: na própria janela ou, se esta página estiver numa moldura (iframe) do programa, na janela de cima. */
    function ponte() {
        for (const w of [raiz, raiz.parent, raiz.top]) {
            try { if (w && w.pywebview && w.pywebview.api && w.pywebview.api.salvar_arquivo) return w.pywebview.api; } catch (e) { /* outra origem */ }
        }
        return null;
    }
    const avisar = m => { try { if (raiz.showToast) raiz.showToast(m); } catch (e) { /* sem toast */ } };

    async function salvarBlob(blob, nome) {
        const p = ponte();
        const r = await p.salvar_arquivo(nome || 'arquivo', paraBase64(await blob.arrayBuffer()), '');
        if (r && r.ok) avisar('Arquivo salvo em: ' + r.caminho);
        else if (r && !r.cancelado) avisar('Não foi possível salvar o arquivo' + (r.erro ? ': ' + r.erro : '.'));
        return r || {};
    }
    async function salvarUrl(href, nome) {
        const blob = await (await raiz.fetch(href)).blob();
        return salvarBlob(blob, nome);
    }
    const baixavel = a => a && a.hasAttribute && a.hasAttribute('download') && /^(blob:|data:)/i.test(a.href || '');

    /* ── downloads ── */
    const cliqueOriginalA = raiz.HTMLAnchorElement.prototype.click;
    raiz.HTMLAnchorElement.prototype.click = function () {
        if (ponte() && baixavel(this)) { salvarUrl(this.href, this.getAttribute('download') || 'arquivo').catch(e => avisar('Erro ao salvar: ' + e.message)); return; }
        return cliqueOriginalA.apply(this, arguments);
    };
    doc.addEventListener('click', e => {
        const a = e.target && e.target.closest ? e.target.closest('a[download]') : null;
        if (a && ponte() && baixavel(a)) { e.preventDefault(); salvarUrl(a.href, a.getAttribute('download') || 'arquivo').catch(err => avisar('Erro ao salvar: ' + err.message)); }
    }, true);

    /* ── showSaveFilePicker (quem usa: Exportar Excel) ── */
    function instalarSalvarComo() {
        if (!ponte() || raiz.__salvarComoDesktop) return;
        raiz.__salvarComoDesktop = true;
        raiz.showSaveFilePicker = async function (opcoes) {
            const nome = (opcoes && opcoes.suggestedName) || 'arquivo';
            return {
                name: nome,
                createWritable: async () => {
                    const partes = [];
                    return {
                        write: async d => { partes.push(d instanceof Blob ? d : new Blob([d])); },
                        close: async () => {
                            const r = await salvarBlob(new Blob(partes), nome);
                            if (r.cancelado) { const e = new Error('Salvar cancelado'); e.name = 'AbortError'; throw e; }
                        }
                    };
                }
            };
        };
    }

    /* ── escolher arquivo ── */
    async function escolher(input) {
        const p = ponte();
        const r = await p.abrir_arquivos(input.getAttribute('accept') || '', !!input.multiple);
        if (!r || r.cancelado) return;
        if (!r.ok) { avisar('Não foi possível abrir o arquivo' + (r.erro ? ': ' + r.erro : '.')); return; }
        const dt = new DataTransfer();
        r.arquivos.forEach(a => dt.items.add(new File([deBase64(a.base64)], a.nome)));
        input.files = dt.files;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    }
    const cliqueOriginalInput = raiz.HTMLInputElement.prototype.click;
    raiz.HTMLInputElement.prototype.click = function () {
        if (this.type === 'file' && ponte()) { escolher(this).catch(e => avisar('Erro ao abrir: ' + e.message)); return; }
        return cliqueOriginalInput.apply(this, arguments);
    };
    doc.addEventListener('click', e => {
        const i = e.target;
        if (i && i.tagName === 'INPUT' && i.type === 'file' && ponte()) { e.preventDefault(); escolher(i).catch(err => avisar('Erro ao abrir: ' + err.message)); }
    }, true);

    // a ponte do pywebview chega um instante depois da página
    instalarSalvarComo();
    raiz.addEventListener('pywebviewready', instalarSalvarComo);
    let tentativas = 0;
    const espera = raiz.setInterval(() => { instalarSalvarComo(); if (raiz.__salvarComoDesktop || ++tentativas > 40) raiz.clearInterval(espera); }, 250);
})(typeof window !== 'undefined' ? window : globalThis);
