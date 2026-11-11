/**
 * plugin_nanocad.js — TASK-131: o programa principal traz o plugin do nanoCAD.
 * Botão "Plugin nanoCAD" no topo → janela com a versão que o programa traz × a instalada, o caminho para o APPLOAD,
 * "Instalar / atualizar agora" e "Abrir pasta" (programa instalado) ou os links para baixar a DLL (online).
 */
(function () {
    const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

    function modal() {
        let m = document.getElementById('modal-plugin-nanocad');
        if (m) return m;
        m = document.createElement('div');
        m.id = 'modal-plugin-nanocad';
        m.className = 'modal-overlay';
        m.style.cssText = 'display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); z-index:1000; align-items:center; justify-content:center;';
        m.innerHTML = `
          <div style="background:#161b22; border:1px solid #30363d; border-radius:10px; width:min(720px,94vw); max-height:88vh; overflow:auto; padding:20px; color:#e6edf3;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
              <h3 style="margin:0;">Plugin do nanoCAD</h3>
              <button class="btn-secondary" style="padding:6px 14px; font-size:0.85rem;" onclick="document.getElementById('modal-plugin-nanocad').style.display='none'">Fechar</button>
            </div>
            <div id="plugin-nanocad-corpo" style="font-size:0.9rem; color:#c9d1d9;">Carregando…</div>
          </div>`;
        m.addEventListener('click', e => { if (e.target === m) m.style.display = 'none'; });
        document.body.appendChild(m);
        return m;
    }

    function linhaAlvo(a) {
        const situ = !a.instalado ? '<span style="color:#d29922;">não instalado</span>'
            : a.desatualizado
                ? `<span style="color:#d29922;">instalado ${esc(a.versao_instalada || '?')} (o programa traz ${esc(a.versao_programa)})</span>`
                : `<span style="color:#3fb950;">instalado ${esc(a.versao_instalada || '?')}</span>`;
        return `<div style="border:1px solid #30363d; border-radius:8px; padding:10px 12px; margin:8px 0; background:#0d1117;">
            <div style="display:flex; justify-content:space-between; gap:8px; flex-wrap:wrap;"><b>${esc(a.nome)}</b>${situ}</div>
            ${a.versao_programa ? '' : '<div style="color:#8b949e; font-size:0.8rem;">Este programa não traz a DLL para este nanoCAD.</div>'}
            <div style="margin-top:6px; display:flex; gap:6px; align-items:center;">
              <code style="flex:1; background:#161b22; padding:4px 6px; border-radius:4px; font-size:0.78rem; overflow-x:auto; white-space:nowrap;">${esc(a.caminho)}</code>
              <button class="btn-secondary" style="padding:3px 9px; font-size:0.75rem;" data-copiar="${esc(a.caminho)}">Copiar</button>
            </div></div>`;
    }

    async function carregar() {
        const corpo = document.getElementById('plugin-nanocad-corpo');
        corpo.textContent = 'Carregando…';
        let d;
        try {
            const r = await fetch('/api/plugin/instalacao');
            d = await r.json();
            if (!r.ok) throw new Error(d.detail || r.status);
        } catch (e) { corpo.textContent = 'Não foi possível consultar o plugin: ' + e.message; return; }
        let h = `<p style="margin-top:0;">O plugin lê o desenho no nanoCAD com as regras e a configuração de leitura de cada projeto cadastrado aqui,
                 confere com a Base de Orçamento e abre a obra nesta aplicação. Depois de instalado, ele <b>se atualiza sozinho</b>
                 (ORC_CONECTAR → "Atualizar o plugin sozinho").</p>`;
        if (d.desktop) {
            h += (d.alvos || []).map(linhaAlvo).join('');
            h += `<div style="display:flex; gap:8px; margin:10px 0; flex-wrap:wrap;">
                    <button class="btn-primary" id="plugin-nanocad-instalar" style="padding:6px 14px; font-size:0.85rem;">Instalar / atualizar agora</button>
                    <button class="btn-secondary" id="plugin-nanocad-pasta" style="padding:6px 14px; font-size:0.85rem;">Abrir pasta</button>
                  </div><div id="plugin-nanocad-msg" style="font-size:0.85rem; color:#8b949e;"></div>`;
        }
        if ((d.downloads || []).length) {
            h += `<p style="margin-bottom:4px;">${d.desktop ? 'Ou baixe a DLL' : 'Baixe a DLL do seu nanoCAD'}:</p><ul style="margin-top:0;">`
               + d.downloads.map(x => `<li><a href="${esc(x.url)}" style="color:#58a6ff;">${esc(x.nome)} — versão ${esc(x.versao)}</a></li>`).join('') + '</ul>';
        }
        h += `<p style="margin-bottom:4px;"><b>No nanoCAD (uma vez):</b></p>
              <ol style="margin-top:0; padding-left:20px; line-height:1.6;">
                <li>Comando <code>APPLOAD</code> → em <b>Inicialização</b> (Startup Suite) → <b>Adicionar</b> → escolha o arquivo do seu nanoCAD (caminho acima).</li>
                <li>Feche e abra o nanoCAD. Comando <code>ORCAMENTO</code>.</li>
                <li><code>ORC_CONECTAR</code> com o endereço desta aplicação e o seu login: a tela inicial passa a listar todos os projetos.</li>
              </ol>
              <p style="font-size:0.8rem; color:#8b949e;">nanoCAD 5: a DLL desliga sozinha o que a versão 5 não oferece (ex.: números desenhados na tela, MLEADER) e segue funcionando.
              Se não carregar, use <code>plugin_nanocad\\build_nanocad5.bat</code> no computador com o nanoCAD 5.</p>`;
        corpo.innerHTML = h;
        corpo.querySelectorAll('[data-copiar]').forEach(b => b.onclick = () => { navigator.clipboard && navigator.clipboard.writeText(b.dataset.copiar); b.textContent = 'Copiado'; });
        const bi = document.getElementById('plugin-nanocad-instalar');
        if (bi) bi.onclick = async () => {
            bi.disabled = true; bi.textContent = 'Instalando…';
            try {
                const r = await fetch('/api/plugin/instalacao', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ forcar: false }) });
                const j = await r.json();
                if (!r.ok) throw new Error(j.detail || r.status);
                const txt = j.resultado.map(x => `${x.nome}: ${x.resultado}${x.versao ? ' ' + x.versao : ''}${x.detalhe ? ' (' + x.detalhe + ')' : ''}`).join(' · ');
                await carregar();
                document.getElementById('plugin-nanocad-msg').textContent = txt + ' — com o nanoCAD aberto, vale ao reabrir.';
            } catch (e) { alert('Instalação falhou: ' + e.message); bi.disabled = false; bi.textContent = 'Instalar / atualizar agora'; }
        };
        const bp = document.getElementById('plugin-nanocad-pasta');
        if (bp) bp.onclick = async () => {
            const r = await fetch('/api/plugin/instalacao/abrir-pasta', { method: 'POST' });
            const j = await r.json().catch(() => ({}));
            if (!j.aberta) document.getElementById('plugin-nanocad-msg').textContent = 'Pasta: ' + (j.pasta || '');
        };
    }

    window.abrirModalPluginNanoCAD = function () {
        const m = modal();
        m.style.display = 'flex';
        carregar();
    };
})();
