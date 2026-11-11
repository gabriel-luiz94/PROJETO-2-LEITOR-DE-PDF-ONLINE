/**
 * marca.js — identidade do ReadGrid (TASK-132): nome, logo da Engeselt Engenharia e "Desenvolvido por Gabriel Sales".
 *
 *   <div data-marca="cabecalho"></div>   logo da Engeselt (se o arquivo existir) + "ReadGrid"
 *   rodapé discreto em todas as páginas que carregam este arquivo
 *
 * Logo: coloque o arquivo em static/img/logo_engeselt.png (ou .svg). Sem o arquivo, aparece "Engeselt Engenharia" em texto.
 */
(function () {
    const NOME = 'ReadGrid';
    const EMPRESA = 'Engeselt Engenharia';
    const AUTOR = 'Gabriel Sales';
    const LOGOS = ['/static/img/logo_engeselt.svg', '/static/img/logo_engeselt.png'];

    // tenta cada arquivo de logo; devolve a URL do primeiro que carregar (ou null)
    function acharLogo() {
        return new Promise(resolve => {
            let i = 0;
            const tentar = () => {
                if (i >= LOGOS.length) return resolve(null);
                const img = new Image();
                const url = LOGOS[i++];
                img.onload = () => resolve(url);
                img.onerror = tentar;
                img.src = url + '?v=1';
            };
            tentar();
        });
    }

    function logoOuTexto(url, altura) {
        if (url) return `<img src="${url}?v=1" alt="${EMPRESA}" title="${EMPRESA}" style="height:${altura}px; width:auto; display:block;">`;
        return `<span style="font-weight:700; letter-spacing:0.02em; color:#c9d1d9; white-space:nowrap;">${EMPRESA}</span>`;
    }

    function cabecalho(el, url) {
        const compacto = el.dataset.tamanho === 'compacto';
        el.innerHTML = `
          <div style="display:flex; align-items:center; gap:${compacto ? 10 : 14}px; white-space:nowrap; ${compacto ? '' : 'justify-content:center;'}">
            <div style="display:flex; align-items:center;">${logoOuTexto(url, compacto ? 30 : 44)}</div>
            <div style="width:1px; align-self:stretch; background:#30363d;"></div>
            <div style="text-align:left;">
              <div style="font-size:${compacto ? '1.35rem' : '1.7rem'}; font-weight:800; line-height:1.05; color:#e6edf3;">Read<span style="color:#d97706;">Grid</span></div>
              <div style="font-size:0.72rem; color:#8b949e; letter-spacing:0.03em;">Leitura de projetos e orçamento de redes</div>
            </div>
          </div>`;
    }

    function rodape() {
        if (document.getElementById('marca-rodape') || document.body.dataset.semRodape === '1') return;
        const f = document.createElement('div');
        f.id = 'marca-rodape';
        f.style.cssText = 'position:fixed; right:12px; bottom:8px; font-size:0.7rem; color:#6e7681; z-index:5; pointer-events:none; ' +
                          'background:rgba(13,17,23,0.75); padding:2px 8px; border-radius:6px;';
        f.textContent = `${NOME} · ${EMPRESA} · Desenvolvido por ${AUTOR}`;
        document.body.appendChild(f);
    }

    function iniciar() {
        rodape();
        const alvos = document.querySelectorAll('[data-marca="cabecalho"]');
        if (!alvos.length) return;
        acharLogo().then(url => alvos.forEach(el => cabecalho(el, url)));
    }

    window.ReadGrid = { NOME, EMPRESA, AUTOR };
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
    else iniciar();
})();
