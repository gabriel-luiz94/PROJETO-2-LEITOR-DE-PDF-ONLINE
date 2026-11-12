/**
 * marca.js — identidade do GridR (TASK-132; nome GridR e ícone na TASK-152): o nome do programa e "Desenvolvido por Gabriel Sales".
 *
 *   <div data-marca="cabecalho"></div>   ícone + "GridR" + "Leitura de projetos e orçamento de redes" (data-tamanho="compacto" = menor)
 *   rodapé discreto em todas as páginas que carregam este arquivo (<body data-rodape="fino"|"esquerda"> muda o lugar para não cobrir o conteúdo)
 */
(function () {
    const NOME = 'GridR';
    const AUTOR = 'Gabriel Sales';

    function cabecalho(el) {
        const compacto = el.dataset.tamanho === 'compacto';
        el.innerHTML = `
          <div style="white-space:nowrap; ${compacto ? 'text-align:left;' : 'text-align:center;'}">
            <div style="display:flex; align-items:center; gap:10px; ${compacto ? '' : 'justify-content:center;'}">
              <img src="/static/gridr.svg" alt="" width="${compacto ? 34 : 44}" height="${compacto ? 34 : 44}" style="display:block;">
              <div style="font-size:${compacto ? '1.6rem' : '2rem'}; font-weight:800; line-height:1.05; color:#e6edf3;">Grid<span style="color:#d97706;">R</span></div>
            </div>
            <div style="font-size:0.75rem; color:#8b949e; letter-spacing:0.03em;">Leitura de projetos e orçamento de redes</div>
          </div>`;
    }

    function rodape() {
        if (document.getElementById('marca-rodape') || document.body.dataset.semRodape === '1') return;
        const f = document.createElement('div');
        f.id = 'marca-rodape';
        // TASK-153: o rodapé não fica mais por cima do conteúdo — <body data-rodape="fino"> (telas que ocupam a altura toda: cabe na margem de baixo)
        // e <body data-rodape="esquerda"> (telas com um painel fixo no canto direito). A cor #8b949e dá 6:1 de contraste (a anterior dava 4,1:1).
        const modo = document.body.dataset.rodape || '';
        f.style.cssText = 'position:fixed; font-size:0.7rem; color:#8b949e; z-index:5; pointer-events:none; ' +
                          (modo === 'esquerda' ? 'left:12px; ' : 'right:12px; ') +
                          (modo === 'fino' ? 'bottom:1px; line-height:16px; font-size:0.66rem; padding:0 4px;'
                                           : 'bottom:8px; background:rgba(13,17,23,0.75); padding:2px 8px; border-radius:6px;');
        f.textContent = `${NOME} · Desenvolvido por ${AUTOR}`;
        document.body.appendChild(f);
    }

    function iniciar() {
        rodape();
        document.querySelectorAll('[data-marca="cabecalho"]').forEach(cabecalho);
    }

    window.GridR = window.GridReader = { NOME, AUTOR };          // GridReader = nome antigo (compatibilidade)
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
    else iniciar();
})();
