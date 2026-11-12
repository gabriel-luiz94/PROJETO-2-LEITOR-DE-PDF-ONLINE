/**
 * marca.js — identidade do GridR (TASK-132; nome GridR e ícone na TASK-152): o nome do programa e "Desenvolvido por Gabriel Sales".
 *
 *   <div data-marca="cabecalho"></div>   ícone + "GridR" + "Leitura de projetos e orçamento de redes" (data-tamanho="compacto" = menor)
 *   rodapé discreto em todas as páginas que carregam este arquivo
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
        f.style.cssText = 'position:fixed; right:12px; bottom:8px; font-size:0.7rem; color:#6e7681; z-index:5; pointer-events:none; ' +
                          'background:rgba(13,17,23,0.75); padding:2px 8px; border-radius:6px;';
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
