/*
 * campos_segredo.js (TASK-128) — campos de chave/segredo que NÃO são login (chaves de API, senha de usuário novo no admin).
 * `<input class="campo-segredo">` vira um campo de texto mascarado (-webkit-text-security): o navegador não o trata como "senha" e para de
 * colar o e-mail do login nele e de oferecer "salvar a senha". Onde o navegador não mascara por CSS (Firefox, versão online), volta a ser type="password"
 * com autocomplete="new-password" (não preenche credenciais salvas).
 */
(function () {
    'use strict';
    function aplicar() {
        const mascara = !!(window.CSS && CSS.supports && (CSS.supports('-webkit-text-security', 'disc') || CSS.supports('text-security', 'disc')));
        document.querySelectorAll('input.campo-segredo').forEach(inp => {
            inp.setAttribute('autocomplete', mascara ? 'off' : 'new-password');
            inp.setAttribute('data-lpignore', 'true');
            inp.setAttribute('data-1p-ignore', 'true');
            inp.setAttribute('data-form-type', 'other');
            if (mascara) { inp.type = 'text'; inp.style.webkitTextSecurity = 'disc'; inp.style.textSecurity = 'disc'; }
            else inp.type = 'password';
        });
    }
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', aplicar); else aplicar();
})();
