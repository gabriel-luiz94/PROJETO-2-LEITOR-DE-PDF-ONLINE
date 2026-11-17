/**
 * auth_fetch.js — Interceptor global de requisições fetch para rotas /api/*.
 * Anexa automaticamente o cabeçalho 'Authorization: Bearer <token>' se o usuário estiver logado.
 */
(function() {
    // "Lembrar de mim" desmarcado no login: a sessão vale só enquanto o programa está aberto. Numa abertura nova (sessionStorage vazio) o token é descartado e o login volta.
    // Sem a marca (sessão criada antes desta versão) ou com ela marcada, entra direto enquanto o token valer.
    try {
        var naTelaDeLogin = (window.location.pathname || '').indexOf('/login') === 0;
        if (!naTelaDeLogin && localStorage.getItem('auth_token') && localStorage.getItem('auth_lembrar') === 'false' && !sessionStorage.getItem('sessao_ativa')) {
            ['auth_token', 'user_id', 'user_email', 'is_admin', 'auth_lembrar'].forEach(function(k) { localStorage.removeItem(k); });
            var aqui0 = (window.location.pathname || '') + (window.location.search || '');
            window.location.replace(aqui0 && aqui0 !== '/' ? '/login?next=' + encodeURIComponent(aqui0) : '/login');
        }
    } catch (e) { /* sem storage: segue como antes */ }
    const origFetch = window.fetch;
    window.fetch = function(url, options) {
        options = options || {};
        const token = localStorage.getItem('auth_token');
        if (token && typeof url === 'string' && url.includes('/api/')) {
            options.headers = options.headers || {};
            if (options.headers instanceof Headers) {
                if (!options.headers.has('Authorization')) {
                    options.headers.append('Authorization', 'Bearer ' + token);
                }
            } else if (Array.isArray(options.headers)) {
                let hasAuth = false;
                for (let i = 0; i < options.headers.length; i++) {
                    if (options.headers[i][0].toLowerCase() === 'authorization') { hasAuth = true; break; }
                }
                if (!hasAuth) options.headers.push(['Authorization', 'Bearer ' + token]);
            } else {
                let hasAuth = false;
                for (let k in options.headers) {
                    if (k.toLowerCase() === 'authorization') { hasAuth = true; break; }
                }
                if (!hasAuth) options.headers['Authorization'] = 'Bearer ' + token;
            }
        }
        const ehApi = typeof url === 'string' && url.includes('/api/');
        const p = origFetch(url, options);
        if (!ehApi || url.includes('/api/auth/login')) return p;
        // TASK-076: 401 = sessão expirada/inválida → volta ao login (uma vez só), em vez de falhar em silêncio.
        return p.then(function(res) {
            if (res && res.status === 401 && token && !window.__sessaoExpiradaTratada) {
                window.__sessaoExpiradaTratada = true;
                ['auth_token', 'user_id', 'user_email', 'is_admin', 'auth_lembrar'].forEach(function(k) { localStorage.removeItem(k); });
                try { alert('Sessão expirada. Entre novamente.'); } catch (e) {}
                // TASK-106: volta para onde estava depois do login (ex.: obra aberta pelo plugin do nanoCAD)
                var aqui = (window.location.pathname || '') + (window.location.search || '');
                window.location.href = aqui && aqui !== '/' && aqui.indexOf('/login') !== 0 ? '/login?next=' + encodeURIComponent(aqui) : '/login';
            }
            return res;
        });
    };
})();
