/**
 * auth_fetch.js — Interceptor global de requisições fetch para rotas /api/*.
 * Anexa automaticamente o cabeçalho 'Authorization: Bearer <token>' se o usuário estiver logado.
 */
(function() {
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
        return origFetch(url, options);
    };
})();
