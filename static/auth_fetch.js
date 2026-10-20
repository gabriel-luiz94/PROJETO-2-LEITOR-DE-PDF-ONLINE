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
        const ehApi = typeof url === 'string' && url.includes('/api/');
        const p = origFetch(url, options);
        if (!ehApi || url.includes('/api/auth/login')) return p;
        // TASK-076: 401 = sessão expirada/inválida → volta ao login (uma vez só), em vez de falhar em silêncio.
        return p.then(function(res) {
            if (res && res.status === 401 && token && !window.__sessaoExpiradaTratada) {
                window.__sessaoExpiradaTratada = true;
                ['auth_token', 'user_id', 'user_email', 'is_admin'].forEach(function(k) { localStorage.removeItem(k); });
                try { alert('Sessão expirada. Entre novamente.'); } catch (e) {}
                window.location.href = '/login';
            }
            return res;
        });
    };
})();
