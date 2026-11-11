document.getElementById('loginForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const email = document.getElementById('email').value;
    const password = document.getElementById('password').value;
    const lembrar = document.getElementById('lembrar').checked;
    const errorMsg = document.getElementById('errorMsg');
    const btnSubmit = document.getElementById('btnSubmit');

    errorMsg.classList.add('hidden');
    btnSubmit.disabled = true;
    btnSubmit.innerText = "Autenticando...";

    try {
        const res = await fetch('/api/auth/login', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, password, lembrar })
        });

        const data = await res.json();

        if (res.ok) {
            // Salva na sessão
            localStorage.setItem('auth_token', data.access_token);
            localStorage.setItem('user_id', data.user_id);
            localStorage.setItem('user_email', data.email);
            localStorage.setItem('is_admin', data.is_admin);
            
            // TASK-133: oferece salvar a senha no navegador/janela (gerenciador de senhas); falha silenciosa onde não houver suporte
            try {
                if (window.PasswordCredential && navigator.credentials && navigator.credentials.store) {
                    await navigator.credentials.store(new PasswordCredential({ id: email, password, name: email }));
                }
            } catch (_) { /* sem gerenciador de senhas: segue */ }

            // Redireciona — TASK-106: volta para a página pedida (?next=, ex.: obra aberta pelo plugin do nanoCAD); só caminho local
            const proximo = new URLSearchParams(window.location.search).get('next') || '';
            window.location.href = (proximo.startsWith('/') && !proximo.startsWith('//') && !proximo.startsWith('/\\')) ? proximo : '/';
        } else {
            errorMsg.innerText = data.detail || "Erro ao fazer login";
            errorMsg.classList.remove('hidden');
        }
    } catch (err) {
        errorMsg.innerText = "Erro de conexão com o servidor local.";
        errorMsg.classList.remove('hidden');
    } finally {
        btnSubmit.disabled = false;
        btnSubmit.innerText = "Entrar";
    }
});


// TASK-133: senhas salvas — se houver uma credencial guardada, preenche o e-mail e a senha (o usuário só clica em Entrar).
// O preenchimento automático do navegador também funciona (campos com autocomplete="username"/"current-password").
(async function preencherSenhaSalva() {
    try {
        if (!(window.PasswordCredential && navigator.credentials && navigator.credentials.get)) return;
        const email = document.getElementById('email');
        const senha = document.getElementById('password');
        if (senha.value) return;                                   // já preenchido pelo navegador
        const c = await navigator.credentials.get({ password: true, mediation: 'optional' });
        if (c && c.type === 'password' && !senha.value) {
            if (!email.value || email.value === c.id) { email.value = c.id; senha.value = c.password || ''; }
        }
    } catch (_) { /* sem suporte: o usuário digita */ }
})();
