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
            
            // Redireciona
            window.location.href = '/';
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
