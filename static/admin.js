document.addEventListener('DOMContentLoaded', () => {
    const token = localStorage.getItem('auth_token');
    if (!token) {
        window.location.href = '/login';
        return;
    }

    const is_admin = localStorage.getItem('is_admin') === 'true';
    if (!is_admin) {
        alert("Acesso negado. Apenas administradores podem ver o painel.");
        window.location.href = '/';
        return;
    }

    loadUsers();

    document.getElementById('addUserForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const email = document.getElementById('newEmail').value;
        const password = document.getElementById('newPassword').value;
        const role = document.getElementById('newRole').value;
        
        try {
            const res = await fetch(`/api/admin/users?token=${token}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email, password, role: role })
            });
            const data = await res.json();
            
            if (res.ok) {
                showMessage('success', data.msg);
                document.getElementById('addUserForm').reset();
                loadUsers();
            } else {
                showMessage('error', data.detail || 'Erro ao criar usuário');
            }
        } catch (err) {
            showMessage('error', 'Erro de conexão com o servidor.');
        }
    });
});

async function loadUsers() {
    const token = localStorage.getItem('auth_token');
    const tbody = document.getElementById('userTableBody');
    tbody.innerHTML = '<tr><td colspan="3" class="px-6 py-4 text-center">Carregando...</td></tr>';
    
    try {
        const res = await fetch(`/api/admin/users?token=${token}`);
        if (!res.ok) {
            if (res.status === 403) {
                window.location.href = '/';
                return;
            }
            throw new Error('Falha ao carregar');
        }
        const users = await res.json();
        
        tbody.innerHTML = '';
        users.forEach(user => {
            const tr = document.createElement('tr');
            
            const tdEmail = document.createElement('td');
            tdEmail.className = 'px-6 py-4';
            tdEmail.textContent = user.email;
            
            const tdType = document.createElement('td');
            tdType.className = 'px-6 py-4';
            const badge = document.createElement('span');
            badge.className = user.role === 'admin' ? 'px-2 py-1 bg-[#d97706] text-black text-xs font-bold rounded' : 'px-2 py-1 bg-[#30363d] text-gray-300 text-xs rounded';
            badge.textContent = user.role === 'admin' ? 'Admin' : 'Operador';
            tdType.appendChild(badge);
            
            const tdAction = document.createElement('td');
            tdAction.className = 'px-6 py-4 text-right space-x-2';
            
            const btnToggle = document.createElement('button');
            btnToggle.className = 'text-sm text-blue-400 hover:text-blue-300 transition-colors';
            btnToggle.textContent = user.role === 'admin' ? 'Rebaixar p/ Operador' : 'Promover p/ Admin';
            btnToggle.onclick = () => updateRole(user.id, user.role === 'admin' ? 'operador' : 'admin');
            
            const btnReset = document.createElement('button');
            btnReset.className = 'text-sm text-yellow-500 hover:text-yellow-400 transition-colors ml-4';
            btnReset.textContent = 'Mudar Senha';
            btnReset.onclick = () => resetPassword(user.id, user.email);
            
            const btnDelete = document.createElement('button');
            btnDelete.className = 'text-sm text-red-400 hover:text-red-300 transition-colors ml-4';
            btnDelete.textContent = 'Excluir';
            btnDelete.onclick = () => deleteUser(user.id, user.email);
            
            tdAction.appendChild(btnToggle);
            tdAction.appendChild(btnReset);
            tdAction.appendChild(btnDelete);
            
            tr.appendChild(tdEmail);
            tr.appendChild(tdType);
            tr.appendChild(tdAction);
            tbody.appendChild(tr);
        });
        
        if (users.length === 0) {
            tbody.innerHTML = '<tr><td colspan="3" class="px-6 py-4 text-center">Nenhum usuário local encontrado.</td></tr>';
        }
    } catch (err) {
        tbody.innerHTML = '<tr><td colspan="3" class="px-6 py-4 text-center text-red-500">Erro ao carregar usuários locais.</td></tr>';
    }
}

async function updateRole(userId, newRole) {
    const token = localStorage.getItem('auth_token');
    try {
        const res = await fetch(`/api/admin/users/${userId}/role?token=${token}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ role: newRole })
        });
        if (res.ok) {
            loadUsers();
        } else {
            const data = await res.json();
            showMessage('error', data.detail || 'Erro ao alterar privilégios');
        }
    } catch (err) {
        showMessage('error', 'Erro de conexão.');
    }
}

async function resetPassword(userId, email) {
    const newPassword = prompt(`Digite a nova senha para ${email}: (Mín. 6 caracteres)`);
    if (!newPassword) return; // Cancelou ou deixou em branco
    
    if (newPassword.length < 6) {
        showMessage('error', 'A senha precisa ter no mínimo 6 caracteres.');
        return;
    }
    
    const token = localStorage.getItem('auth_token');
    try {
        const res = await fetch(`/api/admin/users/${userId}/password?token=${token}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ password: newPassword })
        });
        
        if (res.ok) {
            showMessage('success', `Senha de ${email} redefinida com sucesso!`);
        } else {
            const data = await res.json();
            showMessage('error', data.detail || 'Erro ao redefinir senha');
        }
    } catch (err) {
        showMessage('error', 'Erro de conexão.');
    }
}

async function deleteUser(userId, email) {
    if (!confirm(`Tem certeza que deseja excluir o usuário ${email}?`)) return;
    
    const token = localStorage.getItem('auth_token');
    try {
        const res = await fetch(`/api/admin/users/${userId}?token=${token}`, {
            method: 'DELETE'
        });
        if (res.ok) {
            loadUsers();
        } else {
            const data = await res.json();
            showMessage('error', data.detail || 'Erro ao excluir usuário');
        }
    } catch (err) {
        showMessage('error', 'Erro de conexão.');
    }
}

function showMessage(type, msg) {
    const errorMsg = document.getElementById('errorMsg');
    const successMsg = document.getElementById('successMsg');
    
    errorMsg.classList.add('hidden');
    successMsg.classList.add('hidden');
    
    if (type === 'error') {
        errorMsg.textContent = msg;
        errorMsg.classList.remove('hidden');
    } else {
        successMsg.textContent = msg;
        successMsg.classList.remove('hidden');
    }
    
    setTimeout(() => {
        errorMsg.classList.add('hidden');
        successMsg.classList.add('hidden');
    }, 5000);
}
