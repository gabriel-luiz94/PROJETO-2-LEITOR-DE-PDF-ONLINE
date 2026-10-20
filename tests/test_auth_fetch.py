"""TASK-076 — static/auth_fetch.js: 401 de /api/* limpa a sessão e vai ao login (uma vez); outros status e o login não."""
import json
import os
import subprocess

import pytest

from tests.oraculo_tela import NODE

pytestmark = pytest.mark.skipif(not NODE, reason="node indisponível")
JS = os.path.join(os.path.dirname(__file__), "..", "static", "auth_fetch.js")

ROTEIRO = r"""
const fs = require('fs');
const store = {auth_token: 'tok', user_id: '1', user_email: 'a@b', is_admin: 'true'};
const out = [];
global.localStorage = {getItem: k => (k in store ? store[k] : null), removeItem: k => { delete store[k]; }};
global.alert = () => {};
const status = {};
global.window = global; window.location = {href: '/'};
window.fetch = async (u, o) => ({status: status[u] || 200});
eval(fs.readFileSync(process.argv[1], 'utf8'));
(async () => {
  await window.fetch('/api/x'); out.push(['ok', window.location.href, !!store.auth_token]);
  status['/api/proibido'] = 403; await window.fetch('/api/proibido'); out.push(['403', window.location.href, !!store.auth_token]);
  status['/api/auth/login'] = 401; await window.fetch('/api/auth/login'); out.push(['login', window.location.href, !!store.auth_token]);
  status['/api/y'] = 401; await window.fetch('/api/y'); out.push(['401', window.location.href, !!store.auth_token]);
  window.location.href = '/'; await window.fetch('/api/y'); out.push(['2a vez', window.location.href]);
  console.log(JSON.stringify(out));
})();
"""


def test_401_redireciona_ao_login_so_uma_vez():
    r = subprocess.run([NODE, "-e", ROTEIRO, JS], capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    ok, c403, login, c401, segunda = json.loads(r.stdout)
    assert ok == ["ok", "/", True]
    assert c403 == ["403", "/", True]
    assert login == ["login", "/", True]
    assert c401 == ["401", "/login", False]
    assert segunda == ["2a vez", "/"]
