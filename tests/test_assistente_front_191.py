"""TASK-191 — lado do navegador do assistente (static/assistente.js): stream NDJSON, cópia das tabelas, aplicação do plano com checagem; e a ligação com o Resumo."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
ARQ = os.path.join(config.STATIC_DIR, "assistente.js")
pytestmark = pytest.mark.skipif(not NODE, reason="node indisponível")


def _ler(nome):
    with open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8") as f:
        return f.read()


def _node(js, entrada=None):
    r = subprocess.run([NODE, "-e", "const A = require(%s); const e = JSON.parse(require('fs').readFileSync(0, 'utf8')); %s" % (json.dumps(ARQ), js)],
                       input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def test_sintaxe_e_ligacao_no_resumo():
    assert subprocess.run([NODE, "--check", ARQ], capture_output=True, text=True).returncode == 0
    html, js = _ler("resumo.html"), _ler("resumo.js")
    assert html.index("/static/assistente.js") < html.index("/static/resumo.js")
    assert "btn-assistente-skills" in html
    assert "/api/assistente/chat" in js and "/api/gemini/chat" not in js
    assert "Assistente.linhasNdjson" in js and "Assistente.aplicarPlano" in js and "Assistente.renderPlano" in js


def test_linhas_ndjson_junta_pedacos_e_ignora_lixo():
    r = _node("""
        const a = A.linhasNdjson('{"tipo":"texto","texto":"oi"}\\n{"tipo":"fer');
        const b = A.linhasNdjson(a.resto + 'ramenta","nome":"x"}\\nlixo\\n\\n[1,2]\\n{"tipo":"fim"}\\n');
        console.log(JSON.stringify({ a, b }));""", None)
    assert r["a"]["eventos"] == [{"tipo": "texto", "texto": "oi"}] and r["a"]["resto"] == '{"tipo":"fer'
    assert [e["tipo"] for e in r["b"]["eventos"]] == ["ferramenta", "fim"] and r["b"]["resto"] == ""


def test_tabela_para_servidor_guarda_a_posicao_real_e_so_os_campos_uteis():
    estados = {"cabos": {"data": [{"entidade": "CABO", "operacao": "I", "ativo": "CAA2 3 35 m", "_x": 1, "_y": 2, "vinculoEstruturas": [1]}]},
               "outros": {"data": [None, {"entidade": "0", "operacao": "R", "ativo": "1-CFU", "congelada": True, "grande": "x" * 50}, {"operacao": "I", "ativo": "11300D"}]}}
    r = _node("console.log(JSON.stringify(A.tabelaParaServidor(e)))", estados)
    assert r["cabos"] == [{"id": 0, "entidade": "CABO", "operacao": "I", "ativo": "CAA2 3 35 m", "vinculoEstruturas": [1]}]
    assert [x["id"] for x in r["outros"]] == [1, 2] and r["outros"][0]["congelada"] is True and "grande" not in r["outros"][0]


def _aplicar(estados, itens):
    return _node("""
        const est = e.estados;
        const r = A.aplicarPlano(est, e.itens, (a, t) => t === 'cabos' ? 'CABO' : (/^DT/.test(a) ? 'POSTE' : '0'));
        console.log(JSON.stringify({ r, est, resumo: A.resumoAplicacao(r) }));""", {"estados": estados, "itens": itens})


def test_aplicar_plano_adiciona_edita_remove_e_filtra_nulos():
    est = {"cabos": {"data": [{"entidade": "CABO", "operacao": "I", "ativo": "CAA2 3 35 m"}]},
           "outros": {"data": [{"entidade": "0", "operacao": "I", "ativo": "1-CFU"}, {"entidade": "0", "operacao": "I", "ativo": "1-PR220"}, {"entidade": "0", "operacao": "I", "ativo": "1-X"}]}}
    itens = [
        {"n": 1, "acao": "adicionar", "tabela": "outros", "operacao": "I", "ativo": "DT11/300 1-TR125"},
        {"n": 2, "acao": "editar", "tabela": "outros", "id": 0, "operacao": "R", "ativo": "2-CFU", "antes": {"operacao": "I", "ativo": "1-CFU"}},
        {"n": 3, "acao": "remover", "tabela": "outros", "id": 1, "antes": {"operacao": "I", "ativo": "1-PR220"}},
        {"n": 4, "acao": "adicionar", "tabela": "cabos", "operacao": "*I", "ativo": "CAA2 3 10 m"},
    ]
    r = _aplicar(est, itens)
    assert (r["r"]["adicionadas"], r["r"]["editadas"], r["r"]["removidas"], r["r"]["ignoradas"]) == (2, 1, 1, [])
    assert [x["ativo"] for x in r["est"]["outros"]["data"]] == ["2-CFU", "1-X", "DT11/300 1-TR125"]
    assert r["est"]["outros"]["data"][0]["operacao"] == "R" and r["est"]["outros"]["data"][2]["entidade"] == "POSTE"
    assert r["est"]["cabos"]["data"][1] == {"entidade": "CABO", "operacao": "*I", "ativo": "CAA2 3 10 m"}
    assert r["resumo"] == "2 adicionada(s), 1 editada(s), 1 removida(s)"


def test_aplicar_plano_ignora_linha_que_mudou_ou_sumiu():
    est = {"cabos": {"data": []}, "outros": {"data": [{"entidade": "0", "operacao": "I", "ativo": "3-CFU"}]}}
    itens = [{"n": 1, "acao": "editar", "tabela": "outros", "id": 0, "operacao": "I", "ativo": "2-CFU", "antes": {"operacao": "I", "ativo": "1-CFU"}},      # o usuário já mexeu
             {"n": 2, "acao": "remover", "tabela": "outros", "id": 7, "antes": {"operacao": "I", "ativo": "x"}},                                              # a linha não existe mais
             {"n": 3, "acao": "remover", "tabela": "outros", "id": None}]
    r = _aplicar(est, itens)
    assert r["est"]["outros"]["data"][0]["ativo"] == "3-CFU" and len(r["r"]["ignoradas"]) == 3
    assert {x["motivo"] for x in r["r"]["ignoradas"]} == {"a linha mudou depois do pedido", "a linha não existe mais"}
    assert r["resumo"].startswith("nada aplicado; 3 ignorada(s)")


def test_remover_varias_linhas_nao_desloca_os_ids():
    est = {"cabos": {"data": []}, "outros": {"data": [{"ativo": f"{i}-A", "operacao": "I", "entidade": "0"} for i in range(5)]}}
    itens = [{"n": i, "acao": "remover", "tabela": "outros", "id": i, "antes": {"ativo": f"{i}-A"}} for i in (0, 1, 3)]
    assert [x["ativo"] for x in _aplicar(est, itens)["est"]["outros"]["data"]] == ["2-A", "4-A"]


def test_descrever_item():
    r = _node("""console.log(JSON.stringify(e.map(A.descreverItem)))""", [
        {"acao": "editar", "tabela": "outros", "operacao": "I", "ativo": "2-CFU", "antes": {"ativo": "1-CFU"}, "avisos": ["x"]},
        {"acao": "remover", "tabela": "cabos", "operacao": "I", "ativo": "CAA2 3 35 m"},
        {"acao": "obra", "modo": "carregar", "nome": "Rede Norte", "avisos": []}])
    assert r[0]["simbolo"] == "✎" and r[0]["tabela"] == "Outros" and r[0]["antes"] == "1-CFU" and r[0]["avisos"] == ["x"]
    assert r[1]["classe"] == "remover" and r[1]["tabela"] == "Cabos"
    assert "Carregar (substitui a tela)" in r[2]["texto"] and "Rede Norte" in r[2]["texto"]


def test_servido_sem_cache(tmp_path):
    from fastapi.testclient import TestClient
    import app as appmod
    r = TestClient(appmod.app).get("/static/assistente.js")
    assert r.status_code == 200 and "no-store" in r.headers["cache-control"] and "aplicarPlano" in r.text


def test_presets_e_corpo_do_reserva():
    r = _node("console.log(JSON.stringify({ p: A.PRESETS_RESERVA, a: A.corpoReserva({ ativa: true, provider: 'openai', url: ' https://api.deepseek.com ', modelo: ' deepseek-chat ', chave: '  ' }), "
              "b: A.corpoReserva({ ativa: false, provider: 'gemini', url: 'https://x', modelo: '', chave: 'k', limparChave: true }) }))")
    ids = {p["id"]: p for p in r["p"]}
    assert {"deepseek", "qwen", "openrouter", "ollama", "gemini", "claude"} <= set(ids)
    assert ids["deepseek"]["url"].startswith("https://api.deepseek.com") and ids["qwen"]["url"].startswith("https://dashscope-intl.aliyuncs.com/compatible-mode")
    assert all(p["provider"] in ("openai", "gemini", "claude") for p in r["p"]) and all(p["provider"] == "openai" for p in r["p"] if p["url"])
    assert r["a"] == {"ativa": True, "provider": "openai", "base_url": "https://api.deepseek.com", "model": "deepseek-chat", "api_key": None, "limpar_chave": False}
    assert r["b"] == {"ativa": False, "provider": "gemini", "base_url": "", "model": "", "api_key": "k", "limpar_chave": True}      # URL só vale para OpenAI-compatível


def test_resumo_trata_o_aviso_do_reserva_e_monta_o_formulario():
    js, html = _ler("resumo.js"), _ler("resumo.html")
    assert "ev.tipo === 'aviso'" in js and "Assistente.reservaUI" in js and 'id="assistente-reserva"' in html


def _auth_fetch(corpo_401, status=401):
    """Roda static/auth_fetch.js numa «janela» falsa e devolve {redirecionou, tokenRestou}."""
    js = r"""
        const fs = require('fs'); const store = { auth_token: 'T', auth_lembrar: 'true' }; let destino = null;
        global.window = { location: { pathname: '/resumo', search: '', set href(v) { destino = v; }, replace(v) { destino = v; } }, fetch: async () => ({ status: %d, clone() { return { json: async () => (%s) }; } }) };
        global.localStorage = { getItem: k => (k in store ? store[k] : null), removeItem: k => { delete store[k]; }, setItem: (k, v) => { store[k] = v; } };
        global.sessionStorage = { getItem: () => '1' }; global.Headers = class {}; global.alert = () => {};
        eval(fs.readFileSync(%s, 'utf8'));
        window.fetch('/api/assistente/chat', {}).then(() => console.log(JSON.stringify({ redirecionou: destino, tokenRestou: !!store.auth_token })));
    """ % (status, json.dumps(corpo_401) if isinstance(corpo_401, dict) else corpo_401, json.dumps(os.path.join(config.STATIC_DIR, "auth_fetch.js")))
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def test_401_de_falta_de_chave_da_ia_nao_derruba_a_sessao():
    for detalhe in ("API Key não encontrada.", "API Key não fornecida, não salva e sem chave padrão do sistema."):
        r = _auth_fetch({"detail": detalhe})
        assert r["redirecionou"] is None and r["tokenRestou"] is True, detalhe


def test_401_de_sessao_expirada_continua_voltando_ao_login():
    r = _auth_fetch({"detail": "Não autenticado."})
    assert r["redirecionou"] == "/login?next=%2Fresumo" and r["tokenRestou"] is False
    r = _auth_fetch("(() => { throw new Error('sem json'); })()")                 # corpo ilegível: trata como sessão expirada, como sempre
    assert r["redirecionou"] == "/login?next=%2Fresumo"
    assert _auth_fetch({"detail": "x"}, status=500)["redirecionou"] is None


def test_chat_tem_andamento_visivel_e_botao_parar():
    js = _ler("resumo.js")
    assert "new AbortController()" in js and "signal: chatAbort.signal" in js and "modoParar(true)" in js and "modoParar(false)" in js
    assert "'Pensando'" in js and "setInterval(desenharStatus, 1000)" in js and "clearInterval(relogio)" in js
    assert "error.name === 'AbortError'" in js and "if (chatAbort) chatAbort.abort(); else sendToGemini();" in js
