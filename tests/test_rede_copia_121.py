"""TASK-121: rede adaptativa (proxy × direto × offline, só desktop) e cópia local do banco (nuvem → SQLite, com trabalho offline preservado)."""
import http.server
import json
import os
import threading

import pytest
from fastapi.testclient import TestClient

import database
from services import connectivity_monitor as cm
from services import copia_local as cl
from services import rede_adaptativa as rd
from services import supabase_client as sc
from services import sync_service as ss


# ───────────────────────── rede adaptativa ─────────────────────────
@pytest.fixture
def rede(monkeypatch):
    """Estado limpo da rede (e restaura o ambiente de proxy ao final)."""
    for k in list(os.environ):
        if k.upper() in ("HTTPS_PROXY", "HTTP_PROXY", "ALL_PROXY", "NO_PROXY"):
            monkeypatch.delenv(k, raising=False)
    monkeypatch.setattr(sc, "reset_supabase_client", lambda: None)
    rd.parar()
    yield rd
    rd.parar()


def _sondas(monkeypatch, ok_por_proxy):
    """`_sondar(proxy)` de mentira: ok_por_proxy = {proxy_ou_None: bool}."""
    monkeypatch.setattr(rd, "_sondar", lambda proxy: ok_por_proxy.get(proxy, False))


def test_so_roda_no_desktop_e_nunca_sob_teste_ou_sem_supabase(monkeypatch, rede):
    monkeypatch.setenv("SUPABASE_URL", "https://x.supabase.co"); monkeypatch.setenv("SUPABASE_KEY", "k")
    monkeypatch.delenv("PYTEST_CURRENT_TEST", raising=False)
    monkeypatch.setattr(rd, "APP_MODE", "desktop")
    assert rd.habilitada() is True
    monkeypatch.setattr(rd, "APP_MODE", "server")                      # versão online: desligada
    assert rd.habilitada() is False and rd.iniciar() is None and rd.ativa() is False and rd.offline() is False
    monkeypatch.setattr(rd, "APP_MODE", "desktop")
    monkeypatch.setenv("REDE_ADAPTATIVA", "0")
    assert rd.habilitada() is False
    monkeypatch.delenv("REDE_ADAPTATIVA"); monkeypatch.delenv("SUPABASE_KEY")
    assert rd.habilitada() is False                                    # sem Supabase configurado
    monkeypatch.setenv("SUPABASE_KEY", "k"); monkeypatch.setenv("PYTEST_CURRENT_TEST", "x")
    assert rd.habilitada() is False                                    # sob teste


def test_escolha_do_caminho(monkeypatch, rede):
    P = "http://192.168.100.1:3128"
    _sondas(monkeypatch, {P: True, None: True})
    assert rd.avaliar(P) == (True, "proxy", P)                         # os dois funcionam: prefere o proxy configurado
    _sondas(monkeypatch, {P: False, None: True})
    assert rd.avaliar(P) == (True, "direto", None)                     # proxy "esquecido" (casa): vai direto
    _sondas(monkeypatch, {P: True, None: False})
    assert rd.avaliar(P) == (True, "proxy", P)                         # empresa: só passa pelo proxy
    _sondas(monkeypatch, {P: False, None: False})
    assert rd.avaliar(P) == (False, "nenhum", P)                       # nada responde: offline
    _sondas(monkeypatch, {None: True})
    monkeypatch.setattr(rd, "_PROXY_ORIGINAL", None)
    assert rd.avaliar() == (True, "direto", None)                      # sem proxy configurado
    _sondas(monkeypatch, {})
    assert rd.avaliar() == (False, "nenhum", None)


def test_aplicar_o_caminho_no_ambiente(monkeypatch, rede):
    P = "http://192.168.100.1:3128"
    rd._aplicar("proxy", P)
    assert os.environ["HTTPS_PROXY"] == P and os.environ["HTTP_PROXY"] == P and "NO_PROXY" not in os.environ
    rd._aplicar("direto", None)
    assert os.environ["NO_PROXY"] == "*" and "HTTPS_PROXY" not in os.environ and "HTTP_PROXY" not in os.environ      # "*" ignora até o proxy do Windows
    rd._aplicar("proxy", P)
    assert "NO_PROXY" not in os.environ and os.environ["HTTPS_PROXY"] == P


def test_disjuntor_offline_ouvintes_e_reconexao(monkeypatch, rede):
    eventos = []
    monkeypatch.setitem(rd._ouvintes, "online", [lambda: eventos.append("online")])
    monkeypatch.setitem(rd._ouvintes, "offline", [lambda: eventos.append("offline")])
    monkeypatch.setitem(rd._ouvintes, "mudanca", [lambda e: eventos.append(("mudanca", e["caminho"]))])
    rd._estado.update(iniciado=True, online=True, caminho="direto", proxy=None)
    assert rd.offline() is False
    rd._atualizar(False, "nenhum", None)                               # a nuvem caiu
    assert rd.offline() is True and eventos == ["offline", ("mudanca", "nenhum")]
    # disjuntor: get_supabase() devolve None na hora, sem tocar na rede
    monkeypatch.setattr(sc, "_client_initialized", False)
    assert sc.get_supabase() is None
    rd._atualizar(False, "nenhum", None)                               # sem mudança: nenhum evento novo
    assert len(eventos) == 2
    rd._atualizar(True, "proxy", "http://p:1")                         # voltou
    assert rd.offline() is False and eventos[2:] == ["online", ("mudanca", "proxy")]
    assert rd.estado()["proxy"] == "p:1" and rd.estado()["ativa"] is True       # o estado exposto não leva o endereço completo


def test_get_supabase_da_versao_online_nao_muda(monkeypatch, rede):
    """Sem a vigia ligada (versão online/servidor) nada muda: o cliente é criado normalmente, mesmo que o estado interno diga offline."""
    monkeypatch.setenv("SUPABASE_URL", "https://x.supabase.co"); monkeypatch.setenv("SUPABASE_KEY", "k")
    criado = []
    monkeypatch.setattr(sc, "create_client", lambda u, k, *a, **kw: criado.append((u, kw)) or object())
    monkeypatch.setattr(sc, "_supabase_client", None); monkeypatch.setattr(sc, "_client_initialized", False)
    rd._estado.update(iniciado=False, online=False)                    # "offline" mas a vigia desligada → ignorado
    assert sc.get_supabase() is not None and criado == [("https://x.supabase.co", {})]      # cliente padrão, sem opções de timeout


def test_cliente_com_timeout_curto_so_no_desktop_com_vigia(monkeypatch, rede):
    monkeypatch.setattr(sc, "_supabase_client", None); monkeypatch.setattr(sc, "_client_initialized", False)
    monkeypatch.setenv("SUPABASE_URL", "https://x.supabase.co"); monkeypatch.setenv("SUPABASE_KEY", "k")
    chamadas = []
    monkeypatch.setattr(sc, "create_client", lambda u, k, *a, **kw: chamadas.append(kw) or object())
    rd._estado.update(iniciado=True, online=True)
    sc.get_supabase()
    assert chamadas and chamadas[0].get("options") is not None and chamadas[0]["options"].postgrest_client_timeout.connect == 5.0


def test_falha_de_rede_acorda_a_vigia(monkeypatch, rede):
    import httpx
    chamou = []
    monkeypatch.setattr(rd, "sinalizar_falha", lambda: chamou.append(1))
    sc.registrar_falha("teste", httpx.ConnectTimeout("timeout"))
    assert chamou == [1]
    sc.registrar_falha("teste", ValueError("dado ruim"))              # erro que não é de rede não mexe na vigia
    assert chamou == [1]


def test_monitor_de_conectividade_espelha_a_vigia(monkeypatch, rede):
    rd._estado.update(iniciado=True, online=False, caminho="nenhum")
    assert cm._testar_conexao() is False
    msg = json.loads(cm._mensagem(False))
    assert msg["status"] == "offline" and msg["caminho"] == "nenhum" and "copia" in msg
    rd._estado.update(online=True, caminho="proxy")
    assert cm._testar_conexao() is True
    rd.parar()
    assert json.loads(cm._mensagem(True)) == {"type": "connectivity", "status": "online"}      # versão online: mensagem como sempre


def test_sonda_real_responde_em_http_e_reprova_porta_fechada(monkeypatch, rede):
    class H(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(404); self.end_headers(); self.wfile.write(b'{"error":"requested path is invalid"}')
        def log_message(self, *a): pass
    srv = http.server.HTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    try:
        monkeypatch.setenv("SUPABASE_URL", f"http://127.0.0.1:{srv.server_address[1]}"); monkeypatch.setenv("SUPABASE_KEY", "k")
        assert rd._sondar(None) is True                                # qualquer resposta HTTP (até 404) = a nuvem responde
        assert rd._sondar("http://127.0.0.1:9") is False               # proxy que não responde: reprova rápido
    finally:
        srv.shutdown()
    monkeypatch.setenv("SUPABASE_URL", "http://127.0.0.1:9")
    assert rd._sondar(None) is False


def test_iniciar_liga_a_vigia_e_aplica(monkeypatch, rede):
    monkeypatch.setattr(rd, "habilitada", lambda: True)
    monkeypatch.setattr(rd, "avaliar", lambda *a, **k: (True, "direto", None))
    monkeypatch.setattr(rd, "INTERVALO_ONLINE", 3600)
    est = rd.iniciar()
    assert est["ativa"] is True and est["online"] is True and est["caminho"] == "direto" and os.environ.get("NO_PROXY") == "*"
    assert rd.iniciar() is None                                        # só uma vez
    rd.parar()
    monkeypatch.setattr(rd, "avaliar", lambda *a, **k: (False, "nenhum", None))
    est = rd.iniciar()
    assert est["online"] is False and rd.offline() is True


# ───────────────────────── cópia local ─────────────────────────
class FakeNuvem:
    """Supabase de mentira em memória (select/order/range, upsert(on_conflict), delete().eq())."""
    def __init__(self, tabelas=None):
        self.t = {k: [dict(r) for r in v] for k, v in (tabelas or {}).items()}
        self.falha_upsert = False
        self.upserts, self.deletes = [], []

    def table(self, nome):
        return _Q(self, nome)


class _Q:
    def __init__(self, db, nome):
        self.db, self.nome, self.modo, self.ordem, self.faixa, self.filtros = db, nome, "select", [], None, {}

    def select(self, *_): self.modo = "select"; return self
    def order(self, c): self.ordem.append(c); return self
    def range(self, a, b): self.faixa = (a, b); return self
    def upsert(self, rows, on_conflict=None): self.modo = "upsert"; self.rows = rows if isinstance(rows, list) else [rows]; self.conf = (on_conflict or "").split(","); return self
    def delete(self): self.modo = "delete"; return self
    def eq(self, c, v): self.filtros[c] = v; return self

    def execute(self):
        class R: pass
        r = R(); r.data = []
        if self.nome not in self.db.t:
            raise RuntimeError(f"relation {self.nome} does not exist")
        tab = self.db.t[self.nome]
        if self.modo == "select":
            linhas = sorted(tab, key=lambda x: tuple(str(x.get(c)) for c in self.ordem)) if self.ordem else list(tab)
            r.data = [dict(x) for x in (linhas[self.faixa[0]:self.faixa[1] + 1] if self.faixa else linhas)]
        elif self.modo == "upsert":
            if self.db.falha_upsert:
                raise ConnectionError("sem rede")
            for novo in self.rows:
                self.db.upserts.append((self.nome, dict(novo)))
                alvo = next((x for x in tab if all(x.get(c) == novo.get(c) for c in self.conf)), None)
                if alvo is None:
                    tab.append(dict(novo))
                else:
                    alvo.update(novo)
        elif self.modo == "delete":
            self.db.deletes.append((self.nome, dict(self.filtros)))
            tab[:] = [x for x in tab if not all(x.get(c) == v for c, v in self.filtros.items())]
        return r


@pytest.fixture
def banco(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    monkeypatch.setattr(cl, "APP_MODE", "desktop")
    monkeypatch.setattr(ss, "get_supabase", lambda: None)            # a base de orçamento (master) tem teste próprio
    return database


def _rodar(monkeypatch, nuvem):
    monkeypatch.setattr(sc, "get_supabase", lambda: nuvem)
    rd.parar()
    return cl.sincronizar("teste")


def _local(tabela, where="", args=()):
    c = database.get_row_connection()
    try:
        return [dict(r) for r in c.execute(f"SELECT * FROM {tabela} {where}", args).fetchall()]
    finally:
        c.close()


def _nuvem_base():
    return FakeNuvem({
        "projetos": [{"nome": "PERNAMBUCO", "codigo": "026", "updated_at": "2026-11-01T10:00:00+00:00"}],
        "obras": [{"id": "o1", "nome": "Obra 1", "data": "d", "dados_json": "{}", "user_id": "7", "projeto": "026", "publica": False, "dono_nome": "", "tipo": "obra"},
                  {"id": "o2", "nome": "Obra 2", "data": "d", "dados_json": "{}", "user_id": "7", "projeto": "026", "publica": True, "dono_nome": "ana", "tipo": "obra"}],
        "atalhos_modais": [{"projeto_codigo": "026", "modal": "montaveis", "dados_json": '{"versao":1,"modais":[]}'}],
        "regras_conversao": [{"projeto_codigo": "026", "regras_json": '[{"a":1}]'}],
        "usuarios_nuvem": [{"id": 7, "email": "ana@x.com", "senha_hash": database.hash_password("segredo1"), "role": "operador"},
                           {"id": 8, "email": "adm@x.com", "senha_hash": database.hash_password("adm12345"), "role": "admin"}],
        **{t: [] for t in ["historico_rec", "contextos", "ajustes_contextos", "ajustes_planilhas", "prompts_validacao", "regras_dominio", "regras_leitor_classificacao",
                           "regras_leitor_processamento", "regras_vinculacao", "vinculacao_config", "ajustes_contextos_historico", "ajustes_planilhas_historico",
                           "prompts_validacao_historico", "regras_dominio_historico", "regras_leitor_classificacao_historico", "regras_leitor_processamento_historico", "regras"]},
    })


def test_primeira_copia_deixa_o_local_igual_a_nuvem_e_guarda_o_que_so_existia_aqui(banco, monkeypatch):
    c = database.get_connection()
    c.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES ('gemini_api_key', 'segredo-local')")        # NUNCA é tocada nem enviada
    c.commit(); c.close()
    nuvem = _nuvem_base()
    res = _rodar(monkeypatch, nuvem)
    assert res["projetos"]["aplicadas"] == 1
    assert [p["nome"] for p in _local("projetos")] == ["PERNAMBUCO"]                      # PARAIBA/RONDONIA locais (sementes) não existem na nuvem → saem
    assert {o["id"] for o in _local("obras")} == {"o1", "o2"} and _local("obras", "WHERE id='o2'")[0]["publica"] in (1, True)
    assert _local("atalhos_modais")[0]["dados_json"] == '{"versao":1,"modais":[]}'
    cfg = {r["chave"]: r["valor"] for r in _local("configuracoes")}
    assert cfg["regras_conversao_026"] == '[{"a":1}]' and cfg["gemini_api_key"] == "segredo-local"        # regras de conversão → configuracoes; o resto intacto
    perdidos = _local("conflitos_locais", "WHERE tabela='projetos'")
    assert {json.loads(p["dados_json"])["nome"] for p in perdidos} >= {"PARAIBA", "RONDONIA"}      # o que só existia aqui ficou guardado
    assert nuvem.upserts == [] and nuvem.deletes == []                                      # a primeira cópia não envia nada para a nuvem
    assert cl.estado()["ultima"] and cl.estado()["erro"] is None


def test_contas_de_usuario_e_login_offline_com_o_mesmo_user_id(banco, monkeypatch):
    _rodar(monkeypatch, _nuvem_base())
    u = {x["email"]: x for x in _local("usuarios_locais")}
    assert u["ana@x.com"]["nuvem_id"] == "7" and u["adm@x.com"]["role"] == "admin" and u["adm@x.com"]["is_admin"] == 1
    import app as appmod
    import routers.auth as auth
    monkeypatch.setattr(auth, "get_supabase", lambda: None)                              # sem internet
    cli = TestClient(appmod.app)
    r = cli.post("/api/auth/login", json={"email": "ana@x.com", "password": "segredo1"})
    assert r.status_code == 200 and r.json()["user_id"] == "7"                          # o MESMO id do login online (as obras dela continuam dela)
    assert cli.post("/api/auth/login", json={"email": "ana@x.com", "password": "errada"}).status_code == 401
    assert cli.post("/api/auth/login", json={"email": "adm@x.com", "password": "adm12345"}).json()["role"] == "admin"


def test_trabalho_offline_e_preservado_e_enviado_depois(banco, monkeypatch):
    nuvem = _nuvem_base()
    _rodar(monkeypatch, nuvem)                                                           # 1ª cópia (cria a "foto")
    # ── offline: edita a obra o1, cria o3, apaga o2 (só no banco local) ──
    c = database.get_connection()
    c.execute("UPDATE obras SET nome = 'Obra 1 editada', dados_json = '{\"x\":1}' WHERE id = 'o1'")
    c.execute("INSERT INTO obras (id, nome, data, dados_json, user_id, projeto, publica, dono_nome, tipo) VALUES ('o3','Nova','d','{}','7','026',0,'','obra')")
    c.execute("DELETE FROM obras WHERE id = 'o2'")
    c.commit(); c.close()
    _rodar(monkeypatch, nuvem)                                                           # a nuvem voltou
    ids = {o["id"]: o for o in nuvem.t["obras"]}
    assert set(ids) == {"o1", "o3"} and ids["o1"]["nome"] == "Obra 1 editada"           # editou, criou e apagou na nuvem também
    assert {o["id"] for o in _local("obras")} == {"o1", "o3"}
    assert ("obras", {"id": "o2"}) in nuvem.deletes
    _rodar(monkeypatch, nuvem)                                                           # 3ª cópia: tudo em dia, nada a enviar
    n_up = len(nuvem.upserts)
    _rodar(monkeypatch, nuvem)
    assert len(nuvem.upserts) == n_up and not nuvem.deletes[1:]


def test_mudanca_so_na_nuvem_e_aplicada_e_conflito_vence_a_nuvem_com_backup(banco, monkeypatch):
    nuvem = _nuvem_base()
    _rodar(monkeypatch, nuvem)
    # outra pessoa muda o2 e o1 na nuvem; aqui (offline) só o1 também foi editado → conflito no o1
    for o in nuvem.t["obras"]:
        o["nome"] = "Nuvem " + o["id"]
    c = database.get_connection(); c.execute("UPDATE obras SET nome = 'Local o1' WHERE id = 'o1'"); c.commit(); c.close()
    _rodar(monkeypatch, nuvem)
    nomes = {o["id"]: o["nome"] for o in _local("obras")}
    assert nomes == {"o1": "Nuvem o1", "o2": "Nuvem o2"}                                 # nuvem venceu nos dois (o2 só mudou lá; o1 foi conflito)
    perdido = _local("conflitos_locais", "WHERE tabela='obras'")
    assert len(perdido) == 1 and json.loads(perdido[0]["dados_json"])["nome"] == "Local o1" and "conflito" in perdido[0]["motivo"]
    assert nuvem.upserts == []


def test_envio_que_falha_nao_perde_o_trabalho_e_tenta_de_novo(banco, monkeypatch):
    nuvem = _nuvem_base()
    _rodar(monkeypatch, nuvem)
    c = database.get_connection(); c.execute("UPDATE obras SET nome = 'Editada offline' WHERE id = 'o1'"); c.commit(); c.close()
    nuvem.falha_upsert = True                                                            # a rede cai de novo no meio
    _rodar(monkeypatch, nuvem)
    assert {o["id"]: o["nome"] for o in _local("obras")}["o1"] == "Editada offline"      # continua no local
    assert [o["nome"] for o in nuvem.t["obras"] if o["id"] == "o1"] == ["Obra 1"]
    assert cl.estado()["pendencias"] != "0"
    nuvem.falha_upsert = False
    _rodar(monkeypatch, nuvem)                                                           # volta: envia
    assert [o["nome"] for o in nuvem.t["obras"] if o["id"] == "o1"] == ["Editada offline"]
    assert not _local("conflitos_locais", "WHERE tabela='obras'")                        # sem conflito falso


def test_regras_conversao_vai_e_volta_pelo_configuracoes(banco, monkeypatch):
    nuvem = _nuvem_base()
    _rodar(monkeypatch, nuvem)
    c = database.get_connection()
    c.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES ('regras_conversao_026', '[{\"a\":2}]')")      # editou offline
    c.commit(); c.close()
    _rodar(monkeypatch, nuvem)
    assert nuvem.t["regras_conversao"] == [{"projeto_codigo": "026", "regras_json": '[{"a":2}]'}]
    assert ("configuracoes", ) not in [(t,) for t, _ in nuvem.upserts]                   # a tabela configuracoes (com a chave da API) nunca é enviada


def test_espelho_do_historico_e_tabela_ausente_na_nuvem(banco, monkeypatch):
    nuvem = _nuvem_base()
    nuvem.t["regras_dominio_historico"] = [{"id": 5, "projeto_codigo": "026", "regras_json": "[]", "criado_em": "2026-01-01", "criado_por": "a"}]
    del nuvem.t["regras_vinculacao"]                                                     # tabela que não existe na nuvem: só essa falha, o resto segue
    res = _rodar(monkeypatch, nuvem)
    assert [h["id"] for h in _local("regras_dominio_historico")] == [5]
    assert "regras_vinculacao" not in res and "projetos" in res
    assert "tabela(s) com falha" in (cl.estado()["erro"] or "")


def test_versao_online_nao_copia_nada(banco, monkeypatch):
    monkeypatch.setattr(cl, "APP_MODE", "server")
    nuvem = _nuvem_base()
    assert _rodar(monkeypatch, nuvem) is None and cl.iniciar_em_segundo_plano() is False
    assert [p["nome"] for p in _local("projetos")] != ["PERNAMBUCO"]                    # banco intocado
    import app as appmod
    import routers.rede as rr
    monkeypatch.setattr(rr, "APP_MODE", "server")
    assert TestClient(appmod.app).get("/api/rede/estado", headers={"Authorization": _tok()}).json() == {"modo": "server", "ativa": False, "online": True}


def _tok():
    from middleware.auth_middleware import create_jwt_token
    return f"Bearer {create_jwt_token('u', 'u@x.com', 'admin')}"


def test_endpoint_de_estado_no_desktop(banco, monkeypatch):
    import app as appmod
    import routers.rede as rr
    monkeypatch.setattr(rr, "APP_MODE", "desktop")
    rd.parar()
    d = TestClient(appmod.app).get("/api/rede/estado", headers={"Authorization": _tok()}).json()
    assert d["modo"] == "desktop" and d["ativa"] is False and d["online"] is True and set(d["copia"]) >= {"ultima", "em_andamento", "erro"}


def test_nao_roda_offline(banco, monkeypatch):
    nuvem = _nuvem_base()
    rd._estado.update(iniciado=True, online=False)
    monkeypatch.setattr(sc, "get_supabase", lambda: nuvem)
    assert cl.sincronizar("x") is None                                                    # sem nuvem não copia nada
    rd.parar()
