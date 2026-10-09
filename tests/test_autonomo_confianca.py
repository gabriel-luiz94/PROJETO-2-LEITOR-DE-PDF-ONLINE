"""TASK-059 etapa 3 — nível de confiança do autônomo: taxa de acerto por tipo de arquivo, janela, reinício na aprovação e confirmação automática."""
import json

import pytest
from fastapi.testclient import TestClient

pytest.importorskip("quickjs")
import app as appmod  # noqa: E402
import database  # noqa: E402
from middleware.auth_middleware import create_jwt_token  # noqa: E402
from services.autonomo import aprendizado_repo as repo  # noqa: E402
from services.autonomo import config_autonomo, confianca, pipeline as pl  # noqa: E402
from tests.test_autonomo_pipeline import RECEITA_REMOVE, RECEITA_SUBST, ambiente, leitor, rodar  # noqa: E402,F401


def rev(i, correcoes=0, arquivo=None, dia=None):
    return {"arquivo": arquivo or f"a{i}.dxf", "atualizado_em": f"2026-10-{(dia or i):02d}T10:00:00", "correcoes": correcoes}


# ── unidades ────────────────────────────────────────────────────────────────
def test_tipo_do_arquivo():
    assert [confianca.tipo_do_arquivo(n) for n in ("a.DXF", "b.pdf", "c", "", None, "d.tar.gz")] == ["dxf", "pdf", "(sem extensão)", "(sem extensão)", "(sem extensão)", "gz"]


def test_observando_ate_ter_a_janela_cheia():
    e = confianca.avaliar([rev(i) for i in range(1, 10)], 10, 0.9)
    assert e["nivel"] == "observando" and e["revisadas"] == 9 and e["taxa"] == 1.0
    assert confianca.avaliar([], 10, 0.9)["nivel"] == "observando" and confianca.avaliar([], 10, 0.9)["taxa"] is None


def test_confiavel_na_taxa_minima_e_revisar_abaixo():
    ok = [rev(i) for i in range(1, 11)]
    assert confianca.avaliar(ok, 10, 0.9)["nivel"] == "confiavel"
    um_ruim = [rev(i, correcoes=(3 if i == 5 else 0)) for i in range(1, 11)]
    assert confianca.avaliar(um_ruim, 10, 0.9)["nivel"] == "confiavel" and confianca.avaliar(um_ruim, 10, 0.9)["taxa"] == 0.9          # 90% exatos passa
    dois_ruins = [rev(i, correcoes=(1 if i in (3, 8) else 0)) for i in range(1, 11)]
    e = confianca.avaliar(dois_ruins, 10, 0.9)
    assert e["nivel"] == "revisar" and e["acertos"] == 8 and e["rotulo"] == "Precisa de revisão"


def test_so_as_mais_recentes_contam():
    antigas_ruins = [rev(i, correcoes=5) for i in range(1, 6)]
    recentes_boas = [rev(i, dia=10 + i) for i in range(1, 11)]
    e = confianca.avaliar(antigas_ruins + recentes_boas, 10, 0.9)
    assert e["nivel"] == "confiavel" and e["revisadas"] == 10 and e["acertos"] == 10
    # e a ordem de entrada não importa
    assert confianca.avaliar(recentes_boas + antigas_ruins, 10, 0.9)["nivel"] == "confiavel"


CFG = {"autoconfirmar_projetos": ["P1"], "confianca_janela": 10, "confianca_taxa": 0.9}


def semear(user, projeto, n, correcoes=lambda i: 0, ext="dxf", dia0=1):
    conn = repo._conectar()
    try:
        for i in range(n):
            conn.execute("INSERT OR REPLACE INTO aprendizado_correcoes (execucao_id, user_id, projeto_codigo, arquivo, eventos_json, base_json, linhas_json, criado_em, atualizado_em) "
                         "VALUES (?, ?, ?, ?, ?, '{}', '{}', ?, ?)",
                         (f"e{ext}{i}", user, projeto, f"a{i}.{ext}", json.dumps([{"tipo": "linha_excluida"}] * correcoes(i)), "x", f"2026-10-{dia0 + i:02d}T10:00:00"))
        conn.commit()
    finally:
        conn.close()


@pytest.fixture
def banco(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    return tmp_path


def test_estados_por_tipo_e_so_do_autonomo(banco):
    semear("u1", "P1", 10, ext="dxf")
    semear("u1", "P1", 4, ext="pdf", correcoes=lambda i: 1)
    conn = repo._conectar()          # sessão manual não é obra do autônomo
    conn.execute("INSERT INTO aprendizado_correcoes (execucao_id, user_id, projeto_codigo, arquivo, eventos_json, atualizado_em) VALUES ('manual:x', 'u1', 'P1', '(trabalho manual)', '[]', '2026-10-20T10:00:00')")
    conn.commit(); conn.close()
    est = confianca.estados("u1", "P1", CFG)
    assert set(est) == {"dxf", "pdf"} and est["dxf"]["nivel"] == "confiavel" and est["pdf"]["nivel"] == "observando" and est["pdf"]["acertos"] == 0
    assert confianca.estados("u2", "P1", CFG) == {} and confianca.estados("u1", "P2", CFG) == {}                      # por usuário e por projeto


def test_aprovar_proposta_reinicia_a_contagem(banco):
    semear("u1", "P1", 10)
    assert confianca.estado("u1", "P1", "dxf", CFG)["nivel"] == "confiavel"
    conn = repo._conectar()
    conn.execute("INSERT INTO aprendizado_propostas (user_id, projeto_codigo, chave, status, decidido_em) VALUES ('u1', 'P1', 'k', 'aprovada', '2026-10-15T00:00:00')")
    conn.commit(); conn.close()
    assert confianca.estado("u1", "P1", "dxf", CFG)["revisadas"] == 0 and confianca.estado("u1", "P1", "dxf", CFG)["nivel"] == "observando"
    semear("u1", "P1", 10, dia0=16)              # revisões depois da aprovação voltam a contar
    assert confianca.estado("u1", "P1", "dxf", CFG)["nivel"] == "confiavel"


def test_decisao_so_com_projeto_ligado_confiavel_e_sem_anomalia(banco, monkeypatch):
    semear("u1", "P1", 10)
    d = lambda cfg=CFG, proj="P1", pend=2, linhas=40, arq="x.dxf": confianca.decidir_autoconfirmacao("u1", proj, arq, linhas, pend, cfg)
    assert d()["auto"] is True and d()["estado"]["nivel"] == "confiavel"
    assert d(cfg={**CFG, "autoconfirmar_projetos": []})["auto"] is False and "desligada" in d(cfg={**CFG, "autoconfirmar_projetos": []})["motivo"]
    assert d(arq="x.pdf")["auto"] is False and "insuficiente" in d(arq="x.pdf")["motivo"]          # outro tipo de arquivo, sem histórico
    assert d(proj="P2", cfg={**CFG, "autoconfirmar_projetos": ["P2"]})["auto"] is False
    assert d(pend=5, linhas=10)["auto"] is True                                                    # até 5 exclusões é sempre aceitável
    assert d(pend=6, linhas=10)["auto"] is False and "anormal" in d(pend=6, linhas=10)["motivo"]
    assert d(pend=10, linhas=40)["auto"] is True and d(pend=11, linhas=40)["auto"] is False        # 25% das linhas


def test_decisao_nunca_levanta(banco, monkeypatch):
    monkeypatch.setattr(confianca, "estado", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("x")))
    d = confianca.decidir_autoconfirmacao("u1", "P1", "x.dxf", 10, 1, CFG)
    assert d["auto"] is False and d["estado"] is None


# ── pipeline ────────────────────────────────────────────────────────────────
TEXTOS = ["1-U3 1-U4", "2-CFU", "1-U4 1-SUP-L"]


def ligar(projetos=("P1",), **extra):
    config_autonomo.salvar({**config_autonomo.carregar(), "autoconfirmar_projetos": list(projetos), **extra})


def test_padrao_continua_pedindo_confirmacao(ambiente, leitor):
    semear("u1", "P1", 10)
    ex, _ = rodar(ambiente, leitor, TEXTOS, [RECEITA_SUBST, RECEITA_REMOVE])
    assert ex["status"] == "aguardando_confirmacao" and ex["relatorio"]["confianca"]["auto"] is False                # desligado por padrão, mesmo confiável


def test_confiavel_e_ligado_confirma_sozinho_e_registra(ambiente, leitor):
    semear("u1", "P1", 10)
    ligar()
    ex, ctx = rodar(ambiente, leitor, TEXTOS, [RECEITA_SUBST, RECEITA_REMOVE])
    assert ex["status"] in ("ok", "com_pendencias") and ex["obra_id"] and "confirmada(s) automaticamente" in ex["mensagem"]
    assert ex["relatorio"]["confianca"]["auto"] is True and ex["relatorio"]["confianca"]["estado"]["nivel"] == "confiavel"
    rel = json.load(open(f"{ex['pasta_saida']}/relatorio.json", encoding="utf-8"))
    assert rel["confianca"]["auto"] is True and any("Editar" in f or "Excluir" in f for f in rel["ajustes"]["aplicados"])      # as exclusões foram aplicadas
    assert pl.reverter(ex["id"], ctx)["status"] == "revertida"                                                          # e dá para reverter


def test_sem_confianca_suficiente_ou_com_acerto_baixo_pede_confirmacao(ambiente, leitor):
    ligar()
    semear("u1", "P1", 9)                                          # falta uma revisada
    ex, _ = rodar(ambiente, leitor, TEXTOS, [RECEITA_SUBST, RECEITA_REMOVE], nome="a.dxf")
    assert ex["status"] == "aguardando_confirmacao" and "insuficiente" in ex["relatorio"]["confianca"]["motivo"]
    semear("u1", "P1", 10, correcoes=lambda i: 1 if i < 2 else 0, dia0=10)      # 80% < 90%
    ex2, _ = rodar(ambiente, leitor, TEXTOS, [RECEITA_SUBST, RECEITA_REMOVE], nome="b.dxf")
    assert ex2["status"] == "aguardando_confirmacao" and ex2["relatorio"]["confianca"]["estado"]["nivel"] == "revisar"


def test_projeto_nao_ligado_ou_tipo_diferente_nao_confirma(ambiente, leitor):
    semear("u1", "P1", 10, ext="pdf")
    ligar(projetos=("OUTRO",))
    ex, _ = rodar(ambiente, leitor, TEXTOS, [RECEITA_SUBST, RECEITA_REMOVE])
    assert ex["status"] == "aguardando_confirmacao" and "desligada" in ex["relatorio"]["confianca"]["motivo"]
    ligar()
    ex2, _ = rodar(ambiente, leitor, TEXTOS, [RECEITA_SUBST, RECEITA_REMOVE], nome="outro.dxf")          # histórico é de pdf; este é dxf
    assert ex2["status"] == "aguardando_confirmacao" and "insuficiente" in ex2["relatorio"]["confianca"]["motivo"]


def test_anomalia_pede_confirmacao_mesmo_confiavel(ambiente, leitor, monkeypatch):
    semear("u1", "P1", 10)
    ligar()
    monkeypatch.setattr(confianca, "MINIMO_ANOMALIA", 0)
    monkeypatch.setattr(confianca, "LIMITE_ANOMALIA", 0.1)
    ex, _ = rodar(ambiente, leitor, TEXTOS, [RECEITA_SUBST, RECEITA_REMOVE])
    assert ex["status"] == "aguardando_confirmacao" and "anormal" in ex["relatorio"]["confianca"]["motivo"]


def test_arquivo_sem_exclusao_nao_consulta_a_confianca(ambiente, leitor, monkeypatch):
    monkeypatch.setattr(confianca, "decidir_autoconfirmacao", lambda *a, **k: (_ for _ in ()).throw(AssertionError("não devia consultar")))
    ex, _ = rodar(ambiente, leitor, ["1-U4 1-SUP-L", "2-CFU"], [RECEITA_SUBST])
    assert ex["status"] == "ok" and ex["relatorio"].get("confianca") is None


# ── configuração e API ──────────────────────────────────────────────────────
@pytest.fixture
def client(banco):
    return TestClient(appmod.app)


def cab(role="admin", uid="adm"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@x.com', role)}"}


def test_configuracao_valida_limites(banco):
    assert config_autonomo.carregar()["autoconfirmar_projetos"] == [] and config_autonomo.carregar()["confianca_janela"] == 10
    ok = config_autonomo.salvar({**config_autonomo.carregar(), "autoconfirmar_projetos": [" P1 ", "P1", "", "229"], "confianca_taxa": 0.95, "confianca_janela": 20})
    assert ok["autoconfirmar_projetos"] == ["P1", "229"] and ok["confianca_taxa"] == 0.95
    for ruim in ({"confianca_janela": 2}, {"confianca_janela": 101}, {"confianca_taxa": 0.49}, {"confianca_taxa": 1.01}, {"confianca_taxa": "x"}, {"autoconfirmar_projetos": "P1"},
                 {"autoconfirmar_projetos": [1]}):
        with pytest.raises(config_autonomo.ErroConfig):
            config_autonomo.salvar({**config_autonomo.carregar(), **ruim})


def test_api_confianca_e_toggle_pela_config(client):
    semear("adm", "P1", 10)
    r = client.put("/api/autonomo/config", json={"user_id": "adm"}, headers=cab()).json()
    assert r["config"]["autoconfirmar_projetos"] == []
    c = client.get("/api/aprendizado/confianca?projeto=P1", headers=cab()).json()
    assert c["autoconfirmar"] is False and c["tipos"][0]["tipo"] == "dxf" and c["tipos"][0]["nivel"] == "confiavel" and c["janela"] == 10 and c["taxa_minima"] == 0.9
    assert client.put("/api/autonomo/config", json={"autoconfirmar_projetos": ["P1"]}, headers=cab()).status_code == 200
    assert client.get("/api/aprendizado/confianca?projeto=P1", headers=cab()).json()["autoconfirmar"] is True
    assert client.put("/api/autonomo/config", json={"confianca_taxa": 0.1}, headers=cab()).status_code == 400
    assert client.get("/api/aprendizado/confianca?projeto=P1", headers=cab("operador", "op")).status_code == 403
    assert client.get("/api/aprendizado/confianca?projeto=P1").status_code == 401


def test_confianca_e_do_usuario_do_autonomo_nao_de_quem_olha(client):
    semear("dono", "P1", 10)
    client.put("/api/autonomo/config", json={"user_id": "dono"}, headers=cab(uid="outro_admin"))
    assert client.get("/api/aprendizado/confianca?projeto=P1", headers=cab(uid="outro_admin")).json()["tipos"][0]["nivel"] == "confiavel"
