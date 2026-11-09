"""TASK-123: a validação confere contra a base que o cálculo usa (não a cópia pessoal) e valida parâmetros também em pedidos com APELIDO."""
import pytest
from fastapi.testclient import TestClient

import database
from middleware.auth_middleware import create_jwt_token
from services.validacao_planilhas import validar_base

HDR = "ATIVO;DESC ATIVO;COMPONENTE;PROJETO;MDO;ORIGEM;CODIGO;DESC CODIGO;FATOR I;FATOR R;FILTRO;LV;ARRED"


def _csv(*linhas):
    return "\n".join([HDR] + [f"{a};D;{c};PERNAMBUCO;MATERIAL;;{k};d;1;0;;A;NORMAL" for a, c, k in linhas]).encode()


@pytest.fixture
def api(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    import app as appmod
    import routers.admin as adm
    import services.sync_service as ss
    monkeypatch.setattr(adm, "get_supabase", lambda: None)
    monkeypatch.setattr(ss, "get_supabase", lambda: None)
    cli = TestClient(appmod.app)
    cli.headers.update({"Authorization": f"Bearer {create_jwt_token('u1', 'u@x.com', 'admin')}"})
    return cli


def _importar(api, *linhas):
    r = api.post("/api/admin/upload-master", files={"file": ("b.csv", _csv(*linhas), "text/csv")}, data={"projeto": "PERNAMBUCO"})
    assert r.status_code == 200, r.text


def test_validacao_usa_a_base_do_calculo_e_nao_a_copia_pessoal(api):
    _importar(api, ("N1", "N1", "n1"))
    dados = api.get("/api/orcamento/dados?projeto=PERNAMBUCO").json()["dados"]
    assert api.post("/api/orcamento/salvar", json={"dados": dados, "projeto": "PERNAMBUCO"}).status_code == 200        # cópia pessoal com a base ANTIGA
    _importar(api, ("FLABIT=FLABIT(poste=D)", "", ""), ("FLABIT", "FLABIT", "f1"), ("FLABIT(poste=%D)", "FLABIT.D", "f2"), ("N1", "N1", "n1"))   # base nova (oficial)
    outros = [{"id": "L1", "ativo": "1-FLABIT(D)", "operacao": "I"}]
    calc = api.post("/api/orcamento/calcular", json={"cabos": [], "outros": [{"ativo": "1-FLABIT(D)", "operacao": "I", "entidade": "0"}], "projeto": "PERNAMBUCO"}).json()
    assert calc["nao_encontrados"] == [] and calc["resultado"]                                                         # o cálculo acha
    v = api.post("/api/validacao/planilhas", json={"cabos": [], "outros": outros, "projeto": "PERNAMBUCO", "payload_calculo": {"cabos": [], "outros": outros}}).json()
    assert v["achados"] == []                                                                                          # e a validação concorda


def _lp(ativo, comp="", cod="", origem=""):
    return {"ativo": ativo, "desc_ativo": "", "componente": comp, "codigo": cod, "mdo": "MATERIAL", "fator_i": 1, "fator_r": 1, "desc_codigo": "d", "projeto": "PERNAMBUCO", "origem": origem, "lv": "A"}


BASE = [_lp("IT_2=IT_2(poste=D,cabobt=35), IT-2, FLABIT"), _lp("IT_2", "IT_2", "1"), _lp("IT_2(poste=%D)", "IT_2.D", "2")]


def test_c1_param_tambem_para_pedido_com_apelido():
    def val(t):
        return validar_base([], [{"id": "L", "ativo": t, "operacao": "I"}], "PERNAMBUCO", [dict(r) for r in BASE])
    assert val("1-FLABIT(D,35)") == [] and val("IT-2(D,35)") == []
    for t, citado in (("1-FLABIT(D,A,35)", "FLABIT(D,A,35)"), ("1-FLABIT(XX,35)", "FLABIT(XX,35)"), ("IT-2(XX)", "IT-2(XX)")):
        a = val(t)
        assert [x["regra_id"] for x in a] == ["C1-PARAM"], t                                                           # antes: nenhum aviso para apelido
        assert f"'{citado}'" in a[0]["mensagem"]                                                                      # a mensagem cita o ativo digitado (com o apelido)
