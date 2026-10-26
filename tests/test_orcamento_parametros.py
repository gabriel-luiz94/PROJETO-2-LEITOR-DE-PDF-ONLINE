"""TASK-079 — cálculo com ativos parametrizados (N1(XLPE,DT)): o preview combinado com o usuário vira teste."""
import pytest

from services import orcamento_calc as calc
from tests import referencia_orcamento_calc_antes_lv as antes


def R(ativo, comp, cod, fi, lv="A", mdo="MATERIAL", proj="PERNAMBUCO", desc_ativo=""):
    return {"ativo": ativo, "desc_ativo": desc_ativo, "componente": comp, "projeto": proj, "mdo": mdo, "codigo": cod, "desc_codigo": f"d {cod}",
            "fator_i": fi, "fator_r": 0.0, "filtro": "", "origem": "OUTROS", "lv": lv}


def base_n1(decl="N1=N1(cabo=XLPE, poste=DT)", proj="PERNAMBUCO"):
    return [
        R(decl, "N1", "3340014", 1, proj=proj, desc_ativo="ESTRUTURA N1 COMUM"), R("", "N1", "3493315", 4, proj=proj), R("", "N1", "3480305", 2, proj=proj),
        R("", "N1", "SDEMU", 1, lv="LM", mdo="MAO-DE-OBRA", proj=proj), R("", "N1", "SDEVU", 1, lv="LV", mdo="MAO-DE-OBRA", proj=proj),
        R("N1(cabo=XLPE)", "N1.XLPE", "2312000", 3, proj=proj), R("", "N1.XLPE", "3428220", 3, proj=proj), R("", "N1.XLPE", "2210005", 6, proj=proj),
        R("N1(cabo=ALUM)", "N1.ALUM", "2314003", 3, proj=proj), R("", "N1.ALUM", "3428086", 3, proj=proj), R("", "N1.ALUM", "2200001", 0.27, proj=proj),
        R("", "N1.ALUM", "2278000", 0.12, proj=proj),
        R("N1(cabo=COBRE)", "N1.COBRE", "2314002", 3, proj=proj), R("", "N1.COBRE", "3428086", 3, proj=proj), R("", "N1.COBRE", "2201014", 0.45, proj=proj),
        R("N1(poste=DT)", "N1.DT", "FDT", 1, proj=proj), R("N1(poste=CC)", "N1.CC", "FCC", 1, proj=proj),
        R("N1(cabo=COBRE, poste=CC)", "N1.COBRE.CC", "FCOBCC", 1, proj=proj),
    ]


def pedido(texto, op="I"):
    return {"ativo": texto, "operacao": op, "entidade": "0"}


def totais(res):
    t = {}
    for r in res["resultado"]:
        t[r["codigo"]] = t.get(r["codigo"], 0) + r["total"]
    return {k: round(v, 2) for k, v in t.items()}


def test_preview_combinado_com_o_usuario():
    itens = [pedido("3-N1"), pedido("2-N1(ALUM,CC)", "*I"), pedido("1-N1(poste=CC, cabo=COBRE)")]
    res = calc.processar_calculo([], itens, "PERNAMBUCO", base_n1())
    assert totais(res) == {
        "3340014": 6, "3493315": 24, "3480305": 12, "SDEMU": 4, "SDEVU": 2,
        "2312000": 9, "3428220": 9, "2210005": 18,
        "2314003": 6, "3428086": 9, "2200001": 0.54, "2278000": 0.24,
        "2314002": 3, "2201014": 0.45,
        "FDT": 3, "FCC": 3, "FCOBCC": 1}
    assert res["nao_encontrados"] == [] and "avisos" not in res


def test_pedidos_equivalentes_dao_o_mesmo_resultado():
    a = calc.processar_calculo([], [pedido("3-N1")], "PERNAMBUCO", base_n1())
    for txt in ["3-N1(XLPE,DT)", "3-N1(xlpe, dt)", "3-N1(poste=DT,cabo=XLPE)", "3-N1(XLPE)"]:
        assert calc.processar_calculo([], [pedido(txt)], "PERNAMBUCO", base_n1()) == a, txt


def test_somar_o_mesmo_pedido_em_linhas_separadas():
    juntos = calc.processar_calculo([], [pedido("5-N1(ALUM,CC)")], "PERNAMBUCO", base_n1())
    partes = calc.processar_calculo([], [pedido("3-N1(ALUM,CC)"), pedido("2-N1(ALUM,CC)")], "PERNAMBUCO", base_n1())
    assert totais(juntos) == totais(partes)


def test_sem_padrao_pedido_incompleto_calcula_o_que_da_e_avisa():
    base = base_n1("N1=N1(cabo, poste)")
    res = calc.processar_calculo([], [pedido("2-N1(ALUM)")], "PERNAMBUCO", base)
    t = totais(res)
    assert t["3340014"] == 2 and t["2314003"] == 6 and "FDT" not in t and "FCC" not in t and "FCOBCC" not in t     # sem poste: só o que não depende dele
    assert len(res["avisos"]) == 1 and "POSTE" in res["avisos"][0] and "DT, CC" in res["avisos"][0]
    vazio = calc.processar_calculo([], [pedido("1-N1")], "PERNAMBUCO", base)
    assert totais(vazio) == {"3340014": 1, "3493315": 4, "3480305": 2, "SDEMU": 1} and "CABO" in vazio["avisos"][0] and "POSTE" in vazio["avisos"][0]


def test_declaracao_so_com_nomes_e_ordem_pela_primeira_aparicao():
    sem_decl = [r for r in base_n1() if True]
    sem_decl[0] = dict(sem_decl[0], ativo="N1")                        # sem declaração: a 1ª aparição define a ordem (cabo, poste)
    res = calc.processar_calculo([], [pedido("1-N1(ALUM,CC)")], "PERNAMBUCO", sem_decl)
    assert totais(res)["2314003"] == 3 and totais(res)["FCC"] == 1 and "avisos" not in res


def test_linha_so_de_declaracao_nao_entra_no_calculo():
    base = base_n1()
    base[0] = R("N1=N1(cabo=XLPE, poste=DT)", "", "", 0)               # declaração sozinha
    base.insert(1, R("N1", "N1", "3340014", 1))
    res = calc.processar_calculo([], [pedido("2-N1")], "PERNAMBUCO", base)
    assert totais(res)["3340014"] == 2 and res["nao_encontrados"] == []


def test_erros_viram_nao_encontrados_com_a_lista():
    res = calc.processar_calculo([], [pedido("1-N1(XLIPE)"), pedido("1-N1(ALUM,XX)"), pedido("1-P50(X)"), pedido("1-N1(A,B,C)")], "PERNAMBUCO", base_n1() + [R("P50", "P", "p1", 1)])
    msgs = " | ".join(res["nao_encontrados"])
    assert "XLIPE" in msgs and "existem: XLPE, ALUM, COBRE" in msgs and "POSTE=XX" in msgs and "P50 não tem variantes" in msgs and "parâmetros demais" in msgs
    assert res["resultado"] == []


def test_variante_sem_nenhum_grupo_cadastrado_para_aquele_projeto():
    # declaração e grupos só existem para RONDONIA; PERNAMBUCO não vê nada disso (N1 simples = não encontrado)
    res = calc.processar_calculo([], [pedido("1-N1(ALUM)")], "PERNAMBUCO", base_n1(proj="RONDONIA"))
    assert res["resultado"] == [] and any("N1" in m for m in res["nao_encontrados"])


def test_cruza_com_linha_viva_e_morta():
    morta = calc.processar_calculo([], [pedido("1-N1(ALUM,DT)", "I")], "PERNAMBUCO", base_n1())
    viva = calc.processar_calculo([], [pedido("1-N1(ALUM,DT)", "*I")], "PERNAMBUCO", base_n1())
    assert "SDEMU" in totais(morta) and "SDEVU" not in totais(morta)
    assert "SDEVU" in totais(viva) and "SDEMU" not in totais(viva)


def test_base_sem_parametros_e_pedidos_sem_parenteses_ficam_identicos_ao_calculo_anterior():
    base = [R("POSTE", "PST", "c1", 1.5), R("", "PST", "c2", 2), R("CAA2", "CC", "k1", 1), R("", "", "9001", 1)]
    base = [{k: v for k, v in r.items() if k != "lv"} for r in base]
    itens = [pedido("2-POSTE", "I"), pedido("3-CAA2", "*R"), pedido("1-9001", "R"), pedido("1-NADA", "I")]
    assert calc.processar_calculo([], itens, "", base) == antes.processar_calculo([], itens, "", base)


def test_base_com_parametros_nao_altera_outros_ativos():
    base = base_n1() + [R("P50", "P", "p1", 2), R("", "P", "p2", 1)]
    so_p50 = calc.processar_calculo([], [pedido("3-P50")], "PERNAMBUCO", base)
    assert totais(so_p50) == {"p1": 6, "p2": 3}


def test_tokenizar_outros_protege_parenteses():
    assert calc.tokenizar_outros("3-N1(XLPE,DT) *2-P50") == ["3", "N1(XLPE,DT)", "-2", "P50"]
    assert calc.tokenizar_outros("1-N1(poste=CC, cabo=COBRE) 2-P50") == ["1", "N1(poste=CC, cabo=COBRE)", "2", "P50"]
    assert calc.tokenizar_outros("1-N1(A-B*C)") == ["1", "N1(A-B*C)"]
    assert calc.tokenizar_outros("DT9/150") == ["1", "DT9/150"] and calc.tokenizar_outros("") == []
    itens = calc.extrair_ativos_outros({"ativo": "3-n1(xlpe, dt) *2-p50", "operacao": "i"})
    assert [(x["ativo"], x["qtd"]) for x in itens] == [("N1(XLPE, DT)", 3.0), ("P50", -2.0)]


# ── importação e rota de cálculo ──
import database  # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402
from middleware.auth_middleware import create_jwt_token  # noqa: E402

CAB = "ATIVO;DESC ATIVO;COMPONENTE;PROJETO;MDO;ORIGEM;CODIGO;DESC CODIGO;FATOR I;FATOR R;FILTRO;LV"


def csv_n1(extra=""):
    return "\n".join([CAB,
        "N1=N1(cabo=XLPE, poste=DT);ESTRUTURA N1;N1;PERNAMBUCO;MATERIAL;OUTROS;3340014;CRUZETA;1;0;;A",
        ";;N1;PERNAMBUCO;MAO-DE-OBRA;OUTROS;SDEMU;INSTALAR;1;1;;LM",
        "N1(cabo=XLPE);;N1.XLPE;PERNAMBUCO;MATERIAL;OUTROS;2312000;ISOLADOR;3;0;;A",
        "N1(cabo=ALUM);;N1.ALUM;PERNAMBUCO;MATERIAL;OUTROS;2314003;ISOLADOR;3;0;;A",
        "N1(poste=DT);;N1.DT;PERNAMBUCO;MATERIAL;OUTROS;FDT;FERRAGEM;1;0;;A",
        "N1(poste=CC);;N1.CC;PERNAMBUCO;MATERIAL;OUTROS;FCC;FERRAGEM;1;0;;A"]) + ("\n" + extra if extra else "") + "\n"


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    import routers.admin as adm, routers.orcamento as orc
    for m in (adm, orc):
        monkeypatch.setattr(m, "get_supabase", lambda: None, raising=False)
    import app as appmod
    c = TestClient(appmod.app)
    con = database.get_connection(); con.execute("DELETE FROM tabela_orcamento_master"); con.commit(); con.close()
    return c


def adm():
    return {"Authorization": f"Bearer {create_jwt_token('u1', 'a@x.com', 'admin')}"}


def subir(client, texto):
    return client.post("/api/admin/upload-master", files={"file": ("b.csv", texto.encode("utf-8-sig"), "text/csv")}, headers=adm())


def test_upload_master_com_parametros_e_calculo_pela_rota(client):
    r = subir(client, csv_n1())
    assert r.status_code == 200, r.text
    res = client.post("/api/orcamento/calcular", json={"cabos": [], "outros": [{"entidade": "0", "operacao": "I", "ativo": "2-N1(ALUM,CC)"}], "projeto": "PERNAMBUCO"}, headers=adm())
    assert res.status_code == 200, res.text
    t = {r["codigo"]: r["total"] for r in res.json()["resultado"]}
    assert t == {"3340014": 2.0, "SDEMU": 2.0, "2314003": 6.0, "FCC": 2.0}
    res2 = client.post("/api/orcamento/calcular", json={"cabos": [], "outros": [{"entidade": "0", "operacao": "I", "ativo": "1-N1(XLIPE)"}], "projeto": "PERNAMBUCO"}, headers=adm())
    assert "XLIPE" in " ".join(res2.json()["nao_encontrados"])


@pytest.mark.parametrize("linha_ruim, trecho", [
    ("N1(cabo=XL-PE);;N1.Z;PERNAMBUCO;MATERIAL;OUTROS;z1;Z;1;0;;A", "inválido"),
    ("N1(XLPE);;N1.Z;PERNAMBUCO;MATERIAL;OUTROS;z1;Z;1;0;;A", "levam nome"),
    ("N1(cabo=COBRE);;N1.XLPE;PERNAMBUCO;MATERIAL;OUTROS;z1;Z;1;0;;A", "COMPONENTE N1.XLPE"),
    ("N1(tubo=X);;N1.T;PERNAMBUCO;MATERIAL;OUTROS;z1;Z;1;0;;A", "não está na declaração"),
    ("N1=N1(cabo=XLPE, poste=DT);;N1.D2;PERNAMBUCO;MATERIAL;OUTROS;z1;Z;1;0;;A", "mais de uma declaração"),
])
def test_upload_master_recusa_cadastro_inconsistente_sem_alterar(client, linha_ruim, trecho):
    assert subir(client, csv_n1()).status_code == 200
    r = subir(client, csv_n1(linha_ruim))
    assert r.status_code == 400 and trecho in r.json()["detail"] and "Nada foi alterado" in r.json()["detail"]
    con = database.get_connection()
    try:
        assert con.execute("SELECT COUNT(*) FROM tabela_orcamento_master").fetchone()[0] == 6      # a base anterior continua
    finally:
        con.close()


def test_upload_devolve_avisos_de_sobreposicao(client):
    r = subir(client, csv_n1("N1(poste=DT);;N1.DT;PERNAMBUCO;MATERIAL;OUTROS;2312000;REPETIDO;1;0;;A".replace("N1(poste=DT);;N1.DT", ";;N1.DT")))
    assert r.status_code == 200
    avisos = r.json().get("avisos", [])
    assert len(avisos) == 1 and "2312000" in avisos[0] and "soma duas vezes" in avisos[0]


def test_salvar_sync_e_add_row_com_parametros(client):
    dados = [{"ativo": "N1=N1(cabo=XLPE)", "componente": "N1", "codigo": "c1", "projeto": "", "lv": "A"},
             {"ativo": "N1(cabo=XLPE)", "componente": "N1.X", "codigo": "x1", "projeto": "", "lv": "A"}]
    assert client.post("/api/admin/sync-master-all", json={"dados": dados}, headers=adm()).status_code == 200
    ruim = dados + [{"ativo": "N1(cabo=Z)", "componente": "N1.X", "codigo": "z", "projeto": "", "lv": "A"}]
    r = client.post("/api/admin/sync-master-all", json={"dados": ruim}, headers=adm())
    assert r.status_code == 400 and "COMPONENTE N1.X" in r.json()["detail"]
    assert client.post("/api/orcamento/salvar", json={"dados": ruim}, headers=adm()).status_code == 400
    assert client.post("/api/orcamento/salvar", json={"dados": dados}, headers=adm()).status_code == 200
    assert client.post("/api/admin/add-row", json={"ativo": "N1(cabo=XL-PE)", "codigo": "k"}, headers=adm()).status_code == 400
    assert client.post("/api/admin/add-row", json={"ativo": "N1(cabo=ALUM)", "componente": "N1.A", "codigo": "k"}, headers=adm()).status_code == 200


def test_upload_local_com_parametros_valida_antes_de_apagar(client):
    ok = csv_n1()
    assert client.post("/api/orcamento/upload", files={"file": ("b.csv", ok.encode("utf-8-sig"), "text/csv")}, headers=adm()).status_code == 200
    ruim = csv_n1("N1(cabo=COBRE);;N1.XLPE;PERNAMBUCO;MATERIAL;OUTROS;z1;Z;1;0;;A")
    r = client.post("/api/orcamento/upload", files={"file": ("b.csv", ruim.encode("utf-8-sig"), "text/csv")}, headers=adm())
    assert r.status_code == 400 and "COMPONENTE N1.XLPE" in r.json()["error"]
    con = database.get_connection()
    try:
        assert con.execute("SELECT COUNT(*) FROM tabela_orcamento").fetchone()[0] == 6
    finally:
        con.close()


# ── consumidores do ativo como texto ──
def test_regras_de_dominio_casam_pelo_nome_base_ou_texto_completo():
    from services import regras_dominio as rd
    assert rd._casa("N1(XLPE,DT)", "N1", {}) and rd._casa("N1(XLPE,DT)", "N1(XLPE,DT)", {}) and rd._casa("N1(XLPE,DT)", "N1*", {})
    assert rd._casa("N1(XLPE,DT)", {"regex": "^N1$"}, {}) and rd._casa("N1(XLPE,DT)", ["P50", "N1"], {})
    assert not rd._casa("P50", "N1", {}) and not rd._casa("N1(XLPE,DT)", "N2", {})


def test_validacao_c1_base_marca_variante_invalida_e_a_dica_de_linha_viva_morta():
    from services.validacao_planilhas import validar_base as validar_c1
    base = base_n1() + [R("CAA2", "CC", "k1", 1, lv="LM")]
    achados = validar_c1([], [{"ativo": "1-N1(XLIPE)", "operacao": "I"}, {"ativo": "1-N1(ALUM,CC)", "operacao": "I"}, {"ativo": "1-CAA2", "operacao": "*I"}], "PERNAMBUCO", base)
    textos = [a["mensagem"] for a in achados]
    assert any("N1(XLIPE)" in t for t in textos) and any("CAA2" in t for t in textos)           # a dica entre parênteses não esconde o aviso
    assert not any("N1(ALUM,CC)" in t for t in textos)


def test_autocomplete_de_ativos_das_regras_mostra_o_nome_base(client):
    assert subir(client, csv_n1()).status_code == 200
    r = client.get("/api/validacao/regras/ativos", headers=adm())
    assert r.status_code == 200
    ativos = r.json()["ativos"]
    assert "N1" in ativos and not any("(" in a or "=" in a for a in ativos)


def test_ligacoes_do_front():
    import os, config
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    js = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    assert html.index("/static/ativo_parametros.js") < html.index("/static/resumo.js")
    for trecho in ["AtivoParametros.dividirTokens(item.ativo)", "AtivoParametros.resolverPedido(aName, assinaturasParam)",
                   "AtivoParametros.nomeBase(ativoUpper)", "checkMatch(ruleAtivo, AtivoParametros.nomeBase(item.ativo))", "replace(/\\([^()]*\\)/g, '')"]:
        assert trecho in js, trecho
    r = TestClient(__import__("app").app).get("/static/ativo_parametros.js")
    assert r.status_code == 200 and "no-cache" in r.headers.get("cache-control", "").lower()
    base = open(os.path.join(config.STATIC_DIR, "orcamento.html"), encoding="utf-8").read()
    assert "mostrarAvisosBase(data)" in base
    res = open(os.path.join(config.STATIC_DIR, "resultado_orcamento.html"), encoding="utf-8").read()
    assert "data.avisos" in res and "alertaCalculo(" in res
