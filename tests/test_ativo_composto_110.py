"""TASK-110 — ativo composto: N1C:=N1(C) + 2*N2(C) agrupa outros ativos (Python e JS em paridade; cálculo expande o pedido)."""
import json
import os
import subprocess

import pytest
from fastapi.testclient import TestClient

import config
import database
from middleware.auth_middleware import create_jwt_token
from services import ativo_parametros as ap
from services import orcamento_calc as calc
from services.autonomo.totalizadora import montar_totalizadora
from tests import oraculo_tela as o

JS = os.path.join(config.STATIC_DIR, "ativo_parametros.js")


def L(ativo, comp="", cod="", fi=1.0, proj="", **k):
    r = {"ativo": ativo, "desc_ativo": "d", "componente": comp, "projeto": proj, "mdo": "MATERIAL", "codigo": cod, "desc_codigo": f"d {cod}",
         "fator_i": fi, "fator_r": fi, "filtro": "", "origem": "", "lv": "A"}
    r.update(k)
    return r


BASE = [L("N1=N1(tipo)"), L("N1(tipo=C)", "N1C_", "k1c"), L("N1(tipo=D)", "N1D_", "k1d"),
        L("N2=N2(tipo)"), L("N2(tipo=C)", "N2C_", "k2c"), L("T3", "T3_", "t3"),
        L("N1C:=N1(C) + 2*N2(C)"), L("KIT:=N1C + T3, K"), L("[A,B]", "AB_", "kab"), L("KAB:=3*[A,B]")]


def calcular(texto, base=None, op="I", projeto=""):
    r = calc.processar_calculo([], [{"ativo": texto, "operacao": op, "entidade": "0"}], projeto, base if base is not None else BASE)
    return {x["codigo"]: x["total"] for x in r["resultado"]}, r["nao_encontrados"], r.get("avisos")


# ── leitura da célula ──
def test_analisar_composto():
    a = ap.analisar_celula_base("N1C:=N1(C) + 2*N2(C)")
    assert (a["tipo"], a["base"], a["erro"], a["apelidos"]) == ("composto", "N1C", None, [])
    assert a["itens"] == [{"ativo": "N1(C)", "qtd": 1.0}, {"ativo": "N2(C)", "qtd": 2.0}]
    a = ap.analisar_celula_base("KIT := N1C + 0.5*T3, K")
    assert (a["base"], a["apelidos"], a["itens"]) == ("KIT", ["K"], [{"ativo": "N1C", "qtd": 1.0}, {"ativo": "T3", "qtd": 0.5}])
    a = ap.analisar_celula_base("[S 1]:=2*[CAZ 7,94] + N1(poste=DT)")
    assert (a["base"], a["erro"]) == ("S 1", None) and a["itens"][0] == {"ativo": "[CAZ 7,94]", "qtd": 2.0}
    assert ap.nome_base("N1C:=N1(C)+N2(C), X") == "N1C" and ap.nome_base("N1=N1(a)") == "N1"
    for ruim in ["N1C:=", "N1C:=N1(C)+", "N1C:=0*A", "N1C:=N1=N1(a)", "A B:=X", "N1C:=N1(C", ":=A"]:
        assert ap.analisar_celula_base(ruim)["erro"], ruim


# ── validação da base ──
def erros(*linhas):
    return ap.validar_base([dict(r) for r in linhas])["erros"]


def test_base_valida_sem_erros():
    assert erros(*BASE) == []


def test_validar_composto():
    assert any("não existe" in e and "X9" in e for e in erros(*BASE, L("Z:=X9 + T3")))
    assert any("'N1(Z)'" in e or "N1(Z)" in e for e in erros(*BASE, L("Z:=N1(Z)")))            # valor de parâmetro não cadastrado
    assert any("em branco" in e for e in erros(*BASE, L("Z:=T3", "C", "k")))                  # composto não tem código/componente
    assert any("já foi definido" in e for e in erros(*BASE, L("N1C:=T3")))
    assert any("já existe como ativo" in e for e in erros(*BASE, L("T3:=N1C")))
    assert any("maior que zero" in e for e in erros(*BASE, L("Z:=0*T3")))
    ciclo = erros(*BASE, L("P:=Q"), L("Q:=P + T3"))
    assert any("ciclo" in e for e in ciclo)
    assert any("ciclo" in e for e in erros(*BASE, L("P:=P")))


def test_composto_nao_vira_grupo_nem_ganha_componente():
    novas, n = ap.preencher_componentes([dict(r) for r in BASE])
    comp = [r for r in novas if r["ativo"].startswith(("N1C:=", "KIT:=", "KAB:="))]
    assert n == 0 or all(not r["componente"] for r in comp)
    assert all(not r["componente"] for r in comp)
    assert ap.assinaturas_do_projeto([dict(r) for r in BASE], "").keys() == {"N1", "N2"}         # compostos não são ativos com parâmetros


# ── cálculo ──
def test_pedido_do_composto_expande_os_itens():
    assert calcular("3-N1C")[:2] == ({"k1c": 3.0, "k2c": 6.0}, [])
    assert calcular("2-N1C 1-T3")[0] == {"k1c": 2.0, "k2c": 4.0, "t3": 1.0}
    assert calcular("N1C, 2 T3")[0] == {"k1c": 1.0, "k2c": 2.0, "t3": 2.0}                   # compacto
    assert calcular("3-N1(C)")[0] == {"k1c": 3.0}                                           # o ativo avulso continua valendo


def test_composto_dentro_de_composto_apelido_e_nome_literal():
    assert calcular("2-KIT")[0] == {"k1c": 2.0, "k2c": 4.0, "t3": 2.0}
    assert calcular("1-K")[0] == {"k1c": 1.0, "k2c": 2.0, "t3": 1.0}                         # apelido do composto
    assert calcular("1-KAB")[0] == {"kab": 3.0}                                             # item [A,B] com vírgula
    assert calcular("1-[A,B]")[0] == {"kab": 1.0}


def test_operacao_e_remocao_valem_para_cada_item():
    r = calc.processar_calculo([], [{"ativo": "3-N1C", "operacao": "R", "entidade": "0"}], "", BASE)
    assert sorted((x["operacao"], x["codigo"], x["total"]) for x in r["resultado"]) == [("R", "k1c", 3.0), ("R", "k2c", 6.0)]
    assert calcular("3-N1C", op="*I")[0] == {"k1c": 3.0, "k2c": 6.0}


def test_composto_por_projeto():
    base = [L("N1=N1(tipo)"), L("N1(tipo=C)", "N1C_", "k1c"), L("T3", "T3_", "t3"), L("G:=N1(C)"), L("G:=T3", proj="PE")]
    assert calcular("1-G", base, projeto="PE")[0] == {"t3": 1.0}                           # o do projeto vence o genérico
    assert calcular("1-G", base, projeto="RO")[0] == {"k1c": 1.0}


def test_ciclo_no_calculo_avisa_em_vez_de_travar():
    base = [L("T3", "T3_", "t3"), L("P:=Q + T3"), L("Q:=P")]
    tot, nao, avisos = calcular("1-P", base)
    assert tot == {"t3": 1.0} and any("ciclo" in a for a in avisos)


def test_composto_com_item_inexistente_cai_em_nao_encontrado():
    base = [L("T3", "T3_", "t3"), L("Z:=X9 + T3")]
    tot, nao, _ = calcular("2-Z", base)
    assert tot == {"t3": 2.0} and nao == ["X9"]


def test_sem_composto_nada_muda():
    base = [L("T3", "T3_", "t3")]
    assert calcular("2-T3", base)[0] == {"t3": 2.0}


def test_totalizadora_acha_a_descricao_do_composto():
    r = montar_totalizadora([], [{"entidade": "0", "operacao": "I", "ativo": "2-N1C"}], [], BASE, "")
    assert [(x["ativo"], x["qtd"], x["naoEncontrado"]) for x in r] == [("N1C", 2.0, False)]


# ── paridade com o JS ──
pytestmark_node = pytest.mark.skipif(not o.NODE, reason="node não instalado")

CELULAS = ["N1C:=N1(C) + 2*N2(C)", "KIT := N1C + 0.5*T3, K", "[S 1]:=2*[CAZ 7,94] + N1(poste=DT)", "N1C:=", "N1C:=N1(C)+", "N1C:=0*A", "N1C:=N1=N1(a)", "A B:=X",
           "N1C:=N1(C", ":=A", "N1=N1(a)", "N1(tipo=C), N1C", "[A,B]", "T3", "KAB:=3*[A,B]"]


@pytestmark_node
def test_paridade_js():
    esperado = [{"cel": ap.analisar_celula_base(t), "nome": ap.nome_base(t)} for t in CELULAS]
    r = subprocess.run([o.NODE, "-e", f"const A=require({json.dumps(JS)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify(e.map(t=>({{cel:A.analisarCelulaBase(t),nome:A.nomeBase(t)}}))));"],
                       input=json.dumps(CELULAS), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    assert json.loads(r.stdout) == json.loads(json.dumps(esperado))


# ── importação ──
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


def adm(): return {"Authorization": f"Bearer {create_jwt_token('u1', 'a@x.com', 'admin')}"}


CAB = "ATIVO;DESC ATIVO;COMPONENTE;PROJETO;MDO;CODIGO;DESC CODIGO;FATOR I;FATOR R;FILTRO"


def csv_(*linhas):
    return (CAB + "\n" + "\n".join(linhas) + "\n").encode("utf-8-sig")


def test_importar_base_com_composto(client):
    dados = csv_("T3;d;C1;P;MATERIAL;t3;d;1;1;", "KIT:=2*T3;kit de teste;;P;;;;;;")
    for rota in ("/api/admin/upload-master", "/api/orcamento/upload"):
        r = client.post(rota, files={"file": ("b.csv", dados, "text/csv")}, headers=adm())
        assert r.status_code == 200, (rota, r.text)
    con = database.get_connection()
    assert con.execute("SELECT COUNT(*) FROM tabela_orcamento_master WHERE ativo = 'KIT:=2*T3'").fetchone()[0] == 1
    con.close()


@pytest.mark.parametrize("linha, trecho", [("KIT:=X9;;;P;;;;;;", "não existe"), ("KIT:=T3;;C9;P;;k;;;;", "em branco"), ("KIT:=T3 +;;;P;;;;;;", "item vazio")])
def test_importar_recusa_composto_ruim(client, linha, trecho):
    dados = csv_("T3;d;C1;P;MATERIAL;t3;d;1;1;", linha)
    r = client.post("/api/admin/upload-master", files={"file": ("b.csv", dados, "text/csv")}, headers=adm())
    assert r.status_code == 400 and trecho in r.json()["detail"], r.text
    r = client.post("/api/orcamento/upload", files={"file": ("b.csv", dados, "text/csv")}, headers=adm())
    assert r.status_code == 400 and trecho in r.json()["error"], r.text
