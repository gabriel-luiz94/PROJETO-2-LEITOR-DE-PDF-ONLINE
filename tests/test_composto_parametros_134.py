"""TASK-134 — composto com PARÂMETROS (``CF(cabo, cabo2=S04):=…``) e GRUPO que chama ativos (``CHAVE_FACA(cabo=%S04%):=TERM(cabo)``): leitura, assinaturas, validação, cálculo, vínculo e JS."""
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
from services.anotacao_outros import contexto_da_base, interpretar_linha
from tests import oraculo_tela as o

JS = os.path.join(config.STATIC_DIR, "ativo_parametros.js")
JS_AV = os.path.join(config.STATIC_DIR, "ativo_avaliar.js")


def L(ativo, comp="", cod="", fi=1.0, proj="PERNAMBUCO", **k):
    r = {"ativo": ativo, "desc_ativo": "d", "componente": comp, "projeto": proj, "mdo": "MATERIAL", "codigo": cod, "desc_codigo": f"d {cod}",
         "fator_i": fi, "fator_r": 0.0, "filtro": "", "origem": "OUTROS", "lv": "A", "arred": "NORMAL", "formula_i": "", "formula_r": ""}
    r.update(k)
    return r


BASE = [
    L("CHAVE_FACA=CHAVE_FACA(cabo=S04,cabo2=S04), SUH"), L("CHAVE_FACA", "CHAVE_FACA", "500109", 3),
    L("CHAVE_FACA(cabo=%S04%,cabo2=%S04%)", "CF.S04S04", "2401000", 6), L("CHAVE_FACA(cabo=%A35P%,cabo2=%A35P%)", "CF.A35A35", "2461009", 6),
    L("TERM=TERM(cabo=S04)"), L("TERM(cabo=%S04%)", "T.S04", "2418089", 6), L("TERM(cabo=%A185P%|%S336%)", "T.A185", "2418094", 6), L("TERM(cabo=%A35P%)", "T.A35", "2418099", 6),
    L("CON_S04", "CON_S04", "9999999", 2),
    # opção 2: grupos que chamam ativos
    L("CHAVE_FACA(cabo=%S04%):=TERM(cabo)", proj=""), L("CHAVE_FACA(cabo2=%S04%):=TERM(cabo2)"), L("CHAVE_FACA(cabo=%A35P%):=TERM(cabo)"), L("CHAVE_FACA(cabo2=%A35P%):=TERM(cabo2)"),
    # opção 1: compostos com parâmetros
    L("CFX(cabo, cabo2=S04):=CHAVE_FACA(cabo,cabo2) + 3*CON_S04, CFA"), L("CFY(cabo):=TERM(cabo) + TERM(cabo)"), L("CFZ(cabo):=CFY(cabo)"),
]
BASE = ap.preencher_componentes(BASE)[0]


def total(texto, base=None, op="I"):
    r = calc.processar_calculo([], [{"ativo": texto, "operacao": op, "entidade": "0"}], "PERNAMBUCO", base or BASE)
    return {x["codigo"]: x["total"] for x in r["resultado"]}, r["nao_encontrados"], r.get("avisos")


# ───────────────────────── leitura da célula ─────────────────────────
def test_cabecalho_do_composto():
    a = ap.analisar_celula_base("CF(cabo, cabo2=S04):=CHAVE_FACA(cabo,cabo2) + TERM(cabo)")
    assert (a["tipo"], a["base"], a["erro"], a["cabecalho"]) == ("composto", "CF", None, "CABO,CABO2=S04")
    assert a["formais"] == [{"nome": "CABO", "padrao": None}, {"nome": "CABO2", "padrao": "S04"}]
    assert a["itens"] == [{"ativo": "CHAVE_FACA(CABO,CABO2)", "qtd": 1.0}, {"ativo": "TERM(CABO)", "qtd": 1.0}]
    g = ap.analisar_celula_base("CHAVE_FACA(cabo=%S04%|%C120%):=2*TERM(cabo), K")
    assert g["base"] == "CHAVE_FACA" and g["formais"] is None and g["apelidos"] == ["K"]
    assert [(p["nome"], p["op"], p["alts"]) for p in g["condicoes"]] == [("CABO", "=", ["%S04%", "%C120%"])]
    assert ap.nome_base("CF(cabo,cabo2):=X(cabo), CFA") == "CF"


@pytest.mark.parametrize("ruim", ["CF():=X", "CF(cabo:=X", "CF(cabo=):=X", "CF(cabo):=", "CF(cabo):=TERM(cabo", "CF(%):=X", "A B(c):=X"])
def test_cabecalho_invalido(ruim):
    assert ap.analisar_celula_base(ruim)["erro"], ruim


def test_composto_sem_parenteses_nao_mudou():
    a = ap.analisar_celula_base("N1C:=N1(C) + 2*N2(C)")
    assert (a["cabecalho"], a["formais"], a["condicoes"]) == (None, None, None) and a["base"] == "N1C"


def test_assinaturas_do_composto_e_do_grupo():
    sig = ap.montar_assinaturas([r["ativo"] for r in BASE])
    assert sig["CFX"]["nomes"] == ["CABO", "CABO2"] and sig["CFX"]["padroes"] == {"CABO2": "S04"} and sig["CFX"]["composto"] and sig["CFX"]["fechados"] == []
    assert sig["CHAVE_FACA"]["nomes"] == ["CABO", "CABO2"] and "composto" not in sig["CHAVE_FACA"]       # o grupo só soma condições ao ativo
    assert ap.montar_assinaturas(["N1=N1(a)", "N1(fio=%X%):=T"])["N1"]["nomes"] == ["A"]                   # parâmetro desconhecido não entra nos nomes


# ───────────────────────── cálculo: opção 2 (grupo que chama ativos) ─────────────────────────
def test_grupo_chama_o_terminal_de_cada_cabo():
    t, nao, av = total("1-CHAVE_FACA")                        # padrão S04,S04: chave + conetor S04/S04 + terminal S04 do cabo e do cabo2
    assert t == {"500109": 3.0, "2401000": 6.0, "2418089": 12.0} and nao == []
    t, _, _ = total("1-CHAVE_FACA(A35P,A35P)")
    assert t == {"500109": 3.0, "2461009": 6.0, "2418099": 12.0}
    t, _, _ = total("2-CHAVE_FACA(S04,A35P)")                 # cada lado chama o terminal do seu cabo; quantidade multiplica
    assert t == {"500109": 6.0, "2418089": 12.0, "2418099": 12.0}


def test_grupo_chamado_pelo_apelido_e_sem_regra_casada():
    assert total("1-SUH(S04,A35P)")[0] == total("1-CHAVE_FACA(S04,A35P)")[0]
    t, nao, _ = total("1-CHAVE_FACA(S04,S1/0)")               # S1/0 não está cadastrado em nenhum grupo: o pedido é recusado como sempre foi
    assert t == {} and any("não cadastrado" in n for n in nao)


def test_grupo_chama_outro_ativo_composto_e_o_valor_entra_inteiro():
    base = BASE + [L("KIT:=2*CON_S04")] + [L("CHAVE_FACA(cabo=%C120%):=KIT")]
    base = ap.preencher_componentes(base)[0]
    base = [dict(r, ativo=r["ativo"]) for r in base]
    base.insert(1, L("CHAVE_FACA(cabo=%C120%,cabo2=%C120%)", "CF.C", "2400000", 1))
    base = ap.preencher_componentes(base)[0]
    t, nao, _ = total("1-CHAVE_FACA(C120,C120)", base)
    assert t["9999999"] == 4.0 and "2400000" in t and nao == []


def test_grupo_com_ciclo_nao_trava():
    base = ap.preencher_componentes(BASE + [L("CHAVE_FACA(cabo=%S04%):=CHAVE_FACA(S04,S04)")])[0]
    t, nao, avisos = total("1-CHAVE_FACA", base)
    assert any("funda demais" in a for a in (avisos or []))


# ───────────────────────── cálculo: opção 1 (composto com parâmetros) ─────────────────────────
def test_composto_com_parametros_passa_os_valores_aos_itens():
    t, nao, _ = total("1-CFX(A35P)")                           # = CHAVE_FACA(A35P,S04) + 3*CON_S04 (cabo2 usa o padrão S04)
    assert t == {"500109": 3.0, "2418099": 6.0, "2418089": 6.0, "9999999": 6.0} and nao == []
    t, _, _ = total("1-CFX(A35P,A35P)")
    assert t == {"500109": 3.0, "2461009": 6.0, "2418099": 12.0, "9999999": 6.0}
    assert total("1-CFX(cabo2=A35P,cabo=S04)")[0] == total("1-CFX(S04,A35P)")[0]           # por nome, em qualquer ordem
    assert total("1-CFA(A35P,A35P)")[0] == total("1-CFX(A35P,A35P)")[0]                  # apelido do composto


def test_composto_dentro_de_composto_com_parametros():
    assert total("1-CFZ(S336)")[0] == {"2418094": 12.0}                                    # CFZ → CFY → 2 × TERM(S336)
    assert total("3-CFZ(S336)")[0] == {"2418094": 36.0}


def test_composto_parametro_faltando_ou_demais():
    t, nao, avisos = total("1-CFY")
    assert t == {} and any("faltou informar CABO" in a for a in avisos)
    t, nao, _ = total("1-CFX(a,b,c)")
    assert t == {} and any("parâmetros demais" in n for n in nao)


def test_sem_os_novos_compostos_nada_muda():
    antigo = [r for r in BASE if ":=" not in r["ativo"]]
    assert total("1-CHAVE_FACA", antigo)[0] == {"500109": 3.0, "2401000": 6.0}


def test_estrutura_com_cabos_vinculados_preenche_o_composto():
    ctx = contexto_da_base(BASE, "PERNAMBUCO")
    ctx["parametros_estrutura"] = {"CFX": {"CABO": ["A35P"], "CABO2": ["A35P"]}}
    r = interpretar_linha("9400D, CFX", "I", lambda t: [], ctx)
    assert [i["ativo"] for i in r["itens"]][-1].replace(" ", "") in ("CFX(CABO=A35P,CABO2=A35P)", "CFX(CABO2=A35P,CABO=A35P)")


# ───────────────────────── validação da base ─────────────────────────
def erros(*extra):
    return ap.validar_base([dict(r) for r in ap.preencher_componentes(BASE_LIMPA + list(extra))[0]])["erros"]


BASE_LIMPA = [r for r in BASE if ":=" not in r["ativo"]]


def test_base_com_os_novos_compostos_valida():
    assert ap.validar_base([dict(r) for r in BASE])["erros"] == []


@pytest.mark.parametrize("linha,trecho", [
    (L("CHAVE_FACA(fio=%S04%):=TERM(cabo)"), "FIO não está na declaração"),
    (L("CHAVE_FACA(cabo=%S04%):=NAOEXISTE(cabo)"), "não existe na base"),
    (L("CHAVE_FACA(cabo=%S04%):=TERM(cabo,cabo2)"), "parâmetros demais para TERM"),
    (L("CFQ(cabo=%S04%):=TERM(cabo)"), "não é um ativo da base"),
    (L("CHAVE_FACA(cabo=%S04%):=CHAVE_FACA(cabo,cabo2)"), "ciclo"),
    (L("CHAVE_FACA(cabo=%S04%):=TERM(cabo)", "CF.X", "123"), "só uma definição"),
    (L("CHAVE_FACA:=TERM(S04)"), "já existe como ativo da base"),
])
def test_validacao_recusa(linha, trecho):
    assert any(trecho in e for e in erros(linha)), erros(linha)


def test_validacao_duplicados():
    e = erros(L("CHAVE_FACA(cabo=%S04%):=TERM(cabo)"), L("CHAVE_FACA(cabo=%S04%):=TERM(cabo2)"), L("CFQ(a):=TERM(a)"), L("CFQ(b):=TERM(b)"))
    assert sum("já foi definido" in x for x in e) == 2


def test_grupo_antes_da_declaracao_nao_quebra_a_regra_da_primeira_linha():
    linhas = ap.preencher_componentes([L("CHAVE_FACA(cabo=%S04%):=TERM(cabo)")] + BASE_LIMPA)[0]
    assert ap.validar_base([dict(r) for r in linhas])["erros"] == []


def test_composto_nao_ganha_componente():
    for r in ap.preencher_componentes(BASE)[0]:
        if ":=" in r["ativo"]:
            assert not r["componente"], r["ativo"]


# ───────────────────────── JS: paridade e tela ─────────────────────────
CELULAS = ["CF(cabo, cabo2=S04):=CHAVE_FACA(cabo,cabo2) + TERM(cabo)", "CHAVE_FACA(cabo=%S04%|%C120%):=2*TERM(cabo), K", "CHAVE_FACA(cabo!=S04):=Y", "CF():=X", "CF(cabo:=X",
           "CF(cabo=):=X", "KIT(cabo=%S04%,cabo2=S1/0):=2*X(cabo)", "N1C:=N1(C)+2*N2(C)", "CF(cabo=%S04%):=X", "CHAVE_FACA(cabo=3XS04-7,97kV):=Z"]


@pytest.mark.skipif(not o.NODE, reason="node não instalado")
def test_paridade_js_celulas_e_assinaturas():
    esperado = {"cel": [ap.analisar_celula_base(t) for t in CELULAS], "nome": [ap.nome_base(t) for t in CELULAS], "sig": ap.montar_assinaturas([r["ativo"] for r in BASE])}
    cod = f"const A=require({json.dumps(JS)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({{cel:e.cel.map(t=>A.analisarCelulaBase(t)),nome:e.cel.map(t=>A.nomeBase(t)),sig:A.montarAssinaturas(e.celulas)}}));"
    r = subprocess.run([o.NODE, "-e", cod], input=json.dumps({"cel": CELULAS, "celulas": [x["ativo"] for x in BASE]}), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    assert json.loads(r.stdout) == json.loads(json.dumps(esperado))


@pytest.mark.skipif(not o.NODE, reason="node não instalado")
def test_totalizadora_valida_o_composto_com_parametros_e_o_grupo_normal():
    cod = f"""const AV=require({json.dumps(JS_AV)});const d=JSON.parse(require('fs').readFileSync(0,'utf8'));const idx=AV.construirIndice(d.base,'PERNAMBUCO');
console.log(JSON.stringify(d.casos.map(c=>AV.avaliarAtivo(c,'OUTROS',idx))));"""
    casos = ["CFX(A35P)", "CFX(A35P,A35P)", "CFX(a,b,c)", "CHAVE_FACA(S04,S04)", "CHAVE_FACA(S04,XX)", "NAOEXISTE"]
    r = subprocess.run([o.NODE, "-e", cod], input=json.dumps({"base": BASE, "casos": casos}), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    res = json.loads(r.stdout)
    assert [bool(x["naoEncontrado"]) for x in res] == [False, False, False, False, False, True]
    assert res[0]["erroParametros"] == "" and res[1]["erroParametros"] == ""
    assert "parâmetros demais" in res[2]["erroParametros"]
    assert res[3]["erroParametros"] == "" and "XX" in res[4]["erroParametros"]               # o grupo que chama ativos NÃO desliga a validação do ativo


# ───────────────────────── importação ─────────────────────────
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


CAB = "ATIVO;DESC ATIVO;COMPONENTE;PROJETO;MDO;CODIGO;DESC CODIGO;FATOR I;FATOR R;FILTRO"


def csv_(*linhas):
    return (CAB + "\n" + "\n".join(linhas) + "\n").encode("utf-8-sig")


def adm():
    return {"Authorization": f"Bearer {create_jwt_token('u1', 'a@x.com', 'admin')}"}


def test_importar_base_com_os_novos_compostos(client):
    dados = csv_("T=T(cabo=S04);d;;P;MATERIAL;;;0;0;", "T(cabo=%S04%);d;TS;P;MATERIAL;t1;d;1;0;", "CH=CH(cabo=S04);d;;P;MATERIAL;;;0;0;", "CH;d;CHC;P;MATERIAL;c1;d;1;0;",
                 "CH(cabo=%S04%):=T(cabo);;;P;;;;;;", "CFW(cabo):=CH(cabo) + T(cabo);;;P;;;;;;")
    for rota in ("/api/admin/upload-master", "/api/orcamento/upload"):
        r = client.post(rota, files={"file": ("b.csv", dados, "text/csv")}, headers=adm())
        assert r.status_code == 200, (rota, r.text)


def test_importar_recusa_composto_com_parametro_desconhecido(client):
    dados = csv_("CH=CH(cabo=S04);d;;P;MATERIAL;;;0;0;", "CH;d;CHC;P;MATERIAL;c1;d;1;0;", "CH(fio=%S04%):=CH(S04);;;P;;;;;;")
    r = client.post("/api/admin/upload-master", files={"file": ("b.csv", dados, "text/csv")}, headers=adm())
    assert r.status_code == 400 and "FIO" in r.json()["detail"], r.text
    r = client.post("/api/orcamento/upload", files={"file": ("b.csv", dados, "text/csv")}, headers=adm())
    assert r.status_code == 400 and "FIO" in r.json()["error"], r.text
