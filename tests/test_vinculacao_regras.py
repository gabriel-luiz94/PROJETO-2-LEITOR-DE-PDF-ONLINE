"""TASK-084 — regras de vinculação cabo <-> estrutura (Tabela de Cabos + Tabela de Estruturas): lógica pura Python = JS."""
import json
import os
import random
import shutil
import subprocess

import pytest

import config
from services import vinculacao_regras as vr

NODE = shutil.which("node")

CFG = {
    "versao": 2,
    "cabos": [
        {"padrao": "^C\\d+$", "grupo": "Cobre nu", "valor_cabo": "C", "operacoes": ["I", "M"], "pontas": {"I": 2, "M": 1}, "estruturas": "^[NUB][1-4]"},
        {"padrao": "^(A|S)\\d+P$", "grupo": "Protegido", "valor_cabo": "P", "operacoes": ["I", "M"], "estruturas": "^N[1-4]"},
        {"padrao": "^(S|AA)\\d+(/\\d+)?$", "grupo": "Alumínio nu", "valor_cabo": "A", "estruturas": "^[NUT][1-4]"},
        {"padrao": "RAMA25, PB35, PB70, PB120", "grupo": "Multiplex BT", "valor_cabo": "", "operacoes": ["I"], "estruturas": "B2, R2"},
        {"padrao": "^CAZ", "grupo": "Neutro", "vincula": False},
    ],
    "estruturas": [
        {"padrao": "^N1$", "qtd_cabos": 2, "lados": "IGUAIS", "amarracao": False, "se_diferem": "BLOQUEAR"},
        {"padrao": "^N2$", "qtd_cabos": 2, "lados": "LIVRE", "amarracao": False, "se_diferem": "MAIOR"},
        {"padrao": "^N3$", "qtd_cabos": 1, "lados": "LIVRE", "amarracao": True},
        {"padrao": "^T3$", "qtd_cabos": 1, "amarracao": True, "aceita_i": False, "cabos": "^(S|A)\\d+"},
        {"padrao": "B2, R2", "qtd_cabos": 2, "lados": "IGUAIS", "cabos": "RAMA25, PB35, PB70, PB120", "parametro": ""},
    ],
    "ordem_maior": ["C", "P", "A"],
}


def primeiro(txt):
    return (txt or "").split(" ")[0].upper()


def cab(ativo, op="I", viculos=None):
    d = {"entidade": "CABO", "operacao": op, "ativo": ativo}
    if viculos is not None:
        d["vinculoEstruturas"] = viculos
    return d


def est(ativo, ficticios=None):
    d = {"entidade": "ESTRUTURA", "operacao": "I", "ativo": ativo}
    if ficticios:
        d["cabosFicticios"] = ficticios
    return d


def analisa(cabos, outros, cfg=CFG):
    return vr.analisar(cabos, outros, cfg, primeiro, primeiro)


# ── casa ──
def test_casa_regex_lista_e_vazio():
    assert vr.casa("", "X") and vr.casa(None, "X") and vr.casa([], "X")
    assert vr.casa("N1, N2", "n2") and not vr.casa("N1, N2", "N3")
    assert vr.casa("^N\\d$", "N3") and not vr.casa("^N\\d$", "T3") and not vr.casa("^(", "x")      # regex inválida não casa
    assert vr.casa(["A", "b"], "B")


# ── pode vincular ──
def test_cabo_e_estrutura_precisam_se_aceitar():
    c16 = vr.info_cabo("C16", "I", vr.normalizar_config(CFG))
    n1 = vr.info_estrutura("N1", vr.normalizar_config(CFG))
    assert vr.pode_vincular(c16, n1) == (True, None)
    b2 = vr.info_estrutura("B2", vr.normalizar_config(CFG))
    assert vr.pode_vincular(c16, b2)[0] is False                       # o cabo não aceita B2
    rama = vr.info_cabo("RAMA25", "I", vr.normalizar_config(CFG))
    assert vr.pode_vincular(rama, b2) == (True, None)
    assert vr.pode_vincular(rama, n1)[0] is False
    caz = vr.info_cabo("CAZ7,94", "I", vr.normalizar_config(CFG))
    assert vr.pode_vincular(caz, n1)[0] is False                       # neutro não vincula


def test_cabo_m_so_em_amarracao_e_t3_so_aceita_m():
    cfg = vr.normalizar_config(CFG)
    s10 = vr.info_cabo("S1/0", "M", cfg)
    assert vr.pode_vincular(s10, vr.info_estrutura("N1", cfg))[0] is False       # N1 não é amarração
    assert vr.pode_vincular(s10, vr.info_estrutura("T3", cfg)) == (True, None)
    assert vr.pode_vincular(vr.info_cabo("S1/0", "I", cfg), vr.info_estrutura("T3", cfg))[0] is False      # T3 só cabo M
    assert vr.pode_vincular(vr.info_cabo("C16", "M", cfg), vr.info_estrutura("T3", cfg))[0] is False       # T3 não aceita C16


def test_lados_iguais_e_capacidade():
    cfg = vr.normalizar_config(CFG)
    n1 = vr.info_estrutura("N1", cfg)
    c16, a35 = vr.info_cabo("C16", "I", cfg), vr.info_cabo("A35P", "I", cfg)
    assert vr.pode_com_ocupantes(c16, n1, [{"info": c16, "ficticio": False}])[0]
    assert vr.pode_com_ocupantes(a35, n1, [{"info": c16, "ficticio": False}])[0] is False       # N1 exige iguais
    assert vr.pode_com_ocupantes(c16, n1, [{"info": c16, "ficticio": False}, {"info": c16, "ficticio": False}])[0] is False    # capacidade 2
    n2 = vr.info_estrutura("N2", cfg)
    assert vr.pode_com_ocupantes(a35, n2, [{"info": c16, "ficticio": False}])[0]                # N2 aceita lados diferentes


# ── estados ──
def test_estados_de_cabos_e_estruturas():
    cabos = [cab("C16", "I", [0, 1]), cab("A35P", "I", [2]), cab("S1/0", "M", [3]), cab("RAMA25", "I"), cab("C16", "I", [0, 2])]
    outros = [est("N1"), est("N3"), est("N1"), est("T3"), est("B2")]
    a = analisa(cabos, outros)
    assert {k: v["estado"] for k, v in a["estruturas"].items()} == {0: "gr", 1: "gr", 2: "rd", 3: "gr", 4: "k"}
    assert {k: v["estado"] for k, v in a["cabos"].items()} == {0: "gr", 1: "pu", 2: "gr", 3: "k", 4: "rd"}      # o cabo que rompe o "lados iguais" é o inválido
    assert a["estruturas"][0]["valor_cabo"] == "C" and a["estruturas"][3]["valor_cabo"] == "A"
    assert any(not l["ok"] for l in a["links"])


def test_parcial_roxo_e_ficticio_completa():
    cabos = [cab("C16", "I", [0, 1])]
    outros = [est("N3"), est("N1")]
    a = analisa(cabos, outros)
    assert a["estruturas"][1]["estado"] == "pu" and a["cabos"][0]["estado"] == "gr"
    outros[1]["cabosFicticios"] = ["C16"]
    a = analisa(cabos, outros)
    assert a["estruturas"][1]["estado"] == "gr" and a["estruturas"][1]["ficticios"] == 1 and a["estruturas"][1]["valor_cabo"] == "C"
    outros[1]["cabosFicticios"] = ["A35P"]
    assert analisa(cabos, outros)["estruturas"][1]["estado"] == "rd"                 # N1 exige lados iguais: fictício diferente é inválido


def test_valor_do_parametro_quando_os_lados_diferem():
    cabos = [cab("C16", "I", [0]), cab("A35P", "I", [0])]
    r = analisa(cabos, [est("N2")])["estruturas"][0]
    assert r["valor_cabo"] == "C" and r["avisos"]                                    # N2: usa o maior (C > P > A) e avisa
    cfg = json.loads(json.dumps(CFG)); cfg["estruturas"][1]["se_diferem"] = "AVISAR"
    assert vr.analisar(cabos, [est("N2")], cfg, primeiro, primeiro)["estruturas"][0]["valor_cabo"] is None
    assert vr.parametros_por_linha(cabos, [est("N2")], CFG, primeiro, primeiro) == {0: {"CABO": "C"}}


def test_sem_regra_nao_bloqueia_e_sem_valor_nao_preenche():
    a = vr.analisar([cab("XYZ9", "I", [0])], [est("U9")], vr.config_vazia(), primeiro, primeiro)
    assert a["estruturas"][0]["estado"] == "gr" and a["estruturas"][0]["valor_cabo"] is None and a["cabos"][0]["estado"] == "pu"     # sem regra: cabo I precisa de 2 pontas
    assert vr.parametros_por_linha([cab("RAMA25", "I", [0])], [est("B2")], CFG, primeiro, primeiro) == {}          # multiplex não carrega valor


def test_achados_de_validacao():
    cabos = [cab("C16", "I", [0]), cab("A35P", "I", [0])]
    ach = vr.achados(cabos, [est("N1"), est("N1")], CFG, primeiro, primeiro)
    msgs = " | ".join(a["mensagem"] for a in ach)
    assert "vínculo inválido" in msgs and "sem nenhum cabo vinculado" in msgs
    assert all(a["regra_id"] == "VINCULO-ESTRUTURA" and a["severidade"] == "aviso" for a in ach)
    # a exceção "existe cabo M" cala o aviso de estrutura sem cabo
    ach2 = vr.achados(cabos + [cab("S1/0", "M")], [est("N1"), est("N1")], CFG, primeiro, primeiro)
    assert not any("sem nenhum cabo" in a["mensagem"] for a in ach2)


# ── legado e validação ──
def test_migracao_das_regras_antigas():
    legado = [{"tipo_estrutura": "1", "qtd_cabos": 2, "compatibilidade": "MESMO_TIPO_FASE_OPERACAO", "descricao": "d"},
              {"tipo_estrutura": "3", "qtd_cabos": 1, "compatibilidade": "LIVRE"}]
    cfg = vr.config_de_legado(legado)
    assert vr.validar_config(cfg) == []
    n1, u3, u4 = (vr.info_estrutura(c, cfg) for c in ("N1", "U3", "U4"))
    assert (n1["qtd"], n1["lados"]) == (2, "IGUAIS") and (u3["qtd"], u3["lados"]) == (1, "LIVRE") and u4["regra"] is None      # o 1º dígito do código decide


def test_validar_config():
    assert vr.validar_config(CFG) == []
    assert vr.validar_config({"cabos": [{"padrao": "^("}]})
    assert vr.validar_config({"estruturas": [{"padrao": "N1", "qtd_cabos": 0}]})
    assert vr.validar_config({"estruturas": [{"padrao": "N1", "lados": "X"}]})
    assert vr.validar_config({"cabos": [{"padrao": "C16", "pontas": {"I": -1}}]})
    assert vr.validar_config([]) and vr.validar_config({"ordem_maior": "C"})


# ── paridade Python x JS ──
_JS = """
const fs = require('fs'); const V = require(%(v)s);
const e = JSON.parse(fs.readFileSync(0, 'utf8'));
const f = t => String(t == null ? '' : t).split(' ')[0].toUpperCase();
console.log(JSON.stringify(e.casos.map(c => ({ a: V.analisar(c.cabos, c.outros, c.cfg, f, f), p: V.parametrosPorLinha(c.cabos, c.outros, c.cfg, f, f), x: V.achados(c.cabos, c.outros, c.cfg, f, f) }))));
"""


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_python_js_em_casos_aleatorios():
    rnd = random.Random(84)
    cods_c = ["C16", "C35", "A35P", "S50P", "S1/0", "AA4/0", "RAMA25", "PB35", "CAZ7,94", "XYZ", ""]
    cods_e = ["N1", "N2", "N3", "T3", "B2", "R2", "U4", "N1(XLPE,DT)", ""]
    casos = []
    for _ in range(300):
        cabos = [cab(rnd.choice(cods_c) + " 1 20", rnd.choice(["I", "*I", "R", "M", "*M"])) for _ in range(rnd.randint(0, 6))]
        outros = []
        for _ in range(rnd.randint(0, 6)):
            o = est(rnd.choice(cods_e))
            o["entidade"] = rnd.choice(["ESTRUTURA", "POSTE", "ESTRUTURA", "0"])
            if rnd.random() < .25:
                o["cabosFicticios"] = [rnd.choice(["C16", "A35P", "RAMA25"])]
            outros.append(o)
        for c in cabos:
            if outros and rnd.random() < .8:
                c["vinculoEstruturas"] = sorted(rnd.sample(range(len(outros)), rnd.randint(1, min(2, len(outros)))))
        cfg = json.loads(json.dumps(CFG))
        cfg["estruturas"][rnd.randrange(len(cfg["estruturas"]))]["se_diferem"] = rnd.choice(["BLOQUEAR", "MAIOR", "AVISAR"])
        if rnd.random() < .2:
            cfg = vr.config_vazia()
        casos.append({"cabos": cabos, "outros": outros, "cfg": cfg})
    r = subprocess.run([NODE, "-e", _JS % {"v": json.dumps(os.path.join(config.STATIC_DIR, "vinculacao_regras.js"))}], input=json.dumps({"casos": casos}),
                       capture_output=True, text=True, timeout=120)
    assert r.returncode == 0, r.stderr
    js = json.loads(r.stdout)
    for c, j in zip(casos, js):
        py = {"a": vr.analisar(c["cabos"], c["outros"], c["cfg"], primeiro, primeiro), "p": vr.parametros_por_linha(c["cabos"], c["outros"], c["cfg"], primeiro, primeiro),
              "x": vr.achados(c["cabos"], c["outros"], c["cfg"], primeiro, primeiro)}
        assert json.loads(json.dumps(py)) == j


# ── parâmetro `cabo` na Totalizadora (JS real = espelho Python) ──
from services.autonomo import montagem as mt
from services.autonomo import totalizadora as tt
from tests import oraculo_tela as o


def _l(ativo, comp, cod, proj="P1", desc="D"):
    return {"ativo": ativo, "desc_ativo": desc if ativo else "", "componente": comp, "projeto": proj, "codigo": cod, "desc_codigo": "d"}


BASE_N1 = [_l("N1=N1(cabo=A, poste=11300D)", "N1", "BASE"), _l("N1(cabo=A)", "N1.CABO=A", "FIOA"), _l("N1(cabo=P)", "N1.CABO=P", "FIOP"), _l("N1(cabo=C)", "N1.CABO=C", "FIOC"),
           _l("N3=N3(cabo=A)", "N3", "N3BASE"), _l("N3(cabo=C)", "N3.CABO=C", "N3C"), _l("DT9/200, 9200D", "DT9/200", "POSTE9")]
CFG_P1 = {"versao": 2, "cabos": [{"padrao": "^C\\d+$", "valor_cabo": "C"}, {"padrao": "^A\\d+P$", "valor_cabo": "P"}], "estruturas": [
    {"padrao": "^N1$", "qtd_cabos": 2, "lados": "IGUAIS"}, {"padrao": "^N3$", "qtd_cabos": 1}]}


def _cenario():
    cabos = [{"entidade": "CABO", "operacao": "I", "ativo": "3xC16(16) 30 m", "vinculoEstruturas": [0, 1]}, {"entidade": "CABO", "operacao": "I", "ativo": "3xC16(16) 28 m", "vinculoEstruturas": [0]}]
    outros = [{"entidade": "POSTE", "operacao": "I", "ativo": "9200D, N1"}, {"entidade": "ESTRUTURA", "operacao": "I", "ativo": "1-N3"},
              {"entidade": "ESTRUTURA", "operacao": "I", "ativo": "1-N1(P)"}, {"entidade": "ESTRUTURA", "operacao": "I", "ativo": "1-N1"}]
    mt.recalcular_qtd_ativos(cabos)
    return cabos, outros


def _codigos(tot):
    return sorted(i["ativo"] for i in tot if i["origem"] == "OUTROS")


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_totalizadora_preenche_o_parametro_cabo_igual_ao_js():
    cabos, outros = _cenario()
    js = o.totalizadora_e_payload(cabos, outros, [], BASE_N1, "P1", cfg_vinculo=CFG_P1)
    tot = tt.montar_totalizadora(cabos, outros, [], BASE_N1, "P1", cfg_vinculo=CFG_P1)
    assert tt.normalizar_para_json(tot) == js["totalizadora"] and tt.payload_calculo(tot) == js["payload"]
    ativos = _codigos(tot)
    assert any("CABO=C" in a and "9200D" in a for a in ativos)         # 9200D, N1 ligado a 2 cabos C16 → cabo=C (e poste 9200D)
    assert "N3(C)" in ativos                                           # N3 com 1 cabo C16
    assert any(a.startswith("N1(P") for a in ativos)                   # valor digitado vence (N1(P))
    assert sum(1 for a in ativos if a.startswith("N1(A,")) == 1         # N1 sem vínculo usa o padrão da declaração (A)


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_sem_configuracao_nada_muda():
    cabos, outros = _cenario()
    js = o.totalizadora_e_payload(cabos, outros, [], BASE_N1, "P1")
    tot = tt.montar_totalizadora(cabos, outros, [], BASE_N1, "P1")
    assert tt.normalizar_para_json(tot) == js["totalizadora"] and tt.payload_calculo(tot) == js["payload"]
    assert not any(a.startswith("N1(C") for a in _codigos(tot))


# ── script de vinculação ──
def _xy(d, x, y):
    d["_x"], d["_y"] = x, y
    return d


def _scr(*passos, conflitos="avisar"):
    return {"versao": 1, "conflitos": conflitos, "passos": list(passos)}


def roda(script, cabos, outros, aplicar=False, cfg=CFG):
    return vr.executar_script(script, cabos, outros, cfg, primeiro, primeiro, aplicar)


def test_script_proximidade_liga_as_duas_estruturas_mais_proximas_e_respeita_o_raio():
    cabos = [_xy(cab("C16", "I"), 0, 0)]
    outros = [_xy(est("N1"), 1, 0), _xy(est("N1"), 2, 0), _xy(est("N1"), 50, 0)]
    r = roda(_scr({"modo": "proximidade", "raio_m": 15}), cabos, outros, aplicar=True)
    assert cabos[0]["vinculoEstruturas"] == [0, 1] and cabos[0]["vinculoOrigem"] == "script"
    assert r["passos"][0]["completos"] == 1 and r["vinculos"][0]["completo"]
    cabos = [_xy(cab("C16", "I"), 0, 0)]
    r = roda(_scr({"modo": "proximidade", "raio_m": 1.5}), cabos, outros)
    assert r["vinculos"] == [] and r["sem_candidato"][0]["motivo"].startswith("sem estrutura compatível")


def test_script_parcial_so_com_aceitar_parcial():
    cabos, outros = [_xy(cab("C16", "I"), 0, 0)], [_xy(est("N1"), 1, 0)]
    assert roda(_scr({"modo": "proximidade"}), cabos, outros)["vinculos"] == []
    r = roda(_scr({"modo": "proximidade", "aceitar_parcial": True}), cabos, outros)
    assert r["vinculos"] == [{"passo": 0, "cabo": 0, "estruturas": [0], "completo": False}] and r["passos"][0]["parciais"] == 1


def test_script_cabo_m_so_na_amarracao_e_nao_sobrescreve_manual():
    cabos = [_xy(cab("S1/0", "M"), 0, 0), _xy(cab("S1/0", "M", [2]), 0, 0)]
    outros = [_xy(est("N1"), 1, 0), _xy(est("T3"), 5, 0), _xy(est("T3"), 6, 0)]
    r = roda(_scr({"modo": "proximidade", "cabos": {"operacao": ["M"]}, "estruturas": {"amarracao": True}}), cabos, outros)
    assert [(v["cabo"], v["estruturas"]) for v in r["vinculos"]] == [(0, [1])] and r["passos"][0]["ja_tinham"] == 1
    r = roda(_scr({"modo": "proximidade", "cabos": {"operacao": ["M"]}, "substituir": True}), cabos, outros)
    assert {v["cabo"] for v in r["vinculos"]} == {0, 1}


def test_script_conflito_lados_iguais_e_modos():
    cabos = [_xy(cab("C16", "I", [0]), 0, 0), _xy(cab("A35P", "I"), 1, 0)]
    outros = [_xy(est("N1"), 1, 0), _xy(est("N2"), 3, 0), _xy(est("N2"), 4, 0)]
    r = roda(_scr({"modo": "proximidade"}), cabos, outros)
    assert r["vinculos"][0]["estruturas"] == [1, 2] and "exige cabos iguais" in r["conflitos"][0]["motivo"]     # N1 recusou (A35P × C16): foi para as N2
    assert roda(_scr({"modo": "proximidade"}, conflitos="ignorar"), cabos, outros)["conflitos"] == []
    assert roda(_scr({"modo": "proximidade"}, conflitos="parar"), cabos, outros)["parou"] is True


def test_script_sequencia_multiplex():
    cabos = [cab("RAMA25", "I"), cab("RAMA25", "I")]
    outros = [est("B2"), est("B2"), est("B2")]
    r = roda(_scr({"modo": "sequencia", "cabos": {"ativo": "^RAMA"}, "estruturas": {"ativo": "^B2$"}}), cabos, outros)
    assert [(v["cabo"], v["estruturas"]) for v in r["vinculos"]] == [(0, [0, 1]), (1, [1, 2])]


def test_validar_script():
    assert vr.validar_script(_scr({"modo": "proximidade"})) == []
    assert vr.validar_script({"passos": []}) and vr.validar_script({}) and vr.validar_script([]) and vr.validar_script(_scr({"modo": "x"}))
    assert vr.validar_script(_scr({"modo": "proximidade", "raio_m": 0})) and vr.validar_script(_scr({"modo": "proximidade", "cabos": {"ativo": "^("}}))
    assert vr.validar_script({"conflitos": "x", "passos": [{"modo": "sequencia"}]})


_JS2 = """
const fs = require('fs'); const V = require(%(v)s);
const e = JSON.parse(fs.readFileSync(0, 'utf8'));
const f = t => String(t == null ? '' : t).split(' ')[0].toUpperCase();
console.log(JSON.stringify(e.casos.map(c => ({ s: V.executarScript(c.script, c.cabos, c.outros, c.cfg, f, f, true), c: c.cabos.map(x => x.vinculoEstruturas || null) }))));
"""


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_do_script_python_js():
    rnd = random.Random(842)
    cods_c = ["C16", "C35", "A35P", "S50P", "S1/0", "RAMA25", "PB35", "CAZ7,94", ""]
    cods_e = ["N1", "N2", "N3", "T3", "B2", "U4"]
    casos = []
    for _ in range(300):
        cabos = []
        for _ in range(rnd.randint(1, 6)):
            c = cab(rnd.choice(cods_c) + " 1 20", rnd.choice(["I", "*I", "R", "M", "*M"]))
            if rnd.random() < .85:
                c["_x"], c["_y"] = rnd.randint(0, 30) / 2, rnd.randint(0, 30) / 2
            cabos.append(c)
        outros = []
        for _ in range(rnd.randint(1, 7)):
            o = est(rnd.choice(cods_e)); o["entidade"] = rnd.choice(["ESTRUTURA", "POSTE", "0"])
            if rnd.random() < .9:
                o["_x"], o["_y"] = rnd.randint(0, 30) / 2, rnd.randint(0, 30) / 2
            if rnd.random() < .15:
                o["cabosFicticios"] = [rnd.choice(["C16", "RAMA25"])]
            outros.append(o)
        for c in cabos:
            if rnd.random() < .2:
                c["vinculoEstruturas"] = [rnd.randrange(len(outros))]
        passos = []
        for _ in range(rnd.randint(1, 3)):
            p = {"modo": rnd.choice(["proximidade", "sequencia"]), "aceitar_parcial": rnd.random() < .4, "substituir": rnd.random() < .2}
            if rnd.random() < .5:
                p["raio_m"] = rnd.choice([1, 3, 6, 100])
            if rnd.random() < .4:
                p["cabos"] = {"operacao": rnd.sample(["I", "M", "*I", "R"], rnd.randint(1, 2))}
            if rnd.random() < .3:
                p["estruturas"] = {"ativo": rnd.choice(["^N", "B2, N2", "^T3$"]), "amarracao": rnd.choice([True, False])}
            if rnd.random() < .3:
                p["pontas"] = rnd.randint(1, 3)
            passos.append(p)
        casos.append({"script": {"conflitos": rnd.choice(["avisar", "ignorar", "parar"]), "passos": passos}, "cabos": cabos, "outros": outros, "cfg": CFG})
    r = subprocess.run([NODE, "-e", _JS2 % {"v": json.dumps(os.path.join(config.STATIC_DIR, "vinculacao_regras.js"))}], input=json.dumps({"casos": casos}),
                       capture_output=True, text=True, timeout=120)
    assert r.returncode == 0, r.stderr
    js = json.loads(r.stdout)
    for c, j in zip(casos, js):
        import copy
        cabos = copy.deepcopy(c["cabos"])
        res = vr.executar_script(c["script"], cabos, c["outros"], c["cfg"], primeiro, primeiro, True)
        py = {"s": res, "c": [x.get("vinculoEstruturas") for x in cabos]}
        assert json.loads(json.dumps(py)) == j
