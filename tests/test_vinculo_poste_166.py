"""TASK-166 — vinculação POR POSTE: cada estrutura do poste liga nos cabos que CHEGAM nele (geometria `_geo` × `_poste`). Python = JS."""
import copy
import json
import os
import shutil
import subprocess

import pytest

import config
from services import vinculacao_regras as vr

NODE = shutil.which("node")
CFG = {
    "versao": 2,
    "cabos": [
        {"padrao": "^MT\\d+$", "grupo": "MT", "valor_cabo": "P", "operacoes": ["I", "M"], "pontas": {"I": 2, "M": 1}, "estruturas": "^(B1MT|N3)$"},
        {"padrao": "^BT\\d+$", "grupo": "BT", "valor_cabo": "C", "operacoes": ["I"], "estruturas": "^B2$"},
    ],
    "estruturas": [
        {"padrao": "^B1MT$", "qtd_cabos": 2, "lados": "IGUAIS", "cabos": "^MT\\d+$", "parametro": ""},
        {"padrao": "^N3$", "qtd_cabos": 1, "lados": "LIVRE", "amarracao": True, "cabos": "^MT\\d+$", "parametro": ""},
        {"padrao": "^B2$", "qtd_cabos": 2, "lados": "IGUAIS", "cabos": "^BT\\d+$", "parametro": ""},
    ],
    "ordem_maior": ["C", "P", "A"],
}


def primeiro(txt):
    return (txt or "").replace(",", " ").split()[0].upper() if (txt or "").strip() else ""


def codigos(ativo):
    return [x.strip().upper() for x in (ativo or "").split(",")][1:]          # 1º = poste


def cod_est(ativo, row=None):
    c = codigos(ativo)
    return c[0] if c else primeiro(ativo)


def poste(x, ativo, y=0):
    return {"entidade": "POSTE", "operacao": "I", "ativo": ativo, "_x": x, "_y": y + 5, "_poste": [x, y]}


def vao(a, b, ativo="MT1 40", op="I", y=0, **k):
    d = {"entidade": "CABO", "operacao": op, "ativo": ativo, "_x": (a + b) / 2, "_y": y + 5, "_geo": [[a, y], [b, y]]}
    d.update(k)
    return d


def roda(cabos, outros, **k):
    c = copy.deepcopy(cabos)
    r = vr.vincular_por_poste(c, copy.deepcopy(outros), CFG, primeiro, cod_est, codigos, aplicar=True, **k)
    return r, c


def test_b1mt_liga_nos_dois_cabos_iguais_e_cada_cabo_ganha_a_ponta_do_poste():
    outros = [poste(0, "9100D, N3"), poste(100, "9100D, B1MT"), poste(200, "9100D, N3")]
    cabos = [vao(0, 100), vao(100, 200)]
    r, c = roda(cabos, outros)
    assert c[0]["vinculoEstruturas"] == [0, 1] and c[1]["vinculoEstruturas"] == [1, 2]
    assert c[0]["vinculoOrigem"] == "poste" and r["resumo"]["completos"] == 2 and r["resumo"]["estruturas_incompletas"] == 0
    a = vr.analisar(c, outros, CFG, primeiro, cod_est, codigos)
    assert a["estruturas"][1]["estado"] == "gr" and all(x["estado"] == "gr" for x in a["cabos"].values())


def test_b1mt_com_so_um_cabo_fica_incompleta_e_avisa():
    outros = [poste(0, "9100D, N3"), poste(100, "9100D, B1MT")]
    r, c = roda([vao(0, 100)], outros)
    assert c[0]["vinculoEstruturas"] == [0, 1]          # o cabo liga nas duas pontas; a B1MT recebe 1 de 2
    assert r["resumo"]["estruturas_incompletas"] == 1
    assert any(a["tipo"] == "incompleta" and "B1MT" in a["mensagem"] and "precisa de 2" in a["mensagem"] for a in r["achados"])


def test_cabos_diferentes_nao_se_juntam_na_b1mt_iguais():
    outros = [poste(100, "9100D, B1MT")]
    cabos = [vao(0, 100, "MT1 40"), vao(100, 200, "MT2 40")]
    r, c = roda(cabos, outros)
    assert sum(len(x["vinculoEstruturas"]) for x in c) == 1          # só um dos dois (grupos iguais de tamanho 1)
    assert r["resumo"]["estruturas_incompletas"] == 1


def test_tres_cabos_iguais_escolhe_os_dois_primeiros_e_avisa():
    outros = [poste(100, "9100D, B1MT")]
    cabos = [vao(0, 100, y=0), vao(100, 200), vao(100, 150)]
    cabos[0]["_geo"] = [[0, 0], [100, 0]]
    cabos[1]["_geo"] = [[100, 0], [200, 0]]
    cabos[2]["_geo"] = [[100, 0], [150, 0]]
    r, c = roda(cabos, outros)
    assert [x["vinculoEstruturas"] for x in c] == [[0], [0], []]
    assert any(a["tipo"] == "escolha" for a in r["achados"]) and any(a["tipo"] == "sobra" for a in r["achados"])


def test_vinculo_manual_e_mantido_e_ocupa_a_estrutura():
    outros = [poste(100, "9100D, B1MT")]
    cabos = [vao(0, 100, vinculoEstruturas=[0], vinculoOrigem="manual"), vao(100, 200), vao(100, 150)]
    r, c = roda(cabos, outros)
    assert c[0]["vinculoEstruturas"] == [0] and c[0]["vinculoOrigem"] == "manual"
    assert c[1]["vinculoEstruturas"] == [0] and c[2]["vinculoEstruturas"] == [] and r["resumo"]["manuais_mantidos"] == 1


def test_vinculo_automatico_antigo_e_refeito():
    outros = [poste(0, "9100D, N3"), poste(100, "9100D, B1MT")]
    cabos = [vao(0, 100, vinculoEstruturas=[1, 1], vinculoOrigem="script")]
    _, c = roda(cabos, outros)
    assert c[0]["vinculoEstruturas"] == [0, 1] and c[0]["vinculoOrigem"] == "poste"


def test_cabo_m_so_usa_uma_ponta():
    outros = [poste(0, "9100D, N3"), poste(100, "9100D, N3")]
    _, c = roda([vao(0, 100, op="M")], outros)
    assert c[0]["vinculoEstruturas"] == [0]


def test_ponta_solta_sem_geometria_e_retirada():
    outros = [poste(0, "9100D, N3")]
    cabos = [vao(0, 500), {"entidade": "CABO", "operacao": "I", "ativo": "MT1 40", "vinculoEstruturas": [0]}, vao(0, 10, op="R")]
    r, c = roda(cabos, outros)
    tipos = [a["tipo"] for a in r["achados"]]
    assert "ponta_solta" in tipos and "sem_geometria" in tipos
    assert c[1]["vinculoEstruturas"] == [0] and "vinculoEstruturas" not in c[2]          # sem _geo e R (fora do vínculo): intocados


def test_prioridade_pela_ordem_da_tabela_de_estruturas():
    cfg = copy.deepcopy(CFG)                                   # B1MT (regra 0) e N3 (regra 1) disputam o mesmo cabo: vale a ordem da tabela
    outros = [poste(100, "9100D, N3, B1MT")]
    cabos = [vao(0, 100), vao(100, 200)]
    c = copy.deepcopy(cabos)
    vr.vincular_por_poste(c, outros, cfg, primeiro, cod_est, codigos, aplicar=True)
    assert [x["vinculoEstruturas"] for x in c] == [[0.01], [0.01]]            # B1MT (2ª da linha, 1ª da tabela) ficou com os dois


def test_raio_configuravel():
    outros = [poste(100, "9100D, N3")]
    r, _ = roda([vao(0, 103)], outros, raio=1)
    assert r["resumo"]["pontas_soltas"] == 2
    r2, c2 = roda([vao(0, 103)], outros, raio=5)
    assert c2[0]["vinculoEstruturas"] == [0]


# ── paridade Python = JS ──
_JS = r"""
const v = require(%(v)s); const e = JSON.parse(require('fs').readFileSync(0, 'utf8'));
const primeiro = t => (String(t || '').trim() ? String(t).replace(/,/g, ' ').trim().split(/\s+/)[0].toUpperCase() : '');
const codigos = a => String(a || '').split(',').map(x => x.trim().toUpperCase()).slice(1);
const codEst = (a) => { const c = codigos(a); return c.length ? c[0] : primeiro(a); };
console.log(JSON.stringify(e.casos.map(c => { const cabos = JSON.parse(JSON.stringify(c.cabos)); const r = v.vincularPorPoste(cabos, c.outros, e.cfg, primeiro, codEst, codigos, c.raio, true); return { r, cabos }; })));
"""


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_com_o_js():
    casos = [
        {"outros": [poste(0, "9100D, N3"), poste(100, "9100D, B1MT"), poste(200, "9100D, N3")], "cabos": [vao(0, 100), vao(100, 200)]},
        {"outros": [poste(100, "9100D, B1MT")], "cabos": [vao(0, 100), vao(100, 200), vao(100, 150)]},
        {"outros": [poste(100, "9100D, B1MT")], "cabos": [vao(0, 100, "MT1 40"), vao(100, 200, "MT2 40")]},
        {"outros": [poste(100, "9100D, B1MT")], "cabos": [vao(0, 100, vinculoEstruturas=[0], vinculoOrigem="manual"), vao(100, 200), vao(100, 150)]},
        {"outros": [poste(0, "9100D, N3"), poste(100, "9100D, N3")], "cabos": [vao(0, 100, op="M")]},
        {"outros": [poste(100, "9100D, N3, B1MT")], "cabos": [vao(0, 100), vao(100, 200)]},
        {"outros": [poste(100, "9100D, N3")], "cabos": [vao(0, 103), {"entidade": "CABO", "operacao": "I", "ativo": "MT1 40"}], "raio": 1},
        # Empate entre células diferentes: a primeira estrutura da tabela ainda tem prioridade.
        {"outros": [poste(2, "9100D, N3"), poste(-2, "9100D, N3"), poste(95, "9100D, N3")], "cabos": [vao(0, 100)]},
        # Há um poste na célula vizinha, um candidato mais distante e uma ponta fora do raio.
        {"outros": [poste(-13, "9100D, N3"), poste(-8, "9100D, N3"), poste(5, "9100D, N3")], "cabos": [vao(-10, 10.01)]},
    ]
    for c in casos:
        c.setdefault("raio", None)
    rr = subprocess.run([NODE, "-e", _JS % {"v": json.dumps(os.path.join(config.STATIC_DIR, "vinculacao_regras.js"))}], input=json.dumps({"casos": casos, "cfg": CFG}), capture_output=True, text=True, timeout=30)
    assert rr.returncode == 0, rr.stderr
    js = json.loads(rr.stdout)
    for c, got in zip(casos, js):
        cabos = copy.deepcopy(c["cabos"])
        py = vr.vincular_por_poste(cabos, copy.deepcopy(c["outros"]), CFG, primeiro, cod_est, codigos, raio=c["raio"] or vr.RAIO_POSTE_PADRAO, aplicar=True)
        assert got["r"] == py, c
        assert got["cabos"] == cabos, c
