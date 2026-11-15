"""TASK-171 — a estrutura fictícia do cabo sugere os componentes avulsos (alças, conectores) para a tabela Outros. Python = JS."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import ativo_parametros as ap
from services import avulsos_ficticia as av
from services import obras_arquivo

NODE = shutil.which("node")


def L(ativo, comp="", cod="", proj="PERNAMBUCO"):
    return {"ativo": ativo, "desc_ativo": "d", "componente": comp, "projeto": proj, "mdo": "MATERIAL", "codigo": cod, "desc_codigo": "d", "fator_i": 1, "fator_r": 0,
            "filtro": "", "origem": "OUTROS", "lv": "A", "arred": "NORMAL", "formula_i": "", "formula_r": ""}


BASE = ap.preencher_componentes([
    L("ALCA=ALCA(cabo=S04)"), L("ALCA(cabo=%)", "ALCA", "7"), L("CON=CON(cabo=S04,cabo2=S04,mdo=M)"), L("CON(cabo=%)", "CON", "8"),
    L("N4=N4(cabo=S04,cabo2=S04)"), L("N4(cabo=%,cabo2=%)", "N4", "1"), L("N4(cabo=%,cabo2=%):=3*ALCA(cabo) + 3*ALCA(cabo2)"),
    L("N3=N3(cabo=S04)"), L("N3(cabo=%)", "N3", "2"), L("N3(cabo=%):=3*ALCA(cabo)"),
    L("CE_4=CE_4(poste=D,cabo=S04,cabo2=S04)"), L("CE_4(poste=%D,cabo=%,cabo2=%)", "CE4", "3"), L("CE_4(poste=%D,cabo=%,cabo2=%):=6*ALCA(cabo) + 6*ALCA(cabo2)"),
    L("N1=N1(cabo=S04)"), L("N1(cabo=%)", "N1", "4"),
    L("CHAVE_FACA=CHAVE_FACA(cabo=S04,cabo2=S04)"), L("CHAVE_FACA(cabo=%,cabo2=%)", "CF", "5"), L("CHAVE_FACA(cabo=%,cabo2=%):=2*CON(cabo,cabo2,mdo=N)"),
])[0]
CFG = {"versao": 2, "cabos": [{"padrao": "^(A\\d+P|S04)$", "grupo": "MT", "valor_cabo": "P", "operacoes": ["I", "M", "R"], "pontas": {"I": 2, "M": 1, "R": 2}}],
       "estruturas": [{"padrao": "^(N4|CE_4|CHAVE_FACA)$", "qtd_cabos": 2, "lados": "LIVRE", "parametros": "cabo, cabo2", "parametro": "", "cabos": ""},
                      {"padrao": "^N3$", "qtd_cabos": 1, "amarracao": True, "lados": "LIVRE", "parametros": "cabo", "parametro": "", "cabos": ""},
                      {"padrao": "^N1$", "qtd_cabos": 2, "lados": "IGUAIS", "parametros": "cabo", "parametro": "", "cabos": ""}],
       "ordem_maior": ["C", "P", "A"]}
SIG = ap.assinaturas_do_projeto(BASE, "PERNAMBUCO")
GRUPOS = av.grupos_de_chamada(BASE, "PERNAMBUCO")


def primeiro(txt):
    return (txt or "").replace(",", " ").split()[0].upper() if (txt or "").strip() else ""


def cabo(cod, op="I", **k):
    d = {"entidade": "CABO", "operacao": op, "ativo": f"{cod} 45", "textoOriginal": f"3x{cod}-13,8kV 45 m"}
    d.update(k)
    return d


def sugestao(cod, ficticia, op="I"):
    return av.avulsos_da_ficticia(cabo(cod, op), ficticia, CFG, primeiro, SIG, GRUPOS)


def test_n4_so_o_lado_do_cabo_novo():
    assert sugestao("A35P", "N4") == [{"ativo": "ALCA(cabo=3XA35P-13.8KV)", "qtd": 3.0, "operacao": "I", "completar": False}]


def test_n3_e_ce4_com_poste_padrao():
    assert sugestao("A185P", "N3") == [{"ativo": "ALCA(cabo=3XA185P-13.8KV)", "qtd": 3.0, "operacao": "I", "completar": False}]
    assert sugestao("A70P", "CE_4") == [{"ativo": "ALCA(cabo=3XA70P-13.8KV)", "qtd": 6.0, "operacao": "I", "completar": False}]


def test_n1_sem_linhas_de_chamada_nao_sugere():
    assert sugestao("A35P", "N1") == []


def test_retirada_mantem_a_operacao_do_cabo():
    assert sugestao("A35P", "N4", "R") == [{"ativo": "ALCA(cabo=3XA35P-13.8KV)", "qtd": 3.0, "operacao": "R", "completar": False}]


def test_item_com_cabo_e_o_outro_lado_entra_em_branco_e_marcado():
    assert sugestao("A35P", "CHAVE_FACA") == [{"ativo": "CON(cabo=3XA35P-13.8KV, cabo2=?, mdo=N)", "qtd": 2.0, "operacao": "I", "completar": True}]


def test_duas_ficticias_no_mesmo_cabo_uma_lista_por_ficticia():
    c = cabo("A35P", estruturasFicticias=["N4", "N3"])
    r = av.avulsos_do_cabo(c, CFG, primeiro, SIG, GRUPOS)
    assert [x["ficticia"] for x in r] == ["N4", "N3"] and all(x["itens"][0]["qtd"] == 3.0 for x in r)
    assert av.avulsos_do_cabo(cabo("A35P", estruturasFicticias=["N1"]), CFG, primeiro, SIG, GRUPOS) == []
    assert av.avulsos_do_cabo(cabo("A35P"), CFG, primeiro, SIG, GRUPOS) == []                              # sem fictícia: nada muda


def test_itens_iguais_somam():
    base = BASE + ap.preencher_componentes([L("N3(cabo=%A%):=2*ALCA(cabo)")])[0]
    g = av.grupos_de_chamada(base, "PERNAMBUCO")
    r = av.avulsos_da_ficticia(cabo("A35P"), "N3", CFG, primeiro, ap.assinaturas_do_projeto(base, "PERNAMBUCO"), g)
    assert r[0]["qtd"] == 5.0 and len(r) == 1


def test_origem_avulso_e_aceita_no_obra():
    chave = av.chave_origem(2, "n4")
    assert chave == "2:N4"
    dados = {"cabos": [], "outros": [{"entidade": "OUTROS", "operacao": "I", "ativo": "ALCA(cabo=3XA35P-13.8KV)", "qtdAtivos": 3, "texto": "avulso da N4 fictícia — cabo #3", "origemAvulso": chave}], "totalizadora": []}
    assert "origemAvulso" in obras_arquivo.CAMPOS_VINCULO
    assert obras_arquivo.normalizar_dados(dados)["outros"][0]["origemAvulso"] == "2:N4"


_JS = r"""
const A = require(%(a)s), AP = require(%(ap)s); const e = JSON.parse(require('fs').readFileSync(0, 'utf8'));
const primeiro = t => (String(t || '').trim() ? String(t).replace(/,/g, ' ').trim().split(/\s+/)[0].toUpperCase() : '');
const sig = AP.assinaturasDoProjeto(e.base, 'PERNAMBUCO'), grupos = A.gruposDeChamada(e.base, 'PERNAMBUCO');
console.log(JSON.stringify({ g: Object.keys(grupos).sort().map(k => [k, grupos[k].length]), casos: e.casos.map(([c, f]) => A.avulsosDaFicticia(c, f, e.cfg, primeiro, sig, grupos)),
  doCabo: e.duplas.map(c => A.avulsosDoCabo(c, e.cfg, primeiro, sig, grupos)), est: A.estadoDosAvulsos(2, 'N4', [{ ativo: 'ALCA(cabo=X)', qtd: 3 }], [{}, {}, { origemAvulso: '2:N4', ativo: '3-ALCA(cabo=X)' }, { origemAvulso: '2:N4', ativo: '3-ALCA(cabo=Y)' }]), ok: A.estadoDosAvulsos(2, 'N4', [{ ativo: 'ALCA(cabo=X)', qtd: 3 }], [{}, {}, { origemAvulso: '2:N4', ativo: '3-ALCA(cabo=X)' }]).estado }));
"""


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_com_o_js():
    casos = [[cabo("A35P"), "N4"], [cabo("A185P"), "N3"], [cabo("A70P"), "CE_4"], [cabo("A35P"), "N1"], [cabo("A35P", "R"), "N4"], [cabo("A35P"), "CHAVE_FACA"], [cabo("A35P"), "INEXISTENTE"]]
    duplas = [cabo("A35P", estruturasFicticias=["N4", "N3"]), cabo("A35P")]
    rr = subprocess.run([NODE, "-e", _JS % {"a": json.dumps(os.path.join(config.STATIC_DIR, "avulsos_ficticia.js")), "ap": json.dumps(os.path.join(config.STATIC_DIR, "ativo_parametros.js"))}],
                        input=json.dumps({"base": BASE, "cfg": CFG, "casos": casos, "duplas": duplas}), capture_output=True, text=True, timeout=30)
    assert rr.returncode == 0, rr.stderr
    js = json.loads(rr.stdout)
    assert js["g"] == sorted([[k, len(v)] for k, v in GRUPOS.items()])
    for (c, f), got in zip(casos, js["casos"]):
        assert got == av.avulsos_da_ficticia(c, f, CFG, primeiro, SIG, GRUPOS), (c["ativo"], f)
    for c, got in zip(duplas, js["doCabo"]):
        assert got == av.avulsos_do_cabo(c, CFG, primeiro, SIG, GRUPOS)
    assert js["est"] == {"estado": "desatualizado", "linhas": [2, 3], "velhas": []} and js["ok"] == "adicionado"
