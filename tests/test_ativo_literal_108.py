"""TASK-108 — nome literal entre colchetes ([S336,4], [CAZ 7,94], [A-B]): ativo simples com qualquer caractere (Python e JS em paridade)."""
import json
import os
import subprocess

import pytest

import config
from services import anotacao_outros as ao
from services import ativo_parametros as ap
from services import orcamento_calc as calc
from services.autonomo.totalizadora import montar_totalizadora
from tests import oraculo_tela as o

JS_AP = os.path.join(config.STATIC_DIR, "ativo_parametros.js")
JS_AO = os.path.join(config.STATIC_DIR, "anotacao_outros.js")


def L(ativo, comp="C", cod="k1"):
    return {"ativo": ativo, "desc_ativo": "d", "componente": comp, "projeto": "", "mdo": "MATERIAL", "codigo": cod, "desc_codigo": f"d {cod}",
            "fator_i": 1.0, "fator_r": 1.0, "filtro": "", "origem": "", "lv": "A"}


def total(base, texto, op="I"):
    r = calc.processar_calculo([], [{"ativo": texto, "operacao": op, "entidade": "0"}], "", base)
    return {x["codigo"]: x["total"] for x in r["resultado"]}, r["nao_encontrados"]


def test_celula_com_colchetes():
    a = ap.analisar_celula_base("[S336,4]")
    assert (a["base"], a["tipo"], a["erro"], a["apelidos"], a["params"]) == ("S336,4", "simples", None, [], [])
    a = ap.analisar_celula_base("[CAZ 7,94], 11794D")
    assert (a["base"], a["apelidos"], a["erro"]) == ("CAZ 7,94", ["11794D"], None)
    assert ap.nome_base("[a-b]") == "A-B" and ap.nome_base("[S336,4], X") == "S336,4" and ap.nome_base("N1(XLPE,DT)") == "N1"
    assert ap.separar_apelidos("[S336,4], 11336D") == ("[S336,4]", ["11336D"])
    for ruim in ["[]", "[ ]", "[A(B)]", "[A", "A]", "[A](x)", "[A]B", "[[A]]"]:
        assert ap.analisar_celula_base(ruim)["erro"], ruim


def test_dividir_tokens_nao_quebra_colchetes():
    assert ap.dividir_tokens("2-[CAZ 7,94] 1-N1") == ["2-[CAZ 7,94]", "1-N1"]
    assert ap.dividir_tokens("[A B] [C D]") == ["[A B]", "[C D]"]


@pytest.mark.parametrize("cel, pedidos", [
    ("[S336,4]", ["1-[S336,4]", "[S336,4]", "1-S336,4"]),                # S336,4 sem colchetes (vírgula entre dígitos) continua valendo
    ("[CAZ 7,94]", ["2-[CAZ 7,94]", "[CAZ 7,94], 3 [CAZ 7,94]"]),
    ("[A-B]", ["1-[A-B]", "[A-B]"]),
    ("[S336,A]", ["1-[S336,A]", "[S336,A]"]),
])
def test_calculo_le_o_nome_literal(cel, pedidos):
    for t in pedidos:
        tot, nao = total([L(cel)], t)
        assert nao == [] and list(tot) == ["k1"] and tot["k1"] > 0, (cel, t, nao)


def test_quantidades_e_varios_itens_na_linha():
    base = [L("[CAZ 7,94]", comp="CA", cod="k1"), L("N1", comp="CN", cod="k2")]
    assert total(base, "2-[CAZ 7,94] 3-N1")[0] == {"k1": 2.0, "k2": 3.0}
    assert total(base, "[CAZ 7,94], 3 N1")[0] == {"k1": 1.0, "k2": 3.0}
    assert total(base, "2 [CAZ 7,94], N1")[0] == {"k1": 2.0, "k2": 1.0}


def test_sem_colchetes_o_nome_com_virgula_continua_como_antes():
    tot, nao = total([L("S336,4")], "1-S336,4")                              # sem [ ]: a vírgula separa apelidos (nada mudou)
    assert tot == {} and nao == ["S336,4"]


def test_apelido_junto_do_nome_literal():
    tot, nao = total([L("[S336,4], 11336D")], "1-11336D")
    assert nao == [] and tot == {"k1": 1.0}


def test_neutro_do_cabo_casa_o_ativo_cadastrado_com_colchetes():
    r = calc.processar_calculo([{"ativo": "CAZ7,94 1 30", "operacao": "I", "entidade": "0"}], [], "", [L("[CAZ7,94]")])
    assert r["nao_encontrados"] == [] and [x["codigo"] for x in r["resultado"]] == ["k1"]


def test_validar_base_aponta_colchetes_invalidos():
    assert ap.validar_base([dict(L("[S336,4]"))])["erros"] == []
    assert any("parênteses" in e for e in ap.validar_base([L("[A(B)]")])["erros"])
    assert any("vazios" in e for e in ap.validar_base([L("[]")])["erros"])


def test_totalizadora_descricao_e_regra_com_nome_literal():
    base = [L("[CAZ7,94]", cod="k1")]
    cabos = [{"entidade": "V1", "operacao": "I", "ativo": "3xA35P(7,94) 30"}]
    regra = {"origem": "CABOS", "op_de": "", "ativo_de": "CAZ7,94", "acao": "ADICAO", "op_para": "", "ativo_para": "[S336,A]", "fator": 0.125, "arredondamento": "PARA CIMA", "val_min": "", "val_max": ""}
    r = montar_totalizadora(cabos, [{"entidade": "0", "operacao": "I", "ativo": "2-[CAZ7,94]"}], [regra], base, "")
    por = {(x["ativo"], x["origem"]): x for x in r}
    assert por[("CAZ7,94", "CABOS")]["naoEncontrado"] is False                   # descrição achada pelo nome literal
    assert por[("[CAZ7,94]", "OUTROS")]["naoEncontrado"] is False
    assert por[("[S336,A]", "CABOS")]["qtd"] == 4.0                               # regra casa o neutro; ativo_para também aceita [ ]


# ── paridade com o JS ──
pytestmark_node = pytest.mark.skipif(not o.NODE, reason="node não instalado")


def node(codigo, entrada):
    r = subprocess.run([o.NODE, "-e", f"const A=require({json.dumps(JS_AP)});global.AtivoParametros=A;const AO=require({json.dumps(JS_AO)});"
                        f"const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({codigo}));"],
                       input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


TEXTOS = ["[S336,4]", "[S336,4], 11336D", "[CAZ 7,94], X9", "[A-B]", "[]", "[ ]", "[A(B)]", "[A", "A]", "[A](x)", "[A]B", "[[A]]", "N1(XLPE,DT)", "N1=N1(cabo, poste)", "DT11/300, 11300D",
          "S336,4", "2-[CAZ 7,94] 1-N1", "[a b] [c d]", "[x,y], [z,w]"]


@pytestmark_node
def test_paridade_js_ativo_parametros():
    esperado = [{"sep": list(ap.separar_apelidos(t)), "nome": ap.nome_base(t), "cel": ap.analisar_celula_base(t), "ativo": ap.analisar_ativo(t), "tok": ap.dividir_tokens(t),
                 "lit": ap.literal_de(t), "sem": ap.sem_colchetes(t)} for t in TEXTOS]
    obtido = node("e.map(t=>({sep:A.separarApelidos(t),nome:A.nomeBase(t),cel:A.analisarCelulaBase(t),ativo:A.analisarAtivo(t),tok:A.dividirTokens(t),lit:A.literalDe(t),sem:A.semColchetes(t)}))", TEXTOS)
    assert obtido == json.loads(json.dumps(esperado))


@pytestmark_node
def test_paridade_js_interpretar_linha():
    linhas = ["2-[CAZ 7,94] 1-N1", "[S336,4]", "[S336,A], 2 N1", "1-[A-B]", "[CAZ 7,94], 3 [x y]", "N1, [a,b]", "1-N1 2-[A B]", "9400D, [S336,4]"]
    classico = lambda t: [{"ativo": x["ativo"], "qtd": x["qtd"]} for x in calc._classico_outros(t)]
    esperado = []
    for t in linhas:
        r = ao.interpretar_linha(t, "I", calc._classico_outros, None)
        esperado.append({"formato": r["formato"], "itens": [[i["ativo"], i["qtd"]] for i in r["itens"]], "erros": [e["cod"] for e in r["erros"]]})
    obtido = node("e.map(t=>{const cl=x=>{const R=[];AtivoParametros.dividirTokens(x).forEach(p=>{const m=p.match(/^([\\*\\-]?\\d+(?:[.,]\\d+)?)[Xx\\-](.+)$/);let q=1,a=p.trim();"
                  "if(m){let s=m[1],n=false;if(s[0]=='*'||s[0]=='-'){n=true;s=s.slice(1)}q=parseFloat(s.replace(',','.'));if(n)q=-q;a=m[2].trim()}R.push({ativo:a.toUpperCase(),qtd:q})});return R};"
                  "const r=AO.interpretarLinha(t,'I',cl,null);return {formato:r.formato,itens:r.itens.map(i=>[i.ativo,i.qtd]),erros:r.erros.map(x=>x.cod)}})", linhas)
    assert obtido == esperado
