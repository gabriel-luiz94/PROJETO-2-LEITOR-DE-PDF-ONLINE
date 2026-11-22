"""Linhas congeladas (fora de Totalizadora, orçamento, validação e algoritmos; o cabo congelado ainda é fonte de um vínculo manual) e bloqueadas (sem edição); ativo fora da base."""
import copy
import json
import os
import shutil
import subprocess

import pytest

import config
from services import obras_arquivo
from services import vinculacao_regras as vr
from services.autonomo import totalizadora as tt
from services.autonomo.vinculacao import itens_vinculo_para_totalizadora
from tests.test_equipamento_poste_169 import CFG, _JS, cabo, cod_est, codigos, poste, primeiro

NODE = shutil.which("node")


def _est(ativo, **k):
    return dict({"entidade": "ESTRUTURA", "operacao": "I", "ativo": ativo}, **k)


def test_obra_guarda_as_marcas_e_arquivo_antigo_abre_igual():
    dados = {"cabos": [{"entidade": "CABO", "operacao": "I", "ativo": "A35P ABC 10 m", "congelada": True}],
             "outros": [{"entidade": "0", "operacao": "I", "ativo": "N1", "bloqueada": True, "congelada": "sim"}, {"entidade": "0", "operacao": "I", "ativo": "N3"}]}
    r = obras_arquivo.normalizar_dados(dados)
    assert r["cabos"][0]["congelada"] is True and r["outros"][0]["bloqueada"] is True
    assert "congelada" not in r["outros"][0] and "congelada" not in r["outros"][1] and "bloqueada" not in r["outros"][1]           # só booleano vale
    assert obras_arquivo.normalizar_dados({"cabos": [{"ativo": "X", "operacao": "I", "entidade": "CABO"}], "outros": []})["cabos"] == [{"ativo": "X", "operacao": "I", "entidade": "CABO"}]


def _tot(cabos, outros):
    return tt.montar_totalizadora(cabos, outros, [], [], "T")


def test_totalizadora_do_servidor_nao_ve_linha_congelada():
    cabos = [{"entidade": "CABO", "operacao": "I", "ativo": "A35P ABC 10 m"}, {"entidade": "CABO", "operacao": "I", "ativo": "S04 ABC 99 m", "congelada": True}]
    outros = [{"entidade": "0", "operacao": "I", "ativo": "1-N1"}, {"entidade": "0", "operacao": "I", "ativo": "1-N3", "congelada": True}, {"entidade": "0", "operacao": "I", "ativo": "LIMITE DE SECAO 35"}]
    ativos = [l["ativo"] for l in _tot(cabos, outros)]
    assert any(a.startswith("A35P") or a == "A35P" for a in ativos) and "N1" in " ".join(ativos)
    assert not any("S04" in a or "N3" in a for a in ativos)
    # bloqueada não muda nada no cálculo
    com_bloqueio = [dict(r, bloqueada=True) for r in outros]
    assert [l["ativo"] for l in _tot(cabos, com_bloqueio)] == ativos
    # congelar = tirar a linha (mesmo resultado de a linha não existir)
    sem = [l["ativo"] for l in _tot([cabos[0]], [outros[0], outros[2]])]
    assert ativos == sem


def test_itens_do_vinculo_ignoram_cabo_ou_estrutura_congelados():
    outros = [_est("N3"), _est("N1")]
    cab = dict(cabo(0, 100, "A35P"), vinculoEstruturas=[0, 1], vinculoOrigem="manual")
    normal = itens_vinculo_para_totalizadora([cab], outros)
    assert len(normal) == 2
    assert itens_vinculo_para_totalizadora([dict(cab, congelada=True)], outros) == []
    assert [i["baseId"] for i in itens_vinculo_para_totalizadora([cab], [outros[0], dict(outros[1], congelada=True)])] == ["TOT-VINC-0-0"]


def test_cabo_congelado_com_vinculo_manual_continua_alimentando_a_estrutura():
    outros = [poste(100, "12600D, CE23")]                                                           # CE23 recebe `cabo, cabo2` do vínculo
    manual = dict(cabo(0, 100, "A35P"), vinculoEstruturas=[0], vinculoOrigem="manual")
    base = vr.parametros_por_estrutura([copy.deepcopy(manual)], copy.deepcopy(outros), CFG, primeiro, cod_est, codigos)
    gelado = vr.parametros_por_estrutura([dict(manual, congelada=True)], copy.deepcopy(outros), CFG, primeiro, cod_est, codigos)
    assert base and gelado == base                                                                 # a estrutura segue com o parâmetro do cabo congelado
    # automático: cabo congelado (não manual) não é ligado sozinho; o normal é
    auto = cabo(0, 100, "A35P")
    n3 = [poste(100, "12600D, N3")]
    c1, c2 = [copy.deepcopy(auto)], [dict(copy.deepcopy(auto), congelada=True)]
    vr.vincular_por_poste(c1, copy.deepcopy(n3), CFG, primeiro, cod_est, codigos, aplicar=True)
    vr.vincular_por_poste(c2, copy.deepcopy(n3), CFG, primeiro, cod_est, codigos, aplicar=True)
    assert c1[0].get("vinculoEstruturas") and not c2[0].get("vinculoEstruturas")
    # estrutura congelada: fora do vínculo (o cabo manual ligado a ela não alimenta nada)
    assert vr.parametros_por_estrutura([copy.deepcopy(manual)], [dict(outros[0], congelada=True)], CFG, primeiro, cod_est, codigos) == {}


@pytest.mark.skipif(not NODE, reason="sem node")
def test_paridade_com_o_javascript_das_marcas():
    manual = dict(cabo(0, 100, "A35P"), vinculoEstruturas=[0], vinculoOrigem="manual")
    casos = [
        {"vincular": False, "outros": [poste(100, "12600D, N3")], "cabos": [dict(manual, congelada=True)]},
        {"vincular": True, "outros": [poste(100, "12600D, N3")], "cabos": [dict(cabo(0, 100, "A35P"), congelada=True)]},
        {"vincular": True, "outros": [poste(100, "12600D, N3", ), poste(300, "12700D, N3")], "cabos": [cabo(0, 100, "A35P"), dict(cabo(100, 300, "A35P"), congelada=True)]},
        {"vincular": False, "outros": [dict(poste(100, "12600D, N3"), congelada=True)], "cabos": [manual]},
        {"vincular": True, "outros": [dict(poste(100, "12600D, N3"), congelada=True), poste(300, "12700D, N3")], "cabos": [cabo(0, 100, "A35P"), cabo(100, 300, "A35P")]},
    ]
    rr = subprocess.run([NODE, "-e", _JS % {"v": json.dumps(os.path.join(config.STATIC_DIR, "vinculacao_regras.js"))}], input=json.dumps({"casos": casos, "cfg": CFG}), capture_output=True, text=True, timeout=30)
    assert rr.returncode == 0, rr.stderr
    js = json.loads(rr.stdout)
    for c, got in zip(casos, js):
        cabos, outros = copy.deepcopy(c["cabos"]), copy.deepcopy(c["outros"])
        r = vr.vincular_por_poste(cabos, outros, CFG, primeiro, cod_est, codigos, aplicar=True) if c["vincular"] else None
        a = vr.analisar(cabos, outros, CFG, primeiro, cod_est, codigos)
        assert got["r"] == r, c
        assert {float(k): v for k, v in got["a"]["estruturas"].items()} == {float(k): v for k, v in a["estruturas"].items()}, c
        assert {int(k): v for k, v in got["p"].items()} == vr.parametros_por_estrutura(cabos, outros, CFG, primeiro, cod_est, codigos), c
        assert got["ach"] == vr.achados(cabos, outros, CFG, primeiro, cod_est, codigos), c
        assert got["cabos"] == cabos, c


def _node(expr, **dados):
    js = (f"const NB=require('./static/ativo_na_base.js'),G=require('./static/grade_tabela.js');const D={json.dumps(dados)};console.log(JSON.stringify({expr}));")
    r = subprocess.run([NODE, "-e", js], cwd=config.BASE_DIR, capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


BASE = [{"ativo": "N3=N3(cabo=C), APELIDO3", "componente": "N3", "codigo": "7201", "projeto": "", "origem": ""}, {"ativo": "", "componente": "N3.C", "codigo": "7203", "projeto": "PE", "origem": "OUTROS"},
        {"ativo": "CAA2", "componente": "CAA2", "codigo": "9", "projeto": "OUTRO", "origem": "CABOS"}, {"ativo": "KIT:=N3(C) + 2*CON", "componente": "", "codigo": "", "projeto": "PE", "origem": ""}]


@pytest.mark.skipif(not NODE, reason="sem node")
def test_ativo_na_base_usa_a_cascata_projeto_e_origem():
    assert _node("[...NB.nomesDaBase(D.b,'PE','OUTROS')].sort()", b=BASE) == sorted(["N3", "APELIDO3", "N3.C", "7201", "7203", "KIT"])
    nomes = "NB.nomesDaBase(D.b,'PE','OUTROS')"
    assert _node(f"NB.foraDaBase({nomes}, ['N3(A)','APELIDO3','7201','KIT','XYZ','CAA2','n3(c)',''])", b=BASE) == ["XYZ", "CAA2"]          # CAA2 é de outro projeto/origem; sem acento nem caixa
    assert _node("[...NB.nomesDaBase(D.b,'OUTRO','CABOS')].includes('CAA2')", b=BASE) is True


@pytest.mark.skipif(not NODE, reason="sem node")
def test_linha_bloqueada_nao_recebe_colagem_nem_limpeza():
    dados = [{"entidade": "0", "operacao": "I", "ativo": "A", "bloqueada": True}, {"entidade": "0", "operacao": "I", "ativo": "B"}]
    r = _node("(()=>{const d=D.d;const res=G.aplicarColagem(d,'X\\nY',{tipo:'outros',linha:0,coluna:1});return {d:d.map(x=>x.ativo),b:res.bloqueadas||0}})()", d=dados)
    assert r == {"d": ["A", "Y"], "b": 1}
    r = _node("(()=>{const d=D.d;const alt=G.limparIntervalo(d,'outros',0,1,1,1);return {d:d.map(x=>x.ativo),alt}})()", d=[dict(dados[0]), dict(dados[1])])
    assert r == {"d": ["A", ""], "alt": [1]}


def test_ligacoes_na_tela():
    ler = lambda *p: open(os.path.join(config.STATIC_DIR, *p), encoding="utf-8").read()
    js, html = ler("resumo.js"), ler("resumo.html")
    assert "/static/ativo_na_base.js" in html and 'id="ctx-congelar"' in html and 'id="ctx-bloquear"' in html and "par-nao-base" in html and "linha-congelada" in html
    assert "item.congelada) continue" in js and "alternarMarcaLinhas" in js and "AtivoNaBase.foraDaBase" in js and "r && !r.congelada) linhas.push" in js
    assert "Linha bloqueada: desbloqueie antes de excluir." in js


def test_comentario_persiste_e_vale_como_congelada():
    dados = {"cabos": [], "outros": [{"ativo": "LIMITE DE SECAO", "operacao": "I", "entidade": "0", "comentario": True},
                                      {"ativo": "N1", "operacao": "I", "entidade": "0", "comentario": "sim"}]}
    r = obras_arquivo.normalizar_dados(dados)["outros"]
    assert r[0]["comentario"] is True and r[0]["congelada"] is True           # comentário é uma linha congelada com outra aparência
    assert "comentario" not in r[1] and "congelada" not in r[1]               # valor inválido é descartado, como os outros campos
    assert obras_arquivo.normalizar_dados({"cabos": [], "outros": [{"ativo": "X", "operacao": "I", "entidade": "0"}]})["outros"] == [{"ativo": "X", "operacao": "I", "entidade": "0"}]


def _deep_clone_efetivo():
    """O `deepClone` que vale em static/resumo.js é o ÚLTIMO declarado (o de cima, JSON, é sobrescrito): ele decide o que sobrevive a desfazer/refazer e ao salvar/abrir o projeto."""
    fonte = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    i = fonte.rindex("    function deepClone(obj) {")
    ajuda = fonte[fonte.index("    function idPontoValido("):][:fonte[fonte.index("    function idPontoValido("):].index("\n") + 1]          # deepClone usa este auxiliar (fica fora da função)
    return ajuda + fonte[i:fonte.index("\n    }\n", i) + 6]


@pytest.mark.skipif(not NODE, reason="node não instalado")
def test_marcas_da_linha_sobrevivem_ao_clone_do_historico_e_do_projeto():
    codigo = _deep_clone_efetivo() + """
const linha = { entidade: '0', operacao: 'I', ativo: 'N1', congelada: true, bloqueada: true, comentario: true };
console.log(JSON.stringify([deepClone(linha), deepClone({ entidade: '0', operacao: 'I', ativo: 'N2' })]));"""
    r = subprocess.run([NODE, "-e", codigo], capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    com, sem = json.loads(r.stdout.strip().splitlines()[-1])
    assert com["congelada"] is True and com["bloqueada"] is True and com["comentario"] is True
    assert "congelada" not in sem and "bloqueada" not in sem and "comentario" not in sem            # linha comum: nada a mais


def test_historico_e_projeto_marcam_alteracao_e_hash_vira_comentario():
    fonte = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    corpo = fonte[fonte.index("    function pushHistory() {"):fonte.index("    /** Grava o estado ao vivo")]
    assert "marcarProjetoAlterado()" in corpo                                                          # TASK-182 A: congelar/bloquear/comentar gravam no projeto
    assert "_historicoSemSujo = true" in fonte[fonte.index("function restoreObraSnapshot"):][:1800]    # abrir um projeto não o deixa alterado
    assert "sincronizarHash(" in fonte and "comecaComHash" in fonte                                    # TASK-182 B: "#" = comentário
