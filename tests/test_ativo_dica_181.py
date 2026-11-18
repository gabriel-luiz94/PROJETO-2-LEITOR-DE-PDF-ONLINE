"""TASK-181 — seletor de parâmetros: "(" já fecha o ")", o valor entra PELO NOME, ↑↓ trocam o parâmetro e ←→ o valor, Esc desfaz, parênteses vazios somem (static/ativo_dica.js, parte pura em Node)."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
pytestmark = pytest.mark.skipif(not NODE, reason="node indisponível")
MODULO = os.path.join(config.STATIC_DIR, "ativo_dica.js")
BASE = [{"ativo": "PR=PR(cabo=A35P,cabo2=S04,tipo=RL)", "projeto": ""}, {"ativo": "PR(cabo=A35P)", "projeto": "", "componente": "a", "codigo": "1"}, {"ativo": "PR(cabo=S04)", "projeto": "", "componente": "b", "codigo": "2"},
        {"ativo": "PR(cabo2=S04)", "projeto": "", "componente": "c", "codigo": "3"}, {"ativo": "PR(cabo2=A35P)", "projeto": "", "componente": "d", "codigo": "4"},
        {"ativo": "CE_4=CE_4(poste=12600D,cabo=A35P,cabo2=A35P)", "projeto": ""}, {"ativo": "CE_4(cabo=A35P)", "projeto": "", "componente": "e", "codigo": "5"},
        {"ativo": "CE_4(cabo=S04)", "projeto": "", "componente": "f", "codigo": "6"}, {"ativo": "CE_4(cabo2=S04)", "projeto": "", "componente": "g", "codigo": "7"},
        {"ativo": "OB=OB(obra.regional=X,cabo=A35P)", "projeto": ""}, {"ativo": "OB(cabo=A35P)", "projeto": "", "componente": "h", "codigo": "8"}]


def js(corpo, entrada=None):
    codigo = (f"const D=require({json.dumps(MODULO)});const e=JSON.parse(require('fs').readFileSync(0,'utf8')||'null');const base=e&&e.base;"
              f"const ctx=D.construirContexto(base||[],'');const d=n=>D.dica(n,ctx);console.log(JSON.stringify({corpo}));")
    r = subprocess.run([NODE, "-e", codigo], input=json.dumps({"base": BASE, **(entrada or {})}), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def test_abrir_parentese_ja_fecha_e_nao_duplica():
    assert js("D.fecharParenteseRecemAberto('RL, PR(, ATRINT', 7)") == {"texto": "RL, PR(), ATRINT", "cursor": 7}
    assert js("D.fecharParenteseRecemAberto('RL, PR(), ATRINT', 7)") is None                  # o ")" já está logo depois
    assert js("D.fecharParenteseRecemAberto('PR', 2)") is None                                 # não acabou de abrir


def test_valor_entra_pelo_nome_em_qualquer_ordem_e_na_linha_com_virgulas():
    """RL, PR, ATRINT: a vírgula de FORA nunca é separador de parâmetro (bug da TASK-181)."""
    r = js("(()=>{const a=D.aplicarValor('RL, PR(), ATRINT',7,d('PR'),'cabo2','S04');const b=D.aplicarValor(a.texto,a.cursor,d('PR'),'cabo','A35P');return [a,b];})()")
    assert r[0] == {"texto": "RL, PR(cabo2=S04), ATRINT", "cursor": 16, "fecha": 16}
    assert r[1]["texto"] == "RL, PR(cabo=A35P,cabo2=S04), ATRINT"                              # na ordem da declaração, mesmo gravado ao contrário
    assert js("D.aplicarValor('RL, PR(cabo2=S04), ATRINT',10,d('PR'),'cabo2','A35P')")["texto"] == "RL, PR(cabo2=A35P), ATRINT"        # gravar de novo troca só o valor


def test_traco_tira_o_valor_e_parenteses_vazios_somem():
    assert js("D.aplicarValor('RL, PR(cabo2=S04), ATRINT',10,d('PR'),'cabo2','')")["texto"] == "RL, PR(), ATRINT"
    assert js("D.limparParentesesVazios('RL, PR(), ATRINT')") == "RL, PR, ATRINT"
    assert js("D.limparParentesesVazios('N1( ), (x), PR(cabo=A)')") == "N1, (x), PR(cabo=A)"           # só vazio e colado a um nome


def test_posicional_vira_nome_ao_alterar_um_valor():
    r = js("D.aplicarValor('CE_4(12600D,A35P,A35P)',8,d('CE_4'),'cabo2','S04')")
    assert r["texto"] == "CE_4(poste=12600D,cabo=A35P,cabo2=S04)"
    assert js("D.valoresEscritos('CE_4(12600D,A35P,A35P)',8,d('CE_4'))") == {"POSTE": "12600D", "CABO": "A35P", "CABO2": "A35P"}      # posição → nome como resolver_pedido
    assert js("D.valoresEscritos('CE_4(cabo2=S04,12600D)',8,d('CE_4'))") == {"CABO2": "S04", "POSTE": "12600D"}                         # o posicional pula o que já veio por nome


def test_parametro_sob_o_cursor_e_obra_nao_alcancavel():
    assert js("[D.parametroSobOCursor('CE_4(12600D,A35P,',17,d('CE_4')), D.parametroSobOCursor('CE_4(12600D,A35P,A35P)',10,d('CE_4')), D.parametroSobOCursor('CE_4(cabo2=',11,d('CE_4'))]") == ["CABO2", "POSTE", "CABO2"]
    assert js("D.parametrosEditaveis(d('OB')).map(p=>p.nome)") == ["CABO"]                      # OBRA.* não é alcançado pelas setas
    assert js("D.aplicarValor('OB()',3,d('OB'),'cabo','A35P')")["texto"] == "OB(cabo=A35P)"


def test_grupo_do_cursor():
    assert js("D.grupoDoCursor('RL, PR(cabo=A35P), ATRINT',10)") == {"abre": 6, "fecha": 16, "nome": "PR"}
    assert js("D.grupoDoCursor('RL, PR(cabo=A35P), ATRINT',20)") is None                        # cursor fora dos parênteses
    assert js("D.grupoDoCursor('PR(cabo=A',8)") == {"abre": 2, "fecha": -1, "nome": "PR"}        # ainda aberto
    assert js("D.grupoDoCursor('(x)',2)") is None                                               # sem nome colado antes


def test_valores_do_parametro_sem_repeticao():
    v = js("d('PR').params[0].valores")
    assert len(v) == len(set(v))


def test_ligacao_na_tela():
    fonte = open(MODULO, encoding="utf-8").read()
    for t in ("fecharParenteseRecemAberto", "aplicarValor", "avancar(", "dp-item", "limparParentesesVazios", "cfg.origens"):
        assert t in fonte, t
    resumo = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    assert "origens: (input, ativo)" in resumo and "AtivoStatus.rotuloOrigem" in resumo[resumo.index("origens: (input, ativo)"):][:900]
