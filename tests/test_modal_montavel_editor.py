"""TASK-077 etapa 2 — partes puras do editor visual (static/modal_montavel_editor.js) e blocos novos aceitos pelo validador."""
import json
import os
import subprocess

import pytest

import config
from routers import modais_montaveis as mm
from tests import oraculo_tela as o

EDITOR = os.path.join(config.STATIC_DIR, "modal_montavel_editor.js")
pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado")


def js(expr, **entrada):
    codigo = f"const E=require({json.dumps(EDITOR)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({expr}));"
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def test_mover_e_limites():
    assert js("E.mover([1,2,3], 0, 1)") == [2, 1, 3]
    assert js("E.mover([1,2,3], 2, -1)") == [1, 3, 2]
    assert js("E.mover([1,2,3], 0, -1)") == [1, 2, 3] and js("E.mover([1,2,3], 2, 1)") == [1, 2, 3]


def test_id_unico_normaliza_e_evita_repetir():
    assert js("E.idUnico('Títulos de Contrato', [])") == "titulos_de_contrato"
    assert js("E.idUnico('grade', ['grade','grade_2'])") == "grade_3"
    assert js("E.idUnico('', [])") == "item"


def test_tsv_ida_e_volta_completa_colunas():
    assert js("E.tsvParaLinhas('A\\tB\\n\\nC', 3)") == [["A", "B", ""], ["C", "", ""]]
    assert js("E.linhasParaTsv([['A','B'],['C',null]])") == "A\tB\nC\t"
    assert js("E.tsvParaLinhas('A\\tB\\tC\\tD', 2)") == [["A", "B"]]


def test_cada_bloco_novo_passa_no_validador():
    """Monta um modal só com blocoNovo (na ordem grade → interpretar → destino → ação → campo → planilha → texto → saída) e valida."""
    modal = js("""(()=>{const m={id:'t',nome:'T',tipo:'montagem',blocos:[]};
        ['grade','interpretar','destino','acao_adicionar','campo','tabela_interna','texto','saida_texto'].forEach(t=>m.blocos.push(E.blocoNovo(t,m)));return m;})()""")
    assert mm.validar({"versao": 1, "modais": [modal]}) == []
    assert [b["tipo"] for b in modal["blocos"]][:2] == ["grade", "interpretar"]
    assert modal["blocos"][1]["fonte"] == modal["blocos"][0]["id"]       # interpretar já aponta para a grade existente
    assert modal["blocos"][3]["fonte"] == modal["blocos"][1]["id"]       # ação aponta para o interpretar


def test_blocos_novos_tem_ids_distintos():
    ids = js("""(()=>{const m={blocos:[]};for(let i=0;i<3;i++)m.blocos.push(E.blocoNovo('grade',m));return m.blocos.map(b=>b.id);})()""")
    assert len(set(ids)) == 3
