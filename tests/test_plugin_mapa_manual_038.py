"""Plugin TASK-038 — CONTRATO da montagem manual do nanoCAD (modo «mapa») com a aplicação.

O exemplo é gerado pelo núcleo C# do plugin, sem CAD (plugin_nanocad/tests/Nucleo.Testes, teste Exemplo_da_montagem_manual_esta_em_dia):
cada ponto vira o título «@nome - comentário» + uma linha por campo (operação F, I, *I, M, *M, R, *R; F = fictícia e congelada), e cada cabo
uma linha por campo com a metragem e a geometria. O vínculo é feito depois, no mapa da aplicação.
"""
import json
import shutil
import subprocess
from pathlib import Path

import pytest

import config
from services import anotacao_outros as ao
from services import obras_arquivo as oa
from services import vinculacao_regras as vr

RAIZ = Path(__file__).resolve().parent.parent
EXEMPLO = RAIZ / "plugin_nanocad" / "exemplos" / "Montagem_Manual_Mapa.obra.json"
NODE = shutil.which("node")


@pytest.fixture
def envelope():
    return json.loads(EXEMPLO.read_text(encoding="utf-8"))


def test_importacao_aceita_sem_perder_campos(envelope):
    pronta = oa.validar_importacao(envelope, "081")
    for chave in ("cabos", "outros"):
        assert pronta["dados"][chave] == envelope["dados"][chave], f"{chave}: a importação descartou campos do plugin"


def test_lote_de_anotacoes_passa_na_conferencia_do_receber_anotacoes(envelope):
    """Mesmas checagens de /api/plugin/receber-anotacoes: nada descartado e nenhum ativo vazio."""
    dados = oa.normalizar_dados(envelope["dados"])
    for tabela in ("cabos", "outros"):
        assert len(dados[tabela]) == len(envelope["dados"][tabela])
        assert all(linha.get("ativo", "").strip() for linha in dados[tabela])


def test_pontos_viram_grupos_com_titulo_e_linhas_f_congeladas(envelope):
    outros = envelope["dados"]["outros"]
    decl = [r for r in outros if r.get("grupoDecl")]
    assert [r["ativo"] for r in decl] == ["@P1 - Rua das Flores, 120", "@P2", "@PX - poste existente"]
    gids = {r["gid"] for r in decl}
    membros = [r for r in outros if not r.get("grupoDecl")]
    assert all(r["grupoDe"] in gids for r in membros)
    fict = [r["ativo"] for r in membros if vr.eh_ficticia(r)]
    assert fict == ["TEXTO QUALQUER 123", "12600D"]                                   # fora das regras / marcado F na janela
    assert all(r["congelada"] and r["_congeladaPorF"] for r in membros if vr.eh_ficticia(r))
    p2 = next(r for r in membros if r["grupoDe"] == "plug-P2")
    assert p2["operacao"] == "R" and "_posteAprox" in p2 and "_poste" not in p2       # sem símbolo: posição aproximada
    # o poste F do ponto (12600D) preenche o parâmetro POSTE da N3 do mesmo ponto
    ctx = {"assinaturas": {"N3": {"nomes": ["POSTE"]}}, "poste_do_ponto": "12600D"}
    assert ao.interpretar_linha("N3", "I", None, ctx)["itens"][0]["ativo"] == "N3(POSTE=12600D)"


def test_cabo_com_metragem_e_geometria(envelope):
    (cabo,) = envelope["dados"]["cabos"]
    assert cabo["ativo"] == "3xA35P(7,94)-13,8kV 40 m" and cabo["operacao"] == "I" and cabo["_geoTipo"] == "MT"
    assert cabo["_geo"] == [[1000, 2000], [1038.4, 2000]]
    assert "vinculoEstruturas" not in cabo                                             # o vínculo é feito no mapa da aplicação


def test_secao_cad_guarda_a_montagem_para_continuar(envelope):
    cad = envelope["cad"]
    assert cad["semRegras"] is False
    assert [p["numero"] for p in cad["pontosMapa"]] == ["P1", "P2", "P3"]
    assert cad["cabosMapa"][0]["metros"] == 40 and cad["cabosMapa"][0]["medido"] == 38.4
    assert cad["cabosMapa"][0]["campos"][0]["metros"] == 40                             # 0.16.4: a metragem da própria linha do cabo


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_grupos_do_plugin_sao_lidos_pela_tela(envelope):
    js = """
const M = require('./static/grupos_linhas.js');
const lista = JSON.parse(process.argv[1]);
const p = M.pontoDaLinha(lista, 1), x = M.pontoDaLinha(lista, 7);
const copia = JSON.parse(JSON.stringify(lista)); M.reidentificar(copia);
console.log(JSON.stringify({p: [p.nome, p.comentario, p.formatoAntigo, p.membros], x: [x.nome, x.comentario],
  novos: copia.filter(r => r.grupoDecl).every(r => !String(r.gid).startsWith('plug-')),
  ligados: copia.filter(r => !r.grupoDecl).every(r => copia.some(d => d.grupoDecl && d.gid === r.grupoDe))}));"""
    r = subprocess.run([NODE, "-e", js, json.dumps(envelope["dados"]["outros"])], capture_output=True, text=True, timeout=60, cwd=config.BASE_DIR)
    assert r.returncode == 0, r.stderr
    out = json.loads(r.stdout)
    assert out["p"] == ["P1", "Rua das Flores, 120", False, [1, 2]]
    assert out["x"] == ["PX", "poste existente"]
    assert out["novos"] and out["ligados"]                                             # a aplicação dá ids novos aos grupos ao receber
