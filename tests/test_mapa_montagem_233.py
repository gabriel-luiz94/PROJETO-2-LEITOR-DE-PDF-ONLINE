"""TASK-233: montagem rápida da rede no mapa — posição aproximada, fundo do CAD, pendências, ímã, «Traçar vão», seleção e atalhos."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import obras_arquivo as oa
from services.fundo_cad import fundo_do_cad

NODE = shutil.which("node")
S = config.STATIC_DIR
com_node = pytest.mark.skipif(not NODE, reason="node indisponível")


def _node(modulo, corpo):
    js = "const M = require(%s);\n%s" % (json.dumps(os.path.join(S, modulo)), corpo)
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def _ler(nome):
    return open(os.path.join(S, nome), encoding="utf-8").read()


# ── contas puras (static/mapa_montagem.js) ──
@com_node
def test_casar_um_para_um_pela_menor_distancia():
    r = _node("mapa_montagem.js", """
const as = [{id: 'T1', p: [120, 0]}, {id: 'T2', p: [1110, 0]}, {id: 'T3', p: [1150, -50]}];
const bs = [{id: 'B1', p: [0, 0]}, {id: 'B2', p: [1000, 0]}, {id: 'B3', p: [5000, 0]}];
console.log(JSON.stringify({
  pares: M.casar(as, bs, 200).map(p => [p.a, p.b, Math.round(p.d)]),
  curto: M.casar(as, bs, 80).length,
  empate: M.casar([{id: 'a', p: [0, 0]}, {id: 'b', p: [2, 0]}], [{id: 'x', p: [1, 0]}], 5).map(p => p.a)
}));""")
    assert r["pares"] == [["T2", "B2", 110], ["T1", "B1", 120]]          # T3 perde o B2 para o T2 (mais perto) e não pega outro
    assert r["curto"] == 0
    assert r["empate"] == ["a"]                                           # empate desfaz pela ordem: determinístico


@com_node
def test_mais_perto_blocos_livres_recorte_retangulo_metros_progresso():
    r = _node("mapa_montagem.js", """
const c = [{id: 'a', p: [0, 0]}, {id: 'b', p: [10, 0]}, {id: 'c', p: [30, 0]}];
const pos = new Map([['a', [0, 0]], ['b', [10, 5]], ['c', [50, 50]]]);
const fundo = {redes: [{t: 'BT', p: [[0, 0], [10, 0]]}, {t: 'MT', p: [[9000, 9000], [9100, 9000]]}], postes: [[0, 0], [5, 5], [9000, 9000]]};
console.log(JSON.stringify({
  perto: M.maisPerto([8, 0], c, 25), longe: M.maisPerto([100, 0], c, 25), ignora: M.maisPerto([8, 0], c, 25, new Set(['b'])).id,
  livres: M.blocosLivres([[0, 0], [0.5, 0], [5, 5], [40, 40]], [[0, 0.4]], 1).map(b => b.id),
  recorte: M.recortarFundo(fundo, [0, 0, 100, 100]),
  ret: M.dentroDoRetangulo(pos, [20, 20], [-1, -1]).sort(),
  m: [M.metros([0, 40], [80, 0]), M.metros([0, 0], [0.2, 0]), M.metros(null, [1, 1])],
  prog: [M.progresso(47, 52), M.progresso(0, 0)]
}));""")
    assert r["perto"] == {"id": "b", "d": 2} and r["longe"] is None and r["ignora"] == "a"
    assert r["livres"] == ["b2", "b3"]                                    # o bloco que já é ponto (até 1 m) some
    assert r["recorte"] == {"redes": [{"t": "BT", "p": [[0, 0], [10, 0]]}], "postes": [[0, 0], [5, 5]]}   # fundo de outro lugar não aparece
    assert r["ret"] == ["a", "b"]
    assert r["m"] == [89, 1, None]
    assert r["prog"][0]["texto"] == "47/52 (90%)" and r["prog"][1]["pct"] == 100


# ── posição aproximada no grafo (static/vinculacao_regras.js) ──
@com_node
def test_ponto_sem_bloco_usa_posicao_aproximada_sem_entrar_no_encaixe():
    r = _node("vinculacao_regras.js", """
const outros = [
  {entidade: 'POSTE', operacao: 'I', ativo: '11300D, N4', _poste: [0, 0]},
  {entidade: 'POSTE', operacao: 'I', ativo: '11300D, N4', _posteAprox: [40, 0], _x: 44, _y: 3},
  {entidade: 'POSTE', operacao: 'I', ativo: '11300D, N4', _x: 80, _y: 2},
  {entidade: 'POSTE', operacao: 'I', ativo: '11300D, N4'},
  {entidade: 'POSTE', operacao: 'I', ativo: '11300D, N4', _poste: [120, 0], _x: 125, _y: 9}
];
const cabos = [{entidade: 'CABO', operacao: 'I', ativo: 'CAA2 ABC 40 m', _geo: [[0, 0], [40, 0]]}];
const g = M.grafoDaRede(cabos, outros, () => '', () => [], () => true);
const no = l => g.nosPorLinha.get(l);
console.log(JSON.stringify({
  ids: [0, 1, 2, 3, 4].map(l => no(l).id), pos: [0, 1, 2, 3, 4].map(l => no(l).pos), aprox: [0, 1, 2, 3, 4].map(l => no(l).posAprox || null),
  pontas: g.cabosDoMapa.get(0).ids, util: M.posicaoAproximada({_x: 1, _y: 2}), comPoste: M.posicaoAproximada({_poste: [1, 1], _x: 1, _y: 2})
}));""")
    assert r["ids"] == ["0,0", "linha:1", "linha:2", "linha:3", "120,0"]   # o id não muda (posições salvas e vínculos seguem valendo)
    assert r["aprox"] == [None, [40, 0], [80, 2], None, None]              # `_posteAprox` do plugin vale antes do `_x/_y` da anotação
    assert r["pontas"] == ["0,0", None]                                    # a posição aproximada não liga a ponta do cabo (só o encaixe de 5 m no bloco)
    assert r["util"] == [1, 2] and r["comPoste"] is None


# ── fundo do CAD (services/fundo_cad.py) ──
def _cad():
    return {
        "postes": [
            {"handle": "A1", "x": 0, "y": 0, "ativoLido": "9200D, N1"},
            {"handle": "A2", "x": 40, "y": 0, "ativoLido": ""},                  # bloco livre
            {"handle": "A3", "x": 0.3, "y": 0.2, "ativoLido": ""},                # junto do anotado (troca de poste): não é livre
            {"handle": "", "x": 90, "y": 5, "ativoLido": "9300F, N3"},           # poste só anotação: não é bloco
            {"handle": "A4", "x": 80, "y": 0, "ativoLido": "", "excluido": True},
            {"handle": "A5", "x": 40, "y": 0.4, "ativoLido": ""},                # mesmo lugar do A2
        ],
        "vaos": [
            {"linha": "L1", "trecho": 0, "ativoLido": "CAA2 40 m", "tipo": "BT", "pontos": [[0, 0], [40, 0]]},
            {"linha": "L2", "trecho": 0, "ativoLido": "", "tipo": "MT", "pontos": [[0, 0], [0, 40], [0, 80]]},
            {"linha": "L2", "trecho": 0, "ativoLido": "", "tipo": "MT", "pontos": [[0, 0], [0, 40]]},      # repetido
            {"linha": "L3", "trecho": 1, "ativoLido": "", "tipo": "?", "pontos": [[1, 1]]},              # um ponto só
        ],
    }


def test_fundo_do_cad_so_com_o_que_nao_virou_linha_da_obra():
    f = fundo_do_cad(_cad())
    assert f == {"redes": [{"t": "MT", "p": [[0.0, 0.0], [0.0, 40.0], [0.0, 80.0]]}], "postes": [[40.0, 0.0]]}
    assert fundo_do_cad(None) is None and fundo_do_cad({"postes": [], "vaos": []}) is None
    assert fundo_do_cad({"postes": [{"handle": "X", "x": "nan", "y": 1}], "vaos": "x"}) is None


def test_projeto_do_plugin_abre_com_o_fundo():
    from services import projeto_obra
    env = {"formato": oa.FORMATO, "versao": oa.VERSAO, "tipo": "obra", "nome": "n", "projeto": "081",
           "dados": {"cabos": [], "outros": [{"entidade": "POSTE", "operacao": "I", "ativo": "9300F, N3", "_posteAprox": [90, 5]}], "totalizadora": []},
           "cad": _cad()}
    r = projeto_obra.ler(json.dumps(env), "081")
    assert r["cad"] is True and r["fundo"]["postes"] == [[40.0, 0.0]]
    assert r["dados"]["outros"][0]["_posteAprox"] == [90, 5]                # o campo do plugin sobrevive à importação
    sem = dict(env); sem.pop("cad")
    assert projeto_obra.ler(json.dumps(sem), "081")["fundo"] is None


def test_configuracao_de_leitura_aceita_o_raio_do_bloco_livre():
    from routers.plugin_cad import validar_leitura
    assert validar_leitura({"raio_bloco_livre_alturas": 25})["raio_bloco_livre_alturas"] == 25
    cs = open(os.path.join(config.BASE_DIR, "plugin_nanocad", "src", "Nucleo", "ConfigLeitura.cs"), encoding="utf-8").read()
    assert '"raio_bloco_livre_alturas"' in cs and "RaioBlocoLivreAlturas = 25" in cs
    leitor = open(os.path.join(config.BASE_DIR, "plugin_nanocad", "src", "Nucleo", "LeitorProjeto.cs"), encoding="utf-8").read()
    assert "CasarComBlocosLivres(semBloco, ctx, por)" in leitor
    obra = open(os.path.join(config.BASE_DIR, "plugin_nanocad", "src", "Nucleo", "ObraArquivo.cs"), encoding="utf-8").read()
    assert 'linha.Set("_posteAprox"' in obra


# ── a tela (static/vinculacao_ui.js, resumo.js, projeto.js) ──
def test_mapa_tem_as_ferramentas_da_montagem_rapida():
    ui = _ler("vinculacao_ui.js")
    corpo = ui[ui.index("function abaMapa(corpo)"):]
    for trecho in ("'Traçar vão'", "function cliqueVao(", "function religarPonta(", "dataset.arrastar === 'ponta'", "function calcularPendencias(",
                   "'Vãos vinculados'", "function ligarSugestoes(", "function fixarCoordenada(", "vz-fundo-bloco", "function transformarBloco(",
                   "function excluirSelecao(", "MM.dentroDoRetangulo(", "doc.addEventListener('keydown', teclaMapa, true)",
                   "doc.removeEventListener('keydown', teclaMapa, true)", "function mostrarMaisCartoes(", "st.modeloCabo ="):
        assert trecho in corpo, trecho
    # Esc com ferramenta/seleção ativa não fecha a janela (o atalho roda na captura e para a propagação)
    tecla = corpo[corpo.index("const teclaMapa"):corpo.index("doc.addEventListener('keydown', teclaMapa, true)")]
    assert "e.stopPropagation()" in tecla and "INPUT|TEXTAREA|SELECT" in tecla
    res = _ler("resumo.js")
    assert "copiarPontoNoMapa, copiarCaboNoMapa, removerDoMapa, fundoCad," in res and "window.OrcaFundoCad" in res
    assert "'_poste', '_posteAprox'" in res and "c._posteAprox = aprox.slice()" in res
    assert "aplicarFundoCad(pr)" in _ler("projeto.js")
    html = _ler("resumo.html")
    assert html.index("/static/mapa_montagem.js") < html.index("/static/vinculacao_ui.js")
    from fastapi.testclient import TestClient
    import app as appmod
    r = TestClient(appmod.app).get("/static/mapa_montagem.js")
    assert r.status_code == 200 and "no-store" in r.headers["cache-control"]
