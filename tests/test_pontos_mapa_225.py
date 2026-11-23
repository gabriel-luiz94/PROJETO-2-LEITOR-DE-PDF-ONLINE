"""Pontos do mapa como grupos de Demais Ativos, com comentário opcional e posições estáveis."""
import json
import os
import shutil

import pytest
import subprocess

import config
from services import obras_arquivo as oa

NODE = shutil.which("node")
S = config.STATIC_DIR
pytestmark = pytest.mark.skipif(not NODE, reason="node indisponível")


def _node(modulo, corpo):
    js = "const M = require(%s);\n%s" % (json.dumps(os.path.join(S, modulo)), corpo)
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


# ── Rearranjar (lógica pura) ──
_REDE = """
const NX = 6, NY = 4, nos = [], ar = [], id = (i, j) => 'n' + i + '_' + j;
for (let i = 0; i < NX; i++) for (let j = 0; j < NY; j++) nos.push({id: id(i, j), x: i * 105 + ((i * 7 + j * 3) % 5) * 9, y: j * 105 + ((i * 5 + j * 11) % 7) * 6});
for (let i = 0; i < NX; i++) for (let j = 0; j < NY; j++) { if (i < NX - 1) ar.push([id(i, j), id(i + 1, j)]); if (j < NY - 1) ar.push([id(i, j), id(i, j + 1)]); }
nos.push({id: 'p', x: 10, y: 10}, {id: 'q', x: 11, y: 10}, {id: 'b1', x: 500, y: 500}, {id: 'b2', x: 505, y: 501}); ar.push(['b1', 'b2']);
const dist = (r, a, b) => Math.hypot(r.get(a)[0] - r.get(b)[0], r.get(a)[1] - r.get(b)[1]);
"""


def test_rearranjar_vaos_uniformes_sem_sobreposicao_e_deterministico():
    r = _node("mapa_layout.js", _REDE + """
const L = 60, r1 = M.rearranjar(nos, ar, {comprimento: L}), r2 = M.rearranjar(nos, ar, {comprimento: L});
const d = ar.map(([a, b]) => dist(r1, a, b)).sort((a, b) => a - b), pts = [...r1.values()];
let minimo = Infinity; for (let i = 0; i < pts.length; i++) for (let j = i + 1; j < pts.length; j++) minimo = Math.min(minimo, Math.hypot(pts[i][0] - pts[j][0], pts[i][1] - pts[j][1]));
console.log(JSON.stringify({n: r1.size, igual: JSON.stringify([...r1]) === JSON.stringify([...r2]), mediana: d[d.length >> 1], maximo: d[d.length - 1], minimo, comp: M.componentes(nos.map(n => n.id), ar).map(c => c.length)}));""")
    assert r["n"] == 28 and r["igual"] is True
    assert 55 <= r["mediana"] <= 70 and r["maximo"] <= 150                                               # vãos de tamanho parecido: a maioria é exatamente o pedido (60)
    assert r["minimo"] >= 59                                                                            # nenhum ponto em cima de outro (grade do mesmo passo do vão)
    assert sorted(r["comp"], reverse=True) == [24, 2, 1, 1]


def test_rearranjar_ponto_fixo_nao_sai_do_lugar_e_entradas_vazias():
    r = _node("mapa_layout.js", _REDE + """
const f = M.rearranjar(nos, ar, {comprimento: 60, fixos: ['n0_0', 'p']});
console.log(JSON.stringify({a: f.get('n0_0'), a0: [nos[0].x, nos[0].y], p: f.get('p'), vazio: [...M.rearranjar([], [], {})], um: [...M.rearranjar([{id: 'x', x: 5, y: 7}], [])], vao: M.vaoTipico(f, ar, 1) > 0, semVao: M.vaoTipico(new Map(), [], 42)}));""")
    assert r["a"] == r["a0"] and r["p"] == [10, 10]
    assert r["vazio"] == [] and len(r["um"]) == 1 and r["vao"] is True and r["semVao"] == 42


# ── pontos = grupos em Demais Ativos ──
def test_nome_automatico_nome_do_usuario_e_comentario():
    r = _node("grupos_linhas.js", """
const lista = [{ativo: '@P3 Rua', grupoDecl: true, gid: 'a'}, {ativo: '@RUA X', grupoDecl: true, gid: 'b'}, {ativo: '#nota'}];
const ab = M.abrirPonto('P7', 'R', '  # esquina ');
console.log(JSON.stringify([M.proximoNome(lista), M.proximoNome(lista, 'P', 12), M.nomeLimpo('  @#rua das flores   1 '), M.nomeLimpo('  '), M.nomeLimpo('x'.repeat(90)).length,
  ab.titulo, ab.comentario, ab.titulo.gid === ab.gid,
  [undefined, '', '  ', ' # '].map(texto => M.abrirPonto('P8', 'I', texto).titulo.ativo)]));""")
    assert r[0] == "P4" and r[1] == "P13" and r[2] == "RUA DAS FLORES 1" and r[3] is None and r[4] == 60
    # TASK-235: o comentário vai no título («@P7 - esquina»), sem linha «#» separada
    assert r[5]["ativo"] == "@P7 - esquina" and r[5]["grupoDecl"] and r[5]["congelada"] and r[5]["operacao"] == "R"
    assert r[6] is None and r[7] is True
    assert r[8] == ["@P8", "@P8", "@P8", "@P8"]


def test_agrupar_pontos_mantem_nomes_unicos_e_so_insere_comentario_preenchido():
    r = _node("grupos_linhas.js", """
const a = {ativo: '12600D, N3', operacao: 'I'}, b = {ativo: 'N1', operacao: 'I'}, c = {ativo: '9600D', operacao: 'M'}, d = {ativo: 'N4', operacao: 'I'}, solto = {ativo: 'X'};
const res = M.agruparPontos([a, solto, b, c, d], [{linhas: [a, b], nome: 'p5'}, {linhas: [c], nome: 'P5', comentario: '  esquina  '}, {linhas: [d], comentario: '  '}], 9);
const nomes = l => l.map(x => (x.grupoDecl ? 'D:' : x.grupoDe ? 'M:' : '') + x.ativo);
const p = M.pontoDaLinha(res.lista, res.lista.indexOf(b));
const comentado = M.pontoDaLinha(res.lista, res.lista.indexOf(c));
console.log(JSON.stringify([nomes(res.lista), res.n, p.nome, p.comentarioIdx, p.membros.length,
  M.pontoDaLinha(res.lista, res.lista.indexOf(solto)), comentado.comentario, comentado.formatoAntigo]));""")
    assert r[0] == ["D:@P5", "M:12600D, N3", "M:N1", "X", "D:@P10 - esquina", "M:9600D", "D:@P11", "M:N4"]       # TASK-235: comentário no título
    assert r[1:6] == [3, "P5", None, 2, None]
    assert r[6:] == ["esquina", False]


def test_grupos_do_ponto_sobrevivem_ao_arquivo_da_obra():
    dados = {"cabos": [], "outros": [
        {"ativo": "@P1", "entidade": "0", "grupoDecl": True, "gid": "gabc123", "congelada": True},
        {"ativo": "# comentário", "comentario": True, "congelada": True, "grupoDe": "gabc123"},
        {"ativo": "12600D", "entidade": "POSTE", "grupoDe": "gabc123", "_mapaPoste": [0.5, 0.25]},
        {"ativo": "N3", "entidade": "ESTRUTURA", "grupoDe": 12345, "gid": ["x"]}]}
    r = oa.validar_importacao(json.loads(json.dumps(oa.montar_exportacao({"nome": "R", "projeto": "081", "dados_json": dados}))), "081")["dados"]["outros"]
    assert r[0]["grupoDecl"] is True and r[0]["gid"] == "gabc123" and r[1]["grupoDe"] == "gabc123" and r[2]["grupoDe"] == "gabc123" and r[2]["_mapaPoste"] == [0.5, 0.25]
    assert "grupoDe" not in r[3] and "gid" not in r[3]                                                  # valor inválido é descartado


def test_ligacao_na_tela_posicoes_presas_antes_de_inserir_e_rota_sem_cache():
    ui = open(os.path.join(S, "vinculacao_ui.js"), encoding="utf-8").read()
    res = open(os.path.join(S, "resumo.js"), encoding="utf-8").read()
    html = open(os.path.join(S, "resumo.html"), encoding="utf-8").read()
    corpo = ui[ui.index("function abaMapa(corpo)"):]
    assert corpo.index("fixarPosicoes();") < corpo.index("host.inserirNoMapa(tipo")                       # os pontos existentes são presos ANTES de entrar um novo
    assert "colocarNovoPonto(resultado.indice)" in corpo and "colocarNovoPonto(indice)" in corpo        # o ponto novo cai num lugar livre
    assert "'Rearranjar'" in corpo and "host.agruparPontosNoMapa(" in corpo and "caixaDoPonto(no, status, ehFicticio)" in corpo
    assert "pontoNoMapa, definirPontoNoMapa, marcarPontoNoMapa, agruparPontosNoMapa" in res
    assert "GL.abrirPonto(" in res and "pisoNomePonto" in res
    assert html.index("/static/mapa_layout.js") < html.index("/static/vinculacao_ui.js")
    from fastapi.testclient import TestClient
    import app as appmod
    r = TestClient(appmod.app).get("/static/mapa_layout.js")
    assert r.status_code == 200 and "no-store" in r.headers["cache-control"]
