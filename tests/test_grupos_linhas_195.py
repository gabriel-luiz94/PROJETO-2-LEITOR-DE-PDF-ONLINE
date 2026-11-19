"""TASK-195 — grupos de linhas («@P1») nas tabelas Outros e Cabos: lógica pura (static/grupos_linhas.js) + ligação com a tela."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
pytestmark = pytest.mark.skipif(not NODE, reason="node indisponível")
S = config.STATIC_DIR


def _node(corpo):
    js = "const G = require(%s);\nconst L = (...a) => a.map(x => typeof x === 'string' ? {ativo: x} : x);\nconst nomes = l => l.map(r => (r.grupoDecl ? 'D:' : r.grupoDe ? 'M:' : '') + r.ativo);\n%s" % (json.dumps(os.path.join(S, "grupos_linhas.js")), corpo)
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def test_declaracao_arroba_e_sincronizacao():
    r = _node("""
const a = {ativo: '@P1 Rua'}, b = {ativo: 'X'}, c = {ativo: '#nota'}, d = {ativo: '@'}, e = {ativo: '  @Q'};
const s = [G.sincronizar(a, ''), G.sincronizar(b, ''), G.sincronizar(c, ''), G.sincronizar(d, ''), G.sincronizar(e, '')];
const gid = a.gid; a.ativo = 'P1'; const sai = G.sincronizar(a, '@P1 Rua');
console.log(JSON.stringify([s.map(x => x && x.tipo), !!e.grupoDecl, !!a.grupoDecl, !!a.congelada, sai && sai.tipo, sai && sai.gid === gid, G.nomeDe('@P1 Rua X'), G.rotuloDe({ativo: '@P1 Rua X'})]));""")
    assert r == [["virou", None, None, None, "virou"], True, False, False, "deixou", True, "P1", "P1 Rua X"]


def test_bloqueada_nao_vira_grupo():
    assert _node("const a = {ativo: '@P1', bloqueada: true}; console.log(JSON.stringify([G.sincronizar(a, ''), !!a.grupoDecl]));") == [None, False]


def test_enter_membro_no_fim_e_shift_enter_solta_abaixo_do_grupo():
    r = _node("""
const h = {ativo: '@P1', grupoDecl: true, congelada: true, gid: 'g1'}, m1 = {ativo: 'A', grupoDe: 'g1'}, m2 = {ativo: 'B', grupoDe: 'g1'}, z = {ativo: 'Z'};
const base = [h, m1, m2, z];
const e1 = G.novaLinhaNoGrupo(base, 0, {ativo: ''}), e2 = G.novaLinhaNoGrupo(base, 1, {ativo: ''}), s1 = G.novaLinhaSolta(base, 0, {ativo: ''}), s2 = G.novaLinhaSolta(base, 2, {ativo: ''});
console.log(JSON.stringify([nomes(e1.lista), nomes(e2.lista), nomes(s1.lista), nomes(s2.lista), G.novaLinhaNoGrupo(base, 3, {ativo: ''})]));""")
    assert r[0] == r[1] == ["D:@P1", "M:A", "M:B", "M:", "Z"]
    assert r[2] == r[3] == ["D:@P1", "M:A", "M:B", "", "Z"]
    assert r[4] is None                                  # linha comum: o Enter de sempre


def test_criar_adicionar_tirar_e_desfazer_grupo():
    r = _node("""
const l = L('A', 'B', 'C', 'D');
const c = G.criarGrupo(l, [1, 2], 'p2', t => ({ativo: t})), nc = nomes(c.lista);
const gid = c.lista[1].gid;
const soDecl = G.adicionarAoGrupo(c.lista, [1], gid);
const add = G.adicionarAoGrupo(c.lista, [c.lista.indexOf(l[3])], gid), na = nomes(add.lista);
const fora = G.tirarDoGrupo(add.lista, [add.lista.indexOf(l[1])]), nf = nomes(fora.lista);
const des = G.desfazerGrupo(fora.lista, gid), nd = nomes(des.lista);
const semNome = G.criarGrupo(L('A'), [0], '  @ ', t => ({ativo: t}));
console.log(JSON.stringify([nc, na, nf, nd, semNome, soDecl, c.n]));""")
    assert r[0] == ["A", "D:@p2", "M:B", "M:C", "D"]
    assert r[1] == ["A", "D:@p2", "M:B", "M:C", "M:D"]
    assert r[2] == ["A", "D:@p2", "M:C", "M:D", "B"]            # saiu: logo abaixo do bloco
    assert r[3] == ["A", "C", "D", "B"]
    assert r[4] is None and r[5] is None and r[6] == 2


def test_mover_linhas_e_grupos():
    r = _node("""
const mk = () => { const h = {ativo: '@P', grupoDecl: true, congelada: true, gid: 'g'}, a = {ativo: 'a', grupoDe: 'g'}, b = {ativo: 'b', grupoDe: 'g'}; return {h, a, b, x: {ativo: 'x'}, y: {ativo: 'y'}}; };
let o = mk(), f = [o.x, o.h, o.a, o.b, o.y];
const m1 = G.mover(f, 2, 1);                       // a desce dentro do grupo
o = mk(); f = [o.x, o.h, o.a, o.b, o.y]; const m2 = G.mover(f, 3, 1);          // b (último) desce: sai do grupo
o = mk(); f = [o.x, o.h, o.a, o.b, o.y]; const m3 = G.mover(f, 2, -1);         // a (primeiro) sobe: sai do grupo
o = mk(); f = [o.x, o.h, o.a, o.b, o.y]; const m4 = G.mover(f, 0, 1);          // x desce sobre a declaração: entra no grupo
o = mk(); f = [o.x, o.h, o.a, o.b, o.y]; const m5 = G.mover(f, 1, -1);         // o grupo inteiro sobe
o = mk(); f = [o.x, o.h, o.a, o.b, o.y]; const m6 = G.mover(f, 1, 1);          // o grupo inteiro desce
o = mk(); f = [o.x, o.h, o.a, o.b, o.y]; const m7 = G.mover(f, 0, -1), m8 = G.mover(f, 4, 1);
o = mk(); o.h.grupoRecolhido = true; f = [o.x, o.h, o.a, o.b, o.y]; const m9 = G.mover(f, 0, 1);  // grupo recolhido: x entra no fim do bloco (não no meio)
console.log(JSON.stringify([nomes(m1.lista), nomes(m2.lista), nomes(m3.lista), nomes(m4.lista), nomes(m5.lista), nomes(m6.lista), m7, m8, nomes(m9.lista)]));""")
    assert r[0] == ["x", "D:@P", "M:b", "M:a", "y"]
    assert r[1] == ["x", "D:@P", "M:a", "b", "y"]
    assert r[2] == ["x", "a", "D:@P", "M:b", "y"]
    assert r[3] == ["D:@P", "M:x", "M:a", "M:b", "y"]
    assert r[4] == ["D:@P", "M:a", "M:b", "x", "y"]
    assert r[5] == ["x", "y", "D:@P", "M:a", "M:b"]
    assert r[6] is None and r[7] is None
    assert r[8] == ["D:@P", "M:a", "M:b", "x", "y"]


def test_normalizar_mapa_e_reidentificar():
    r = _node("""
const h = {ativo: '@P', grupoDecl: true, gid: 'g'}, a = {ativo: 'a', grupoDe: 'g'}, x = {ativo: 'x'}, orf = {ativo: 'o', grupoDe: 'sumiu'};
const n = G.normalizar([a, x, h, orf]);             // membro antes da declaração e membro órfão
const mp = G.mapa([a, x, h, orf], n);
const copia = [JSON.parse(JSON.stringify(h)), JSON.parse(JSON.stringify(a))]; const gidAntes = copia[0].gid;
G.reidentificar(copia);
console.log(JSON.stringify([nomes(n), mp, copia[0].gid !== gidAntes, copia[1].grupoDe === copia[0].gid, G.contar([h, a, {ativo: ''}, {ativo: 'c', grupoDe: 'g', comentario: true}], 'g')]));""")
    assert r[0] == ["x", "D:@P", "M:a", "o"]
    assert r[1] == [2, 0, 1, 3]
    assert r[2] is True and r[3] is True and r[4] == 1


def test_sem_linha_arroba_nada_muda():
    r = _node("""
const l = L('A', 'B', '#nota'); const antes = JSON.stringify(l);
console.log(JSON.stringify([G.temGrupos(l), G.bloco(l, 1), G.novaLinhaNoGrupo(l, 1, {ativo: ''}), G.novaLinhaSolta(l, 1, {ativo: ''}), G.normalizar(l).length, JSON.stringify(l) === antes]));""")
    assert r == [False, None, None, None, 3, True]


def test_ligacao_com_a_tela_e_rota_sem_cache():
    js, html = open(os.path.join(S, "resumo.js"), encoding="utf-8").read(), open(os.path.join(S, "resumo.html"), encoding="utf-8").read()
    assert html.index("/static/grupos_linhas.js") < html.index("/static/resumo.js")
    assert "GL.sincronizar(row, anterior)" in js and "enterEmLinhaDeGrupo(type, idx, row, e.shiftKey)" in js
    assert "grupoDecl" in js[js.index("function deepClone(obj) {\n        const c"):]                      # desfazer/salvar levam os grupos
    assert "posicionarMenu(ctxMenu, x, y)" in js and "innerHeight - 140" not in js                  # menu mede e vira para cima: não corta mais
    assert 'id="ctx-grupo"' in html
    from fastapi.testclient import TestClient
    import app as appmod
    r = TestClient(appmod.app).get("/static/grupos_linhas.js")
    assert r.status_code == 200 and "no-store" in r.headers["cache-control"]


def test_congelar_e_bloquear_grupo_age_nos_membros():
    """TASK-196: marcar a linha «@» vale para o grupo todo (os membros entram nos alvos; a declaração continua sempre congelada)."""
    js = open(os.path.join(S, "resumo.js"), encoding="utf-8").read()
    assert "function expandirGrupos(tipo, idxs)" in js
    corpo = js[js.index("function alternarMarcaLinhas"):js.index("function alternarComentario")]
    assert "idxs = expandirGrupos(tipo, idxs)" in corpo and "!dados[i].grupoDecl" in corpo
    assert "'grupo' : 'linha'" in js
