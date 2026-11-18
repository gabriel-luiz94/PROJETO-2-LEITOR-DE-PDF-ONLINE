"""TASK-185 — painel de vinculação rápido (lista virtualizada), aba «Cabos ⇄ Postes» e tabelas grandes do Resumo (hidratação das linhas).

Regras puras (Python = JS): o usuário diz a que poste(s) o cabo chega (`vinculoPostes`) e o algoritmo por poste escolhe a estrutura — sem geometria.
O resto é checagem estática da tela (a medição de desempenho foi feita no navegador; aqui garantimos que as técnicas continuam no código).
"""
import copy
import json
import os
import shutil
import subprocess

import pytest

import config
from services import obras_arquivo, vinculacao_regras as vr

NODE = shutil.which("node")
CFG = {
    "versao": 2,
    "cabos": [{"padrao": "^MT\\d+$", "grupo": "MT", "valor_cabo": "P", "operacoes": ["I", "M"], "pontas": {"I": 2, "M": 1}, "estruturas": "^(B1MT|N3)$"}],
    "estruturas": [
        {"padrao": "^B1MT$", "qtd_cabos": 2, "lados": "IGUAIS", "cabos": "^MT\\d+$", "parametro": ""},
        {"padrao": "^N3$", "qtd_cabos": 1, "lados": "LIVRE", "amarracao": True, "cabos": "^MT\\d+$", "parametro": ""},
    ],
    "ordem_maior": ["C", "P", "A"],
}


def primeiro(txt):
    return (txt or "").replace(",", " ").split()[0].upper() if (txt or "").strip() else ""


def codigos(ativo):
    return [x.strip().upper() for x in (ativo or "").split(",")][1:]


def cod_est(ativo, row=None):
    c = codigos(ativo)
    return c[0] if c else primeiro(ativo)


def cabo(ativo="MT1 40", **k):
    d = {"entidade": "CABO", "operacao": "I", "ativo": ativo}          # sem _geo: só vale o que o usuário atribuiu
    d.update(k)
    return d


def poste(ativo):
    return {"entidade": "POSTE", "operacao": "I", "ativo": ativo}      # sem _poste


def roda(cabos, outros):
    c = copy.deepcopy(cabos)
    r = vr.vincular_por_poste(c, copy.deepcopy(outros), CFG, primeiro, cod_est, codigos, aplicar=True)
    return r, c


OUTROS = [poste("9100D, N3"), poste("9100D, B1MT"), poste("9100D, N3")]


def test_cabo_com_postes_atribuidos_liga_na_estrutura_de_cada_poste_sem_geometria():
    r, c = roda([cabo(vinculoPostes=[0, 1]), cabo(vinculoPostes=[1, 2])], OUTROS)
    assert c[0]["vinculoEstruturas"] == [0, 1] and c[1]["vinculoEstruturas"] == [1, 2]
    assert c[0]["vinculoOrigem"] == "poste"
    assert r["resumo"]["completos"] == 2 and r["achados"] == [] or all(a["tipo"] != "ponta_solta" for a in r["achados"])


def test_cabo_sem_geometria_e_sem_postes_continua_intocado():
    r, c = roda([cabo(vinculoEstruturas=[2], vinculoOrigem="script")], OUTROS)
    assert c[0]["vinculoEstruturas"] == [2]
    assert any(a["tipo"] == "sem_geometria" for a in r["achados"])


def test_um_poste_so_o_cabo_fica_parcial_e_b1mt_avisa_incompleta():
    r, c = roda([cabo(vinculoPostes=[1])], OUTROS)
    assert c[0]["vinculoEstruturas"] == [1]
    assert r["resumo"]["parciais"] == 1 and r["resumo"]["estruturas_incompletas"] == 3          # B1MT (1 de 2) + as duas N3 sem cabo


def test_poste_inexistente_e_ignorado_e_duplicado_conta_uma_vez():
    assert vr.postes_atribuidos(cabo(vinculoPostes=[1, 1, -1, 2.5, True, "x", 3])) == [1, 3]
    r, c = roda([cabo(vinculoPostes=[99])], OUTROS)
    assert c[0].get("vinculoEstruturas") in (None, [])


def test_vinculo_manual_continua_mandando_mesmo_com_postes_atribuidos():
    r, c = roda([cabo(vinculoPostes=[0], vinculoEstruturas=[2], vinculoOrigem="manual")], OUTROS)
    assert c[0]["vinculoEstruturas"] == [2] and c[0]["vinculoOrigem"] == "manual"


def test_cabo_congelado_nao_e_tocado():
    r, c = roda([cabo(vinculoPostes=[0], congelada=True)], OUTROS)
    assert not c[0].get("vinculoEstruturas")


_JS = r"""
const v = require(%(v)s); const e = JSON.parse(require('fs').readFileSync(0, 'utf8'));
const primeiro = t => (String(t || '').trim() ? String(t).replace(/,/g, ' ').trim().split(/\s+/)[0].toUpperCase() : '');
const codigos = a => String(a || '').split(',').map(x => x.trim().toUpperCase()).slice(1);
const codEst = (a) => { const c = codigos(a); return c.length ? c[0] : primeiro(a); };
console.log(JSON.stringify(e.casos.map(c => { const cabos = JSON.parse(JSON.stringify(c.cabos)); const r = v.vincularPorPoste(cabos, c.outros, e.cfg, primeiro, codEst, codigos, undefined, true); return { r, cabos, at: c.cabos.map(x => v.postesAtribuidos(x)) }; })));
"""


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_com_o_js():
    casos = [
        {"outros": OUTROS, "cabos": [cabo(vinculoPostes=[0, 1]), cabo(vinculoPostes=[1, 2])]},
        {"outros": OUTROS, "cabos": [cabo(vinculoPostes=[1])]},
        {"outros": OUTROS, "cabos": [cabo(vinculoPostes=[0], vinculoEstruturas=[2], vinculoOrigem="manual"), cabo(vinculoPostes=[1, 2])]},
        {"outros": OUTROS, "cabos": [cabo(vinculoPostes=[1, 1, 99]), cabo(vinculoEstruturas=[2], vinculoOrigem="script")]},
        {"outros": OUTROS, "cabos": [cabo("MT1 40", operacao="M", vinculoPostes=[0, 2])]},
    ]
    rr = subprocess.run([NODE, "-e", _JS % {"v": json.dumps(os.path.join(config.STATIC_DIR, "vinculacao_regras.js"))}], input=json.dumps({"casos": casos, "cfg": CFG}), capture_output=True, text=True, timeout=30)
    assert rr.returncode == 0, rr.stderr
    js = json.loads(rr.stdout)
    for c, got in zip(casos, js):
        cabos = copy.deepcopy(c["cabos"])
        py = vr.vincular_por_poste(cabos, copy.deepcopy(c["outros"]), CFG, primeiro, cod_est, codigos, aplicar=True)
        assert got["r"] == py, c
        assert got["cabos"] == cabos, c
        assert got["at"] == [vr.postes_atribuidos(x) for x in c["cabos"]], c


def test_obra_guarda_e_valida_vinculo_postes():
    lin = {"entidade": "CABO", "operacao": "I", "ativo": "MT1 40", "vinculoPostes": [0, 3]}
    r = obras_arquivo.normalizar_dados({"cabos": [lin, dict(lin, vinculoPostes="x"), dict(lin, vinculoPostes=[0, "a"])], "outros": []})
    assert r["cabos"][0]["vinculoPostes"] == [0, 3]
    assert "vinculoPostes" not in r["cabos"][1] and "vinculoPostes" not in r["cabos"][2]


# ── a tela (estático) ──
def _ler(nome):
    with open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8") as f:
        return f.read()


def test_painel_do_modal_e_virtualizado_e_sem_medir_layout():
    ui = _ler("vinculacao_ui.js")
    for trecho in ["function criarPainel(", "function ctx()", "function assinatura()", "const metaAtivo = new Map()", "function atualizar()", "faixa(dados.esq, y0, y1)", "id: 'cabopostes'" if False else "['cabopostes', 'Cabos ⇄ Postes'", "abaCabosPostes", "VR.postesAtribuidos"]:
        assert trecho in ui, trecho
    painel = ui.split("function criarPainel(")[1].split("// ── tabelas editáveis")[0]
    assert "getBoundingClientRect" not in painel                             # posição das linhas de ligação sai por aritmética (altura fixa), não por medição
    assert "vz-card" not in ui and "vz-ebox" not in ui                    # blocos grandes antigos
    assert "background:#05080d" in ui                                     # fundo opaco: o translúcido refazia a mistura com a página a cada quadro da rolagem
    assert "const RH = 30" in ui                                          # uma linha por cabo e por poste


def test_resumo_hidrata_as_linhas_de_tabelas_grandes_aos_poucos():
    js = _ler("resumo.js")
    for trecho in ["const LIMIAR_LAZY = 60", "function hidratarLinha(", "function hidratarJanela(", "function hidratarTudo(", "function agendarHidratacao(", "tr._montar = montar", "agendarResumoRede();",
                   "if (lazy) { hidratarJanela(type, 500); agendarHidratacao(type); }", "vinculoPostes"]:
        assert trecho in js, trecho
    assert "paramsVinc" not in js                                         # cálculo que ninguém usava no desenho
    # só colar/copiar/cortar/arrastar montam tudo antes; clique, Enter, excluir etc. NÃO (uma rajada de ações não paga a montagem de todas as linhas)
    gat = js.split("function ligarGatilhosHidratacao")[1].split("let raf = 0;")[0]
    for ev in ["'paste'", "'copy'", "'cut'", "'dragstart'"]:
        assert ev in gat
    for ev in ["'mousedown'", "'click'", "'contextmenu'"]:
        assert ev not in gat
    # ícone da lixeira sem filhos no DOM
    assert "<polyline" not in js.split("btnDel = document.createElement('button')")[1].split("btnDel.addEventListener")[0]


def test_resumo_html_sem_desfoque_de_fundo_nos_paineis_das_tabelas():
    html = _ler("resumo.html")
    secao = html.split(".report-section {")[1].split("}")[0]
    assert "backdrop-filter" not in secao.replace("/* sem backdrop-filter", "")
    assert "tr.lazy-row" in html and ".btn-row-del::before" in html


def test_regras_deixam_a_inicializacao_declarar_os_caches_no_topo():
    js = _ler("resumo.js")
    topo = js.split("function renderTable(")[0]
    for nome in ["_lazy", "_lazyTm", "_hLinhaReal", "_tmResumoRede", "_medidasTabela", "_cacheCores"]:
        assert nome in topo.split("function ")[0] or nome in topo, nome   # (let/const mais abaixo dariam erro de inicialização: renderTable roda cedo)


# ── melhorias seguintes (TASK-185b) ──
@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_layer_em_aceita_curinga_no_motor_do_leitor():
    js = ("const E = require(%s); const r = JSON.parse(require('fs').readFileSync(0, 'utf8')); "
          "console.log(JSON.stringify(r.map(([l, x]) => E.layerCasa(l, x))));") % json.dumps(os.path.join(config.STATIC_DIR, "regras_leitor_engine.js"))
    casos = [[["POSTE_*_LV"], "poste_instalando_lv"], [["POSTE_*_LV"], "POSTE_INSTALANDO"], [["*_LV"], "X_LV"], [["POSTE_INSTALANDO"], "poste_instalando"],
             [["A.B"], "AXB"], [["A.B*"], "A.BC"], [[], "QUALQUER"], [["POSTE_*"], "POSTE_REMOVENDO_LV"]]
    rr = subprocess.run([NODE, "-e", js], input=json.dumps(casos), capture_output=True, text=True, timeout=30)
    assert rr.returncode == 0, rr.stderr
    assert json.loads(rr.stdout) == [True, False, True, True, False, True, True, True]


def test_aba_cabos_postes_tem_teclado_e_painel_de_achados():
    ui = _ler("vinculacao_ui.js")
    for trecho in ["function tecladoCabosPostes(", "function atualizarAchados(", "class: 'vl-ach'", "st.kbModo", "painel.rolarPara("]:
        assert trecho in ui, trecho
    aba = ui.split("function abaCabosPostes(")[1].split("// ── tabelas editáveis")[0]
    assert "corpo.appendChild(st.elAch)" in aba                         # o painel de achados fica na aba certa (não na aba Vínculo)
    assert "corpo.appendChild(st.elAch)" not in ui.split("function abaVinculo(")[1].split("function atualizarAchados")[0]


def test_tabelas_muito_grandes_montam_so_a_janela_em_segundo_plano():
    js = _ler("resumo.js")
    assert "LIMIAR_PARCIAL = 300" in js and "hidratarJanela(type, 2500)" in js


def test_script_de_desempenho_existe_com_limites():
    t = open(os.path.join(os.path.dirname(config.STATIC_DIR), "scripts", "perf_resumo.js"), encoding="utf-8").read()
    for chave in ["LIMITES", "modal: vincular", "resumo: desfazer", "process.exit(falhou ? 1 : 0)"]:
        assert chave in t


@pytest.mark.skipif(not (NODE and os.environ.get("PERF_URL") and os.environ.get("PERF_TOKEN")), reason="defina PERF_URL e PERF_TOKEN (servidor no ar) para medir o desempenho")
def test_desempenho_com_450_linhas_em_cada_tabela():
    rr = subprocess.run([NODE, os.path.join(os.path.dirname(config.STATIC_DIR), "scripts", "perf_resumo.js"), os.environ["PERF_URL"], os.environ["PERF_TOKEN"], "450"], capture_output=True, text=True, timeout=240)
    assert rr.returncode == 0, rr.stdout + rr.stderr
