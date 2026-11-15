"""TASK-167 — status dos parâmetros do ativo (preenchidos × que faltam), balão e cores opcionais (static/ativo_status.js)."""
import json
import os
import subprocess

import pytest

import config
from tests import oraculo_tela as o

pytestmark = pytest.mark.skipif(not o.NODE, reason="node não instalado")
MOD = os.path.join(config.STATIC_DIR, "ativo_status.js")
AP = os.path.join(config.STATIC_DIR, "ativo_parametros.js")


def L(ativo, comp="", cod=""):
    return {"ativo": ativo, "componente": comp, "codigo": cod, "projeto": "PERNAMBUCO", "desc_ativo": "d", "mdo": "MATERIAL", "desc_codigo": "d", "fator_i": 1, "fator_r": 0, "filtro": "", "origem": "OUTROS"}


BASE = [L("N1=N1(cabo=XLPE,poste=D,tipocabo)"), L("N1(poste=%D)", "N1.D", "10"), L("N4=N4(cabo,tipocabo,cabo2)"), L("N4(cabo=%A35P%)", "N4.A", "11"),
        L("FLBIM=FLBIM(poste=D,obra.regiao=CENTRO)"), L("FLBIM(obra.regiao=CENTRO|OESTE)", "F.A", "12"), L("FLBIM(obra.regiao=SUL)"), L("N3", "N3", "13")]


def st(final, pedido=None, obra=None):
    codigo = (f"const S=require({json.dumps(MOD)}),AP=require({json.dumps(AP)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));"
              "const sig=AP.assinaturasDoProjeto(e.base,'PERNAMBUCO');console.log(JSON.stringify(S.statusDoAtivo(e.f,e.p,sig,e.o)));")
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps({"base": BASE, "f": final, "p": pedido, "o": obra}), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def por_nome(s):
    return {p["nome"]: p for p in s["params"]}


def test_padrao_conta_como_preenchido_e_diz_a_origem():
    s = st("N1(poste=9400D,tipocabo=P)", "N1(poste=9400D)")
    p = por_nome(s)
    assert s["completo"] and s["faltam"] == 0
    assert p["poste"]["origem"] == "pedido" and p["tipocabo"]["origem"] == "vinculo" and p["cabo"]["origem"] == "padrao" and p["cabo"]["valor"] == "XLPE"


def test_falta_quando_nao_ha_valor_nem_padrao_com_motivo():
    s = st("N4")
    assert not s["completo"] and s["faltam"] == 3
    assert all(p["falta"] and "nenhum cabo vinculado" in p["motivo"] for p in s["params"])


def test_vinculo_preenche_parte_e_o_resto_falta():
    s = st("N4(cabo=A35P,tipocabo=P)", "N4")
    p = por_nome(s)
    assert s["faltam"] == 1 and p["cabo2"]["falta"] and p["cabo"]["origem"] == "vinculo" and not p["tipocabo"]["falta"]


def test_ativo_sem_parametros_nao_tem_status():
    assert st("N3") is None and st("[A-B]") is None and st("INEXISTENTE") is None


def test_valor_invalido_nao_conta_como_completo():
    s = st("N4(cabo=ZZZ,tipocabo=P,cabo2=A35P)")
    assert s["invalidos"] == 1 and not s["completo"] and por_nome(s)["cabo"]["invalido"]


def test_parametro_de_obra_vem_da_obra_ou_do_padrao():
    a = st("FLBIM", None, {"OBRA.REGIAO": "SUL"})
    p = por_nome(a)["obra.regiao"]
    assert p["origem"] == "obra" and p["valor"] == "SUL" and a["completo"]
    b = st("FLBIM", None, {})
    assert por_nome(b)["obra.regiao"]["origem"] == "padrao" and b["completo"]


def test_estado_da_linha():
    codigo = (f"const S=require({json.dumps(MOD)});console.log(JSON.stringify([S.estadoDaLinha([]),S.estadoDaLinha([null]),S.estadoDaLinha([{{completo:true}}]),S.estadoDaLinha([{{completo:true}},{{completo:false}}])]));")
    r = subprocess.run([o.NODE, "-e", codigo], capture_output=True, text=True, timeout=30)
    assert json.loads(r.stdout) == [None, None, "ok", "falta"]


def test_ligacoes_da_tela_e_rota_sem_cache():
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    js = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    assert html.index("/static/ativo_parametros.js") < html.index("/static/ativo_status.js") < html.index("/static/resumo.js")
    for chave in ("cfg-colorir-params", "cfg-avancado", "par-balao", "body.colorir-parametros #table-outros .inp-ativo.par-falta"):
        assert chave in html, chave
    for chave in ("aplicarCoresParametros", "definirCorParametros", "data-par-tipo", "AtivoStatus.estadoDaLinha", "colorir_ativos_parametros"):
        assert chave in js, chave
    assert 'type="checkbox" id="cfg-colorir-params"' not in html          # botão (click), não checkbox: preferência de tela não marca o projeto como alterado
