"""TASK-176 — perfil de exportação por projeto: cabeçalho, centro, PDF/XLSX, .obra e OBRA.*. Sem perfil, saídas idênticas às de antes (baseline gravada antes da mudança)."""
import io
import json
import os

import pytest
from openpyxl import load_workbook

import config
from services import exportar_sap_pdf as pdf
from services import exportar_xlsx as xl
from services import obra_params
from services import orcamento_calc as calc
from services import perfil_exportacao as pe
from services import projeto_obra
from services import vinculacao_regras as vr
from tests._resumo_pdf_176 import resumo

BASE = json.load(open(os.path.join(os.path.dirname(__file__), "dados", "exportacao_padrao_176.json"), encoding="utf-8"))
CAB = {"titulo": "Projeto X", "nota": "123", "sap": "E01", "bairro": "Centro", "municipio": "Recife", "utd": "UTD1", "utep": "UTEP1", "projetista": "Fulano", "data": "16/11/2026",
       "id_investimento": "INV1", "id_investimento_odd": "INV2", "contrato": "CT-77"}
ODI = [{"mdo": "MATERIAL" if i % 2 else "MAO-DE-OBRA", "codigo": str(1000 + i), "desc": f"ITEM {i} DESCRICAO", "qtd": i * 1.5} for i in range(60)]
ODD = [{"mdo": "M", "codigo": "9" + str(i), "desc": "REMOVE", "qtd": i} for i in range(5)]
CAB_BASE = {k: v for k, v in CAB.items() if k != "contrato"}
PERFIL = {"cabecalho": [{"chave": "titulo", "rotulo": "TÍTULO:", "coluna": "esq"}, {"chave": "nota", "rotulo": "Nº NOTA:", "coluna": "esq"}, {"chave": "contrato", "rotulo": "CONTRATO:", "tela": "Contrato", "coluna": "esq", "obra": True},
                       {"chave": "bairro", "rotulo": "BAIRRO", "coluna": "esq", "obra": True}, {"chave": "projetista", "rotulo": "PROJETISTA:", "coluna": "esq"},
                       {"chave": "sap", "rotulo": "Nº SAP:", "coluna": "meio"}, {"chave": "municipio", "rotulo": "MUNICIPIO:", "coluna": "meio", "obra": True}, {"chave": "data", "rotulo": "DATA:", "coluna": "meio"},
                       {"chave": "id_investimento", "rotulo": "ID INVESTIMENTO", "coluna": "id"}],
          "centro": {"MATERIAL": "2000", "MAO-DE-OBRA": "3000"}, "xlsx": {"colunas": ["codigo", "desc", "qtd", "centro"]}, "pdf": {"cor_titulo": "#0055AA", "rodape": "Uso interno"}}


def _xlsx(odi, odd, perfil=None):
    wb = load_workbook(io.BytesIO(xl.gerar_xlsx(odi, odd, perfil)))
    return {ws.title: [[c.value for c in r] for r in ws.iter_rows()] for ws in wb}


def test_sem_perfil_pdf_e_xlsx_iguais_aos_de_antes():
    assert resumo(pdf.gerar_pdf_unico(CAB_BASE, ODI, ODD)) == BASE["unico"]
    assert resumo(pdf.gerar_pdf(CAB_BASE, "ODI", ODI[:3])) == BASE["odi"] and resumo(pdf.gerar_pdf(CAB_BASE, "ODD", ODD)) == BASE["odd"]
    assert _xlsx(ODI[:4], ODD[:2]) == BASE["xlsx"]
    # perfil padrão explícito (ou {}) = o mesmo
    for perfil in (pe.PADRAO, {}, None):
        assert resumo(pdf.gerar_pdf_unico(CAB_BASE, ODI, ODD, perfil)) == BASE["unico"]
        assert _xlsx(ODI[:4], ODD[:2], perfil) == BASE["xlsx"]


def test_perfil_do_projeto_muda_centro_cabecalho_e_colunas():
    x = _xlsx(ODI[:2], ODD[:1], PERFIL)
    assert x["ODI"][0] == ["CÓDIGO", "DESCRIÇÃO", "QUANTIDADE", "CENTRO"] and x["ODI"][1][3] == "3000" and x["ODI"][2][3] == "2000" and x["ODD"][1][3] == "2000"
    doc = pdf.gerar_pdf_unico(CAB, ODI[:3], ODD[:1], PERFIL)
    import fitz
    d = fitz.open(stream=doc, filetype="pdf")
    t = d[0].get_text()
    assert "CONTRATO:" in t and "CT-77" in t and "UTD:" not in t and "UTEP:" not in t and "2000" in t and "3000" in t and "1600" not in t and "Uso interno" in t
    assert resumo(doc) != BASE["unico"]
    cor = [dr["fill"] for dr in d[0].get_drawings() if dr.get("fill") and abs(dr["fill"][2] - 0.6667) < 0.01]
    assert cor, "a faixa do título usa a cor do perfil"


def test_obra_grava_as_chaves_do_perfil_e_obra_antigo_abre_igual(monkeypatch):
    import routers.vinculacao_config as vc
    # sem perfil no projeto: só os campos de sempre (campo desconhecido é descartado, como antes)
    assert projeto_obra.sap_limpo({"sap": "E1", "contrato": "CT-77", "campo_estranho": "x", "utd": "U"}, "081") == {"sap": "E1", "utd": "U"}
    assert projeto_obra.sap_limpo({"sap": "E1", "bairro": "B"}) == {"sap": "E1", "bairro": "B"}                      # arquivo antigo igual
    monkeypatch.setattr(vc, "perfil_do_projeto", lambda p: PERFIL if p == "T1" else None)
    assert projeto_obra.sap_limpo({"sap": "E1", "contrato": "CT-77", "campo_estranho": "x", "utd": "U", "nota": "n"}, "T1") == {"sap": "E1", "utd": "U", "contrato": "CT-77"}
    texto = projeto_obra.montar({"nome": "P", "sap": {"contrato": "CT-77", "sap": "E1"}}, None, "T1")
    assert json.loads(texto)["projeto_trabalho"]["sap"] == {"sap": "E1", "contrato": "CT-77"}
    assert pe.campos_sap(PERFIL) == ["contrato", "bairro", "projetista", "sap", "municipio", "data", "id_investimento"]
    assert pe.campos_sap() == ["sap", "bairro", "municipio", "utd", "utep", "projetista", "data", "id_investimento", "id_investimento_odd"]


def test_obra_contrato_disponivel_para_a_base_e_nomes_atuais_no_padrao():
    obra = {"municipio": "Recife", "bairro": "Boa Vista", "utd": "Casa Forte", "utep": "Metropolitano", "contexto": "X", "contrato": "ct 77"}
    assert obra_params.montar(obra, "") == obra_params.montar(obra, "", pe.PADRAO) == {"OBRA.MUNICIPIO": "RECIFE", "OBRA.BAIRRO": "BOA_VISTA", "OBRA.UTD": "CASA_FORTE", "OBRA.UTEP": "METROPOLITANO", "OBRA.CONTEXTO": "X"}
    m = obra_params.montar(obra, "", PERFIL)
    assert m["OBRA.CONTRATO"] == "CT_77" and "OBRA.UTD" not in m and m["OBRA.MUNICIPIO"] == "RECIFE" and m["OBRA.CONTEXTO"] == "X"
    # a base usa OBRA.CONTRATO como parâmetro
    base = [{"ativo": "X=X(obra.contrato=CT_77)", "desc_ativo": "d", "componente": "X", "projeto": "T", "mdo": "MATERIAL", "codigo": "a", "desc_codigo": "d", "fator_i": 1, "fator_r": 0, "filtro": "",
             "origem": "OUTROS", "lv": "A"},
            {"ativo": "", "desc_ativo": "", "componente": "X", "projeto": "T", "mdo": "MATERIAL", "codigo": "b", "desc_codigo": "d", "fator_i": 1, "fator_r": 0, "filtro": "", "origem": "OUTROS", "lv": "A"}]
    r = calc.processar_calculo([], [{"ativo": "1-X", "operacao": "I", "entidade": "0"}], "T", base, obra={"contrato": "ct 77"}, perfil_exportacao=PERFIL)
    assert r["resultado"] is not None


def test_validacao_e_normalizacao():
    assert pe.validar(None) == [] and pe.validar(PERFIL) == [] and pe.validar(pe.PADRAO) == []
    assert len(pe.validar({"cabecalho": [{"chave": "Contrato"}, {"chave": "a"}, {"chave": "a"}], "centro": 5, "pdf": {"cor_titulo": "verde"}, "xlsx": {"colunas": ["x"]}, "arquivos": ["zip"]})) >= 6
    assert pe.validar({"cabecalho": [{"chave": f"c{i}", "coluna": "esq"} for i in range(6)]})
    assert vr.validar_config({"exportacao": {"centro": 5}}) and vr.validar_config({"exportacao": PERFIL}) == [] and vr.normalizar_config({"exportacao": PERFIL})["exportacao"] == PERFIL
    assert pe.oferece("rec") and not pe.oferece("csv", {"arquivos": ["rec"]}) and pe.normalizar({})["centro"] == pe.PADRAO["centro"]


def test_rota_usa_o_perfil_enviado_ou_o_do_projeto():
    from fastapi.testclient import TestClient
    import app as appmod
    from middleware.auth_middleware import create_jwt_token
    c = TestClient(appmod.app)
    h = {"Authorization": f"Bearer {create_jwt_token('u1', 'u1@x.com', 'operador')}"}
    r = c.post("/api/orcamento/exportar-xlsx", json={"odi": ODI[:2], "odd": [], "perfil": PERFIL}, headers=h)
    assert r.status_code == 200
    x = load_workbook(io.BytesIO(r.content))["ODI"]
    assert [c.value for c in x[1]] == ["CÓDIGO", "DESCRIÇÃO", "QUANTIDADE", "CENTRO"]
    assert c.post("/api/orcamento/exportar-xlsx", json={"odi": ODI[:2], "odd": [], "perfil": {"centro": 5}}, headers=h).status_code == 400
    r = c.post("/api/orcamento/exportar-sap-pdf", json={"cabecalho": CAB, "odi": ODI[:3], "odd": [], "perfil": PERFIL}, headers=h)
    assert r.status_code == 200 and r.content[:4] == b"%PDF"
    sem = c.post("/api/orcamento/exportar-xlsx", json={"odi": ODI[:2], "odd": []}, headers=h)
    assert [c.value for c in load_workbook(io.BytesIO(sem.content))["ODI"][1]] == ["TIPO MDO", "CÓDIGO", "QUANTIDADE", "CENTRO", "DESCRIÇÃO"]


def _node(expr, **dados):
    import shutil
    import subprocess
    js = f"const PE=require('./static/perfil_exportacao.js'),SAP=require('./static/orcamento_sap.js');const D={json.dumps(dados)};console.log(JSON.stringify({expr}));"
    r = subprocess.run([shutil.which("node"), "-e", js], cwd=config.BASE_DIR, capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


@pytest.mark.skipif(not __import__("shutil").which("node"), reason="sem node")
def test_paridade_com_o_javascript():
    ruim = {"cabecalho": [{"chave": "Contrato"}, {"chave": "a"}, {"chave": "a"}], "centro": 5, "pdf": {"cor_titulo": "verde"}, "xlsx": {"colunas": ["x"]}, "arquivos": ["zip"]}
    for perfil in (None, {}, pe.PADRAO, PERFIL, {"centro": {"material": "9"}, "arquivos": ["rec"]}):
        assert pe.normalizar(perfil) == _node("PE.normalizar(D.p)", p=perfil), perfil
        assert pe.campos_sap(perfil) == _node("PE.camposSap(D.p)", p=perfil)
        assert pe.chaves_obra(perfil) == _node("PE.chavesObra(D.p)", p=perfil)
        for m in ("M", "N", "x", ""):
            assert pe.centro_de(m, perfil) == _node("PE.centroDe(D.m, D.p)", m=m, p=perfil), (m, perfil)
        for a in pe.ARQUIVOS:
            assert pe.oferece(a, perfil) == _node("PE.oferece(D.a, D.p)", a=a, p=perfil)
    assert pe.validar(ruim) == _node("PE.validar(D.p)", p=ruim) and pe.validar(PERFIL) == _node("PE.validar(D.p)", p=PERFIL) == []
    assert _node("SAP.centroDe('MATERIAL')") == "1600" and _node("SAP.centroDe('MAO-DE-OBRA')") == "1000" and _node("SAP.centroDe('-')") == "1000" and _node("SAP.centroDe('XX')") == ""


def test_ligacoes_na_tela():
    ler = lambda *p: open(os.path.join(config.STATIC_DIR, *p), encoding="utf-8").read()
    resumo_html, resultado = ler("resumo.html"), ler("resultado_orcamento.html")
    assert resumo_html.index("/static/perfil_exportacao.js") < resumo_html.index("/static/projeto.js")
    assert resultado.index("/static/perfil_exportacao.js") < resultado.index("/static/projeto.js") and "relCampos" in resultado and "perfilPronto" in resultado
    assert "projeto: localStorage.getItem('projeto_selecionado_codigo')" in resultado
    ui = ler("vinculacao_ui.js")
    assert "abaExportacao" in ui and "Pré-visualizar uma página" in ui
    assert "camposSap" in ler("projeto.js") and "PerfilExportacao" in ler("resumo.js") and "camposDoPerfil" in ler("obra_params.js")
