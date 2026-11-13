"""TASK-157 — Comparar com REC / .obra / PDFs PB / PDF SAP e Importar PDFs com os formatos PB e SAP."""
import json
import os
import shutil
import subprocess

import fitz
import pytest

import config
from services import importar_sap_pdf as isp
from services.exportar_sap_pdf import gerar_pdf
from tests.test_fator_formula_109 import adm, banco, client  # noqa: F401

NODE = shutil.which("node")
CAB = {"titulo": "-INST SECCION  U B088237 RECIFE", "nota": "9101578720", "sap": "E0483118", "id_investimento": "8.2"}
LINHAS = [{"mdo": "M", "codigo": "500109", "desc": "CH SEC 15KV 630A 1P MAN SECO", "qtd": 3},
          {"mdo": "M", "codigo": "2203008", "desc": "CABO COBRE NU C2A MD 35 MM2", "qtd": "1,95"},
          {"mdo": "N", "codigo": "SDEVU3602II", "desc": "INST CHAVE FACA UNIPOLAR - LV - (PA)", "qtd": "1.234,5"}]


def ler(*partes):
    return open(os.path.join(*partes), encoding="utf-8").read()


# ───────── leitura do PDF SAP ─────────
def test_le_o_pdf_gerado_pelo_programa_odi_e_odd():
    r = isp.ler_pdf(gerar_pdf(CAB, "ODI", LINHAS))
    assert [x["tipo"] for x in r["secoes"]] == ["ODI"] and r["nota"] == "9101578720"
    assert [(l["mdo"], l["codigo"], l["qtd"], l["centro"]) for l in r["secoes"][0]["linhas"]] == [("MATERIAL", "500109", 3.0, "1600"), ("MATERIAL", "2203008", 1.95, "1600"),
                                                                                         ("MAO-DE-OBRA", "SDEVU3602II", 1234.5, "1000")]
    assert r["secoes"][0]["linhas"][2]["desc"] == "INST CHAVE FACA UNIPOLAR - LV - (PA)"
    assert isp.ler_pdf(gerar_pdf(CAB, "ODD", LINHAS))["secoes"][0]["tipo"] == "ODD"


def test_le_pdf_com_varias_paginas_e_descricao_longa():
    muitas = [{"mdo": "M", "codigo": str(1000000 + i), "desc": f"MATERIAL NUMERO {i} " + "X" * 60, "qtd": i + 0.5} for i in range(130)]
    r = isp.ler_pdf(gerar_pdf(CAB, "ODI", muitas))
    linhas = r["secoes"][0]["linhas"]
    assert len(r["secoes"]) == 1 and len(linhas) == 130 and linhas[129]["codigo"] == "1000129" and linhas[129]["qtd"] == 129.5 and linhas[0]["desc"].startswith("MATERIAL NUMERO 0 ")


def pdf_de_fora(linhas_texto):
    """PDF 'feito por fora': decorações, hífen U+2010, descrição com espaços repetidos, ordem diferente — cada célula numa linha de texto."""
    doc = fitz.open()
    pg = doc.new_page(width=595, height=842)
    y = 20
    for t in linhas_texto:
        pg.insert_text((20, y), t, fontsize=9)
        y += 11
    return doc.tobytes()


def test_le_pdf_feito_por_fora_no_mesmo_formato():
    textos = ["111", "111", "TÍTULO:", "-INST SECCION  U B088237 RECIFE", "Nº NOTA:", "9101578720", "TIPO DE ORDEM  -  ODD", "111111",
              "TIPO", "CÓDIGO", "QTD", "CENTRO", "DESCRIÇÃO DO MATERIAL",
              "M", "500109", "3,00", "1600", "CH SEC 15KV 630A 1P MAN SECO",
              "M", "2402000", "6,00", "1600", "CARTUCHO AZUL     FERRAMENTA AMPACT",
              "N", "SDEVU1434II", "5,62", "1000", "INST COND NU LINHA VIVA - (PA)", "#Internal Use"]
    r = isp.ler_pdf(pdf_de_fora(textos))
    linhas = r["secoes"][0]["linhas"]
    assert r["secoes"][0]["tipo"] == "ODD" and len(linhas) == 3
    assert linhas[1]["desc"] == "CARTUCHO AZUL FERRAMENTA AMPACT" and linhas[2]["codigo"] == "SDEVU1434II" and linhas[2]["qtd"] == 5.62
    assert "Internal" not in linhas[2]["desc"]
    for h in "-\u2010\u2011\u2013\u2212":                                                       # o PDF SAP traz hífen "bonito" (U+2010) no cabeçalho
        assert isp._RE_TIPO_ORDEM.search(f"TIPO DE ORDEM  {h}  ODI").group(1) == "ODI"


def pdf_icp_combinado():
    """Como o relatório do ICP: ODI em duas páginas (a 2ª só com a continuação da tabela, sem cabeçalho) e o ODD numa 3ª página com cabeçalho. Quantidades 'soltas' (3, 5,2)."""
    doc = fitz.open()

    def pagina(textos):
        pg = doc.new_page(width=595, height=842)
        for i, t in enumerate(textos):
            pg.insert_text((20, 20 + i * 11), t, fontsize=9)
    pagina(["TÍTULO:", "-PO RL U B088255 RECIFE", "Nº NOTA:", "9101578718", "TIPO DE ORDEM  -  ODI", "TIPO", "CÓDIGO", "QTD", "CENTRO", "DESCRIÇÃO DO MATERIAL",
            "M", "2201014", "0,45", "1600", "FIO COBRE NU", "N", "SDEVU1434II", "15,91", "1000", "INST COND NU  LINHA VIVA  - (PA)", "#Internal Use"])
    pagina(["M", "3419284", "6,00", "1600", "SUPORTE P/EQUIP CHAVE", "M", "2212022", "3", "1600", "CABO ALUM XLPE", "M", "2206000", "5,2", "1600", "CABO NU ACO-COBRE 2 AWG", "#Internal Use"])
    pagina(["TÍTULO:", "Nº NOTA:", "9101578718", "TIPO DE ORDEM  -  ODD", "TIPO", "CÓDIGO", "QTD", "CENTRO", "DESCRIÇÃO DO MATERIAL",
            "M", "7910008", "13,95", "1600", "SUCATA CABO/FIO COBRE NU", "N", "SDEVU3203RD", "1,00", "1000", "RETIRAR EST CRUZ", "#Internal Use"])
    return doc.tobytes()


def test_um_pdf_com_odi_e_odd_vira_duas_secoes():
    r = isp.ler_pdf(pdf_icp_combinado())
    assert [(x["tipo"], len(x["linhas"])) for x in r["secoes"]] == [("ODI", 5), ("ODD", 2)] and r["nota"] == "9101578718"    # a 2ª página (sem cabeçalho) continua o ODI
    assert [l["qtd"] for l in r["secoes"][0]["linhas"]] == [0.45, 15.91, 6.0, 3.0, 5.2]                                     # 3 e 5,2 sem casas decimais fixas
    res = isp.para_resultado([r])
    assert [x["operacao"] for x in res] == ["I"] * 5 + ["R"] * 2 and res[-1]["codigo"] == "SDEVU3203RD"
    with pytest.raises(isp.ErroSapPdf, match="dois PDFs ODD"):                                                              # o ODD do combinado + um ODD solto
        isp.para_resultado([r, isp.ler_pdf(gerar_pdf(CAB, "ODD", LINHAS))])


def test_recusa_o_que_nao_e_pdf_sap():
    with pytest.raises(isp.ErroSapPdf, match="PDF válido"):
        isp.ler_pdf(b"isto nao e pdf")
    with pytest.raises(isp.ErroSapPdf, match="TIPO DE ORDEM"):
        isp.ler_pdf(pdf_de_fora(["Relatório qualquer", "M", "1", "2,00", "1600", "X"]))
    with pytest.raises(isp.ErroSapPdf, match="linhas de material"):
        isp.ler_pdf(pdf_de_fora(["TIPO DE ORDEM - ODI", "TIPO", "CÓDIGO"]))


def test_para_resultado_odi_vira_i_odd_vira_r_e_recusa_tipo_repetido():
    odi, odd = isp.ler_pdf(gerar_pdf(CAB, "ODI", LINHAS)), isp.ler_pdf(gerar_pdf(CAB, "ODD", LINHAS[:1]))
    res = isp.para_resultado([odi, odd])
    assert [x["operacao"] for x in res] == ["I", "I", "I", "R"] and res[0] == {"operacao": "I", "mdo": "MATERIAL", "codigo": "500109", "desc_codigo": "CH SEC 15KV 630A 1P MAN SECO", "filtro": "", "total": 3.0}
    assert [x["operacao"] for x in isp.para_resultado([odi])] == ["I", "I", "I"]                    # só o ODI: só instalação
    with pytest.raises(isp.ErroSapPdf, match="dois PDFs ODI"):
        isp.para_resultado([odi, odi])


# ───────── rota ─────────
def envia(client, *pdfs):
    return client.post("/api/importar-sap-pdf", files=[("arquivos", (f"a{i}.pdf", p, "application/pdf")) for i, p in enumerate(pdfs)], headers=adm())


def test_rota_importar_sap_pdf(client):
    odi, odd = gerar_pdf(CAB, "ODI", LINHAS), gerar_pdf(CAB, "ODD", LINHAS[:2])
    r = envia(client, odi, odd)
    assert r.status_code == 200, r.text
    d = r.json()
    assert len(d["resultado"]) == 5 and [a["tipo"] for a in d["arquivos"]] == ["ODI", "ODD"] and d["arquivos"][0]["itens"] == 3 and d["arquivos"][0]["nota"] == "9101578720"
    assert {x["operacao"] for x in d["resultado"]} == {"I", "R"}
    assert envia(client, odi).status_code == 200
    assert envia(client, odi, odi).status_code == 400 and "dois PDFs ODI" in envia(client, odi, odi).json()["detail"]
    assert envia(client, odi, odd, odi).status_code == 400
    ruim = envia(client, b"nao-pdf")
    assert ruim.status_code == 400 and "a0.pdf" in ruim.json()["detail"]


# ───────── comparação: linhas → itens ─────────
@pytest.mark.skipif(not NODE, reason="node não instalado")
def test_itens_de_linhas_consolida_e_normaliza():
    mod = os.path.join(config.STATIC_DIR, "orcamento_comparar.js")
    codigo = f"""const C=require({json.dumps(mod)});
const it=C.itensDeLinhas([{{operacao:'I',mdo:'MATERIAL',codigo:'500109',desc_codigo:'CH',total:3}},{{operacao:'I',mdo:'M',codigo:'500109',total:'1,5'}},
 {{operacao:'R',mdo:'MAO-DE-OBRA',codigo:'sdevu3602ii',desc:'INST',qtd:1}},{{operacao:'I',mdo:'MATERIAL',codigo:'  '}},null]);
console.log(JSON.stringify(it));"""
    saida = subprocess.run([NODE, "-e", codigo], capture_output=True, text=True, timeout=30)
    assert saida.returncode == 0, saida.stderr
    itens = json.loads(saida.stdout)
    assert [(i["operacao"], i["mdo"], i["codigo"], i["qtd"], i["desc"]) for i in itens] == [("I", "MATERIAL", "500109", 4.5, "CH"), ("R", "MAO-DE-OBRA", "sdevu3602ii", 1, "INST")]
    assert itens[0]["chave"] == "I|MATERIAL|500109"


# ───────── telas ─────────
def test_comparar_pede_o_formato_e_importar_pdfs_tem_pb_e_sap():
    h = ler(config.STATIC_DIR, "resultado_orcamento.html")
    assert h.count('onclick="abrirComparar()"') == 2 and "file-rec-comparar').click(); return;" in h            # o REC segue o fluxo de sempre
    for trecho in ('id="modal-comparar"', 'id="cmp-file-obra"', 'id="cmp-file-pb-orc"', 'id="cmp-file-pb-lista"', 'id="cmp-file-sap"',
                   "projeto_trabalho", "/api/importar-sap-pdf", "/api/importar-rec-pdf", "definirReferenciaDeLinhas", "itensDeLinhas"):
        assert trecho in h, trecho
    assert 'id="imp-fmt-pb"' in h and 'id="imp-fmt-sap"' in h and 'id="file-pdf-sap"' in h and "importarPdfsSap" in h and "Este .obra não tem resultado de orçamento gravado" in h
    assert "document.getElementById('file-rec-comparar').click()" not in h.split("function abrirComparar")[0].split('<div class="barra-comparar"')[1][:600]


def test_rota_le_um_pdf_com_odi_e_odd_juntos(client):
    r = client.post("/api/importar-sap-pdf", files=[("arquivos", ("RELATORIO_ICP.pdf", pdf_icp_combinado(), "application/pdf"))], headers=adm())
    assert r.status_code == 200, r.text
    d = r.json()
    assert [(a["nome"], a["tipo"], a["itens"]) for a in d["arquivos"]] == [("RELATORIO_ICP.pdf", "ODI", 5), ("RELATORIO_ICP.pdf", "ODD", 2)]
    assert [x["operacao"] for x in d["resultado"]] == ["I"] * 5 + ["R"] * 2
    dois = client.post("/api/importar-sap-pdf", files=[("arquivos", ("a.pdf", pdf_icp_combinado(), "application/pdf")), ("arquivos", ("b.pdf", gerar_pdf(CAB, "ODD", LINHAS), "application/pdf"))], headers=adm())
    assert dois.status_code == 400 and "dois PDFs ODD" in dois.json()["detail"]


def test_titulo_escrito_depois_da_tabela_nao_entra_na_descricao():
    """O PDF do ICP escreve o valor do TÍTULO depois da tabela: a leitura por posição não deixa esse texto virar parte da descrição da última linha."""
    doc = fitz.open()
    pg = doc.new_page(width=595, height=842)
    colunas = (47, 116, 199, 255, 300)
    pg.insert_text((20, 20), "TIPO DE ORDEM  -  ODD", fontsize=9)
    for i, (mdo, cod, qtd, centro, desc) in enumerate([("M", "7910008", "13,95", "1600", "SUCATA CABO/FIO COBRE NU"), ("N", "SDEVU3203RD", "1,00", "1000", "RETIRAR EST CRUZ SIMP ISOL PINO/PILAR LV")]):
        for x, t in zip(colunas, (mdo, cod, qtd, centro, desc)):
            pg.insert_text((x, 133 + i * 14.5), t, fontsize=9)
    pg.insert_text((88.4, 31.4), "-PO RL  U B088255 RECIFE", fontsize=9)                 # no texto do PDF isto vem DEPOIS da tabela
    pg.insert_text((267, 837), "#Internal Use", fontsize=9)
    r = isp.ler_pdf(doc.tobytes())
    assert [l["desc"] for l in r["secoes"][0]["linhas"]] == ["SUCATA CABO/FIO COBRE NU", "RETIRAR EST CRUZ SIMP ISOL PINO/PILAR LV"]
