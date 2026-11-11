"""TASK-140 — relatórios ODI/ODD em PDF no formato da nota do SAP (cabeçalho, tabela verde, Calibri 11, rodapé Internal Use)."""
import fitz
import pytest

from services.exportar_sap_pdf import ErroPdf, formatar_qtd, gerar_pdf
from tests.test_fator_formula_109 import adm, banco, client  # noqa: F401

CAB = {"titulo": "-INST SECCION  U B088237 RECIFE", "nota": "9101578720", "sap": "E0483118", "bairro": "BREJO DA GUABIRABA", "municipio": "RECIFE",
       "utd": "CASA FORTE", "utep": "METROPOLITANO", "projetista": "GABRIEL SALES", "data": "11/11/2026", "id_investimento": "8.2", "id_investimento_odd": "9.9"}
LINHAS = [{"mdo": "M", "codigo": "500109", "desc": "CH SEC 15KV 630A 1P MAN SECO", "qtd": 3}, {"mdo": "N", "codigo": "SDEVU3602II", "desc": "INST CHAVE FACA UNIPOLAR - LV - (PA)", "qtd": "1.234,5"}]


def perto(a, b, tol=0.5):
    return all(abs(x - y) <= tol for x, y in zip(a, b))


def spans(pdf, pagina=0):
    saida = {}
    for b in fitz.open(stream=pdf, filetype="pdf")[pagina].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                saida[s["text"].strip()] = (round(s["origin"][0], 1), round(s["origin"][1], 1), round(s["size"], 1), s["color"])
    return saida


def test_formato_da_nota():
    s = spans(gerar_pdf(CAB, "ODI", LINHAS))
    assert perto(s["TÍTULO:"][:3], (18.4, 31.4, 11.0)) and perto(s["-INST SECCION  U B088237 RECIFE"][:2], (88.4, 31.4))
    assert perto(s["Nº SAP:"][:1], (235.9,)) and perto(s["E0483118"][:2], (300.4, 45.96)) and s["BAIRRO"][1] == 60.5 and s["PROJETISTA:"][1] == 89.5
    assert perto(s["ID INVESTIMENTO - 8.2"][:1], (443.5,)) and s["ID INVESTIMENTO - 8.2"][3] == 0x404040 and s["TIPO DE ORDEM  -  ODI"][1] == 60.5
    assert s["TIPO"][3] == 0xFFFFFF and perto(s["DESCRIÇÃO DO MATERIAL"][:2], (300.4, 118.4))
    assert perto(s["3,00"][:2], (199.2, 133.0)) and s["1.234,50"][1] == 147.5 and s["1600"][1] == 133.0 and s["1000"][1] == 147.5
    assert s["Internal Use"][2:] == (12.0, 0x8000) and s["Internal Use"][1] == 837.7


def test_odd_usa_id_proprio_e_a_cor_verde_do_titulo():
    pdf = gerar_pdf(CAB, "ODD", LINHAS)
    assert "ID INVESTIMENTO - 9.9" in spans(pdf) and "TIPO DE ORDEM  -  ODD" in spans(pdf)
    assert "ID INVESTIMENTO - 8.2" in spans(gerar_pdf({**CAB, "id_investimento_odd": ""}, "ODD", LINHAS))
    pix = fitz.open(stream=pdf, filetype="pdf")[0].get_pixmap(dpi=72)
    assert pix.pixel(450, 114)[:3] == (0, 163, 66)                                           # faixa do título: #00A443


def test_varias_paginas_repete_titulo_e_rodape():
    pdf = gerar_pdf(CAB, "ODI", [{"mdo": "M", "codigo": str(1000 + i), "desc": "ITEM " + str(i), "qtd": i} for i in range(80)])
    d = fitz.open(stream=pdf, filetype="pdf")
    assert d.page_count == 2
    s2 = spans(pdf, 1)
    assert "CÓDIGO" in s2 and "Internal Use" in s2 and "TÍTULO:" not in s2


def test_formatar_qtd_e_validacao():
    assert [formatar_qtd(v) for v in (3, "5.84", "1.234,5", 0, None)] == ["3,00", "5,84", "1.234,50", "0,00", "0,00"]
    with pytest.raises(ErroPdf):
        gerar_pdf(CAB, "XYZ", LINHAS)
    with pytest.raises(ErroPdf):
        gerar_pdf(CAB, "ODI", "x")


def test_rota_devolve_pdf(client):
    r = client.post("/api/orcamento/exportar-sap-pdf", json={"tipo": "ODI", "cabecalho": CAB, "linhas": LINHAS}, headers=adm())
    assert r.status_code == 200 and r.headers["content-type"] == "application/pdf" and r.content[:4] == b"%PDF"
    assert client.post("/api/orcamento/exportar-sap-pdf", json={"tipo": "ODX", "cabecalho": CAB, "linhas": LINHAS}, headers=adm()).status_code == 400
    assert client.post("/api/orcamento/exportar-sap-pdf", json={"tipo": "ODI", "cabecalho": CAB, "linhas": LINHAS}).status_code in (401, 403)
