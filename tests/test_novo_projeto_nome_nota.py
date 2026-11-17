"""Pedido de 2026-11-17 — Novo projeto: a NOTA do SAP é o nome do projeto (e do .obra); o título do SAP fica guardado só para os relatórios."""
import json
import os

import config
from services import projeto_obra


def ler(*p):
    return open(os.path.join(config.STATIC_DIR, *p), encoding="utf-8").read()


def test_obra_guarda_o_titulo_do_sap_sem_ele_ser_o_nome():
    sap = {"sap": "E1", "titulo": "OLCF -INST SECCION U B088237 RECIFE", "nota": "9101578720", "utd": "U"}
    limpo = projeto_obra.sap_limpo(sap)
    assert limpo == {"sap": "E1", "titulo": "OLCF -INST SECCION U B088237 RECIFE", "utd": "U"}             # a nota não é gravada no sap: ela é o número da obra
    texto = projeto_obra.montar({"nome": "9101578720", "numero_obra": "9101578720", "sap": sap}, None, "081")
    t = json.loads(texto)["projeto_trabalho"]
    assert t["nome"] == "9101578720" and t["numero_obra"] == "9101578720" and t["sap"]["titulo"].startswith("OLCF")
    assert projeto_obra.nome_arquivo(t["nome"]) == "9101578720.obra"
    lido = projeto_obra.ler(texto, "081")
    assert lido["nome"] == "9101578720" and lido["sap"]["titulo"].startswith("OLCF")
    assert len(projeto_obra.sap_limpo({"titulo": "x" * 500})["titulo"]) == projeto_obra.LIMITE_TEXTO


def test_obra_antigo_sem_titulo_abre_igual():
    texto = projeto_obra.montar({"nome": "Rua das Flores", "numero_obra": "55", "sap": {"sap": "E1"}}, None, "081")
    assert projeto_obra.ler(texto, "081")["sap"] == {"sap": "E1"}


def test_ligacoes_na_tela():
    js, res = ler("projeto.js"), ler("resultado_orcamento.html")
    assert "if (c.nota) { cNome.input.value = c.nota; cNum.input.value = c.nota; }" in js and "if (c.titulo) sap.titulo = c.titulo;" in js
    assert "sap.titulo = cab.titulo.trim()" in js                                                       # o título digitado nos relatórios (diferente do nome) fica no projeto
    assert "GridRProjeto.sap().titulo || GridRProjeto.nome()" in res                                  # o relatório sugere o título do SAP; sem ele, o nome
