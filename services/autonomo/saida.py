"""services/autonomo/saida.py — Grava os resultados de uma execução autônoma em pasta (TASK-031, fase B): JSON + CSV (UTF-8 com BOM, abre no Excel)."""
import csv
import json
import os
import re

_SEGURO = re.compile(r"[^A-Za-z0-9_.\-]+")


def nome_seguro(texto: str, padrao: str = "item") -> str:
    """Nome de arquivo/pasta sem separadores, `..` ou caracteres estranhos (o código do projeto e o nome do arquivo vêm de fora)."""
    limpo = _SEGURO.sub("_", str(texto or "")).strip("._")
    return limpo[:80] or padrao


def pasta_da_execucao(base: str, projeto_codigo: str, arquivo: str, exec_id: str) -> str:
    stem = os.path.splitext(os.path.basename(arquivo))[0]
    return os.path.join(base, nome_seguro(projeto_codigo, "sem_projeto"), f"{nome_seguro(stem)}-{exec_id}")


def _json(caminho, objeto):
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(objeto, f, ensure_ascii=False, indent=2)


def _csv(caminho, colunas, linhas):
    with open(caminho, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(colunas)
        for l in linhas:
            w.writerow([("" if l.get(c) is None else l.get(c)) for c in colunas])


def gravar(pasta: str, *, tabelas: dict = None, totalizadora: list = None, orcamento: dict = None, relatorio: dict = None,
           pendencias: list = None) -> list:
    """Cria/atualiza os arquivos da execução. Devolve os nomes gravados."""
    os.makedirs(pasta, exist_ok=True)
    feitos = []
    if tabelas is not None:
        _csv(os.path.join(pasta, "cabos.csv"), ["entidade", "operacao", "ativo", "qtdAtivos"], tabelas.get("cabos", []))
        _csv(os.path.join(pasta, "outros.csv"), ["entidade", "operacao", "ativo"], tabelas.get("outros", []))
        _json(os.path.join(pasta, "tabelas.json"), {**tabelas, **({"totalizadora": totalizadora} if totalizadora is not None else {})})
        feitos += ["cabos.csv", "outros.csv", "tabelas.json"]
    if orcamento is not None:
        _json(os.path.join(pasta, "orcamento.json"), orcamento)
        _csv(os.path.join(pasta, "orcamento.csv"), ["operacao", "mdo", "codigo", "desc_codigo", "filtro", "total"], orcamento.get("resultado", []))
        feitos += ["orcamento.json", "orcamento.csv"]
    if pendencias is not None:
        _json(os.path.join(pasta, "pendencias.json"), pendencias)
        feitos.append("pendencias.json")
    if relatorio is not None:
        _json(os.path.join(pasta, "relatorio.json"), relatorio)
        feitos.append("relatorio.json")
    return feitos
