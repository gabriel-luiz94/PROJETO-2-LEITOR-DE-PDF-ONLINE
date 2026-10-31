"""tests/oraculo/oraculo_anotacoes.py — roda o código PYTHON DA APLICAÇÃO (services/anotacao_cabos.py e services/vinculacao_regras.py)
sobre as anotações de cabo de casos_anotacoes.json; é o oráculo da paridade de src/Nucleo/Anotacoes.cs.
Uso (na raiz do repositório): python plugin_nanocad/tests/oraculo/oraculo_anotacoes.py > plugin_nanocad/tests/oraculo/esperado_anotacoes.json
"""
import json
import os
import sys

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
sys.path.insert(0, RAIZ)
from services.anotacao_cabos import prefixo_e_comprimento  # noqa: E402
from services.vinculacao_regras import derivar_cabo, descricao_cabo  # noqa: E402

casos = json.load(open(os.path.join(os.path.dirname(__file__), "casos_anotacoes.json"), encoding="utf-8"))
saida = []
for t in casos:
    pc = prefixo_e_comprimento(t)
    d = derivar_cabo(t)
    saida.append({"texto": t, "codigo": pc[0] if pc else None, "comprimento": pc[1] if pc else None,
                  "fases": d["fases"], "tensao": d["tensao"], "nivel": d["nivel"], "descricao": descricao_cabo(t)})
print(json.dumps(saida, ensure_ascii=False, indent=1))
