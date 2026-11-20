"""Compara as unidades de estrutura exportadas pelo núcleo com o recálculo da aplicação."""
import json
import sys
from pathlib import Path

raiz = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(raiz))
from services.autonomo.montagem import prefixo_cabo
from services.autonomo.vinculacao import codigo_estrutura, codigos_estrutura
from services.vinculacao_regras import vincular_por_poste


def conferir():
    arquivo = raiz / "plugin_nanocad/exemplos/Projeto_Rua_das_Flores.obra.json"
    dados = json.loads(arquivo.read_text(encoding="utf-8"))["dados"]
    cfg = json.loads((raiz / "data/vinculacao_pernambuco_seed.json").read_text(encoding="utf-8"))
    resultado = vincular_por_poste(dados["cabos"], dados["outros"], cfg, prefixo_cabo, codigo_estrutura, codigos_estrutura)
    assert resultado["vinculos"], "O exemplo precisa conter cabos com geometria"
    for linha in resultado["vinculos"]:
        cabo = dados["cabos"][linha["cabo"]]
        assert sorted(cabo.get("vinculoEstruturas", [])) == sorted(linha["estruturas"]), (cabo, linha)
    print(len(resultado["vinculos"]), "cabos comparados com a aplicação")


if __name__ == "__main__":
    conferir()
