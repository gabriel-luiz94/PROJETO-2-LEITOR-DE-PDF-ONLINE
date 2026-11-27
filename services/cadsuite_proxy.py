"""Conversão local da representação visual de proxies em uma cópia DXF."""

import json
from pathlib import Path

TIPOS = {"ACAD_PROXY_ENTITY", "SPDSNOTEPOSITION"}


def converter_proxies(doc):
    from ezdxf.proxygraphic import ProxyGraphic

    modelo = doc.modelspace()
    convertidos = 0
    avisos = []
    for entidade in list(modelo):
        if entidade.dxftype() not in TIPOS:
            continue
        identificador = f"{entidade.dxftype()} {entidade.dxf.handle}"
        grafico = getattr(entidade, "proxy_graphic", None)
        adicionadas = []
        try:
            if not grafico:
                raise ValueError("sem representação gráfica")
            partes = list(ProxyGraphic(grafico, doc).virtual_entities())
            if not partes:
                raise ValueError("nenhuma geometria compatível")
            for parte in partes:
                modelo.add_entity(parte)
                adicionadas.append(parte)
            modelo.delete_entity(entidade)
        except Exception as erro:
            for parte in adicionadas:
                modelo.delete_entity(parte)
            avisos.append(f"{identificador}: {erro}")
        else:
            convertidos += 1
    return {"convertidos": convertidos, "avisos": avisos}


def main(argv):
    if len(argv) != 1:
        return 2
    pasta = Path(argv[0])
    relatorio = pasta / "resultado.json"
    try:
        import ezdxf

        ezdxf.options.load_proxy_graphics = True
        doc = ezdxf.readfile(pasta / "entrada.dxf")
        resultado = converter_proxies(doc)
        if resultado["convertidos"]:
            doc.saveas(pasta / "saida.dxf")
    except Exception as erro:
        relatorio.write_text(json.dumps({"erro": str(erro)}, ensure_ascii=False), encoding="utf-8")
        return 1
    relatorio.write_text(json.dumps(resultado, ensure_ascii=False), encoding="utf-8")
    return 0
