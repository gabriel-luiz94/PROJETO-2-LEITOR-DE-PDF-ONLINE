"""tests/oraculo/dxf_para_bruto.py — DXF → "desenho bruto" do plugin (mesmo formato que o ORC_DIAGNOSTICO grava no nanoCAD), para
testar o núcleo C# com desenhos reais/sintéticos. Com --app também grava o que o LEITOR DA APLICAÇÃO faz com o mesmo DXF
(services/dxf_service.extract_dxf_content + motor JS static/regras_leitor_engine.js), que é o oráculo da paridade.
Uso (na raiz do repositório): python plugin_nanocad/tests/oraculo/dxf_para_bruto.py desenho.dxf bruto.json [--app esperado.json --regras 081]
"""
import json
import os
import subprocess
import sys
import tempfile

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
sys.path.insert(0, RAIZ)
import ezdxf  # noqa: E402


def cor(e):
    if e.dxf.hasattr("true_color"):
        return {"m": "rgb", "rgb": e.dxf.true_color}
    aci = e.dxf.get("color", 256)
    if aci == 256:
        return {"m": "layer"}
    if aci == 0:
        return {"m": "block"}
    return {"m": "aci", "aci": aci}


def xy(v):
    return [round(float(v[0]), 4), round(float(v[1]), 4)]


def base(e, tipo):
    return {"t": tipo, "h": e.dxf.handle, "layer": e.dxf.get("layer", "0"), "cor": cor(e)}


def bruta(e, com_definicao=True):
    tp = e.dxftype()
    if tp == "TEXT" or tp == "ATTRIB":
        b = base(e, tp)
        b.update(texto=e.dxf.text, x=round(e.dxf.insert.x, 4), y=round(e.dxf.insert.y, 4), altura=e.dxf.get("height", 0))
        if tp == "ATTRIB" and e.is_invisible:
            b["invisivel"] = True
        return b
    if tp == "MTEXT":
        b = base(e, "MTEXT")
        b.update(texto=e.text, x=round(e.dxf.insert.x, 4), y=round(e.dxf.insert.y, 4), altura=e.dxf.get("char_height", 0))
        return b
    if tp == "MULTILEADER":
        b = base(e, "MLEADER")
        ctx = e.context
        if ctx.mtext is not None:
            b.update(texto=ctx.mtext.default_content, x=round(ctx.mtext.insert.x, 4), y=round(ctx.mtext.insert.y, 4), altura=ctx.char_height)
        linhas = []
        for ld in ctx.leaders:
            for ln in ld.lines:
                pts = [xy(v) for v in ln.vertices] + [xy(ld.last_leader_point)]
                if len(pts) >= 2:
                    linhas.append(pts)
        if linhas:
            b["linhas"] = linhas
        return b
    if tp == "INSERT":
        b = base(e, "INSERT")
        b.update(bloco=e.dxf.name, x=round(e.dxf.insert.x, 4), y=round(e.dxf.insert.y, 4))
        attrs = [bruta(a, False) for a in e.attribs]
        if attrs:
            b["atributos"] = attrs
        if com_definicao:
            blk = e.doc.blocks.get(e.dxf.name)
            defs = [bruta(x, False) for x in blk.query("TEXT MTEXT")] if blk else []
            if defs:
                b["definicao"] = defs
        return b
    if tp == "CIRCLE":
        b = base(e, "CIRCLE")
        b.update(x=round(e.dxf.center.x, 4), y=round(e.dxf.center.y, 4), raio=e.dxf.radius)
        return b
    pts = []
    if tp == "LINE":
        pts = [xy(e.dxf.start), xy(e.dxf.end)]
    elif tp == "LWPOLYLINE":
        pts = [xy(v) for v in e.get_points("xy")]
        if e.closed and pts:
            pts.append(pts[0])
    elif tp == "POLYLINE":
        pts = [xy(v.dxf.location) for v in e.vertices]
        if e.is_closed and pts:
            pts.append(pts[0])
    elif tp == "LEADER":
        pts = [xy(v) for v in e.vertices]
    else:
        return None
    b = base(e, tp)
    b["pontos"] = pts
    return b


def converter(caminho):
    doc = ezdxf.readfile(caminho)
    layers = {}
    for l in doc.layers:
        layers[l.dxf.name] = {"m": "rgb", "rgb": l.dxf.true_color} if l.dxf.hasattr("true_color") else {"m": "aci", "aci": l.dxf.color}
    ents = [b for b in (bruta(e) for e in doc.modelspace()) if b is not None]
    return doc, {"formato": "plugin-desenho-bruto", "versao": 1, "arquivo": os.path.basename(caminho), "layers": layers, "entidades": ents, "avisos": []}


def leitor_da_aplicacao(doc, regras):
    """O que a tela Leitor da aplicação faz com este DXF: extração (Python) + motor de regras (JS)."""
    from services.dxf_service import extract_dxf_content
    linhas = extract_dxf_content(doc)
    itens = [{"texto": r["texto"], "cor": r["cor"], "layer": r["layer"], "pagina": r.get("pagina", 1), "_x": r["_x"], "_y": r["_y"]} for r in linhas]
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
        json.dump([{"nome": "dxf", "regras": regras, "itens": itens}], f, ensure_ascii=False)
        tmp = f.name
    try:
        saida = subprocess.run(["node", os.path.join(os.path.dirname(__file__), "oraculo_motor.js"), tmp], capture_output=True, text=True, check=True).stdout
    finally:
        os.unlink(tmp)
    res = json.loads(saida)[0]["resultados"]
    return [{"texto": i["texto"], "cor": i["cor"], "layer": i["layer"], **r} for i, r in zip(itens, res)]


if __name__ == "__main__":
    doc, bruto = converter(sys.argv[1])
    with open(sys.argv[2], "w", encoding="utf-8") as f:
        json.dump(bruto, f, ensure_ascii=False, indent=1)
    if "--app" in sys.argv:
        regras = sys.argv[sys.argv.index("--regras") + 1] if "--regras" in sys.argv else "081"
        with open(sys.argv[sys.argv.index("--app") + 1], "w", encoding="utf-8") as f:
            json.dump(leitor_da_aplicacao(doc, regras), f, ensure_ascii=False, indent=1)
