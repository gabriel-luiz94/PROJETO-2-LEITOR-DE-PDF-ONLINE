"""Resumo estável de um PDF (palavras + desenhos por página) para comparar saídas sem depender de datas/IDs internos do arquivo."""
import hashlib
import json

import fitz


def resumo(pdf_bytes: bytes) -> dict:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    paginas = []
    for p in doc:
        palavras = [[round(w[0], 2), round(w[1], 2), round(w[2], 2), round(w[3], 2), w[4]] for w in p.get_text("words")]
        desenhos = [[d["type"], [round(x, 2) for x in (d["rect"].x0, d["rect"].y0, d["rect"].x1, d["rect"].y1)], d.get("fill"), d.get("color")] for d in p.get_drawings()]
        paginas.append({"palavras": hashlib.sha1(json.dumps(palavras, ensure_ascii=False).encode()).hexdigest(),
                        "desenhos": hashlib.sha1(json.dumps(desenhos, ensure_ascii=False, default=str).encode()).hexdigest(), "n_palavras": len(palavras), "n_desenhos": len(desenhos)})
    doc.close()
    return {"paginas": paginas}
