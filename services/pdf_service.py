"""
services/pdf_service.py — Funções de extração de conteúdo de arquivos PDF.
"""


def color_to_hex(c):
    """Converte valor de cor do PyMuPDF para hex string."""
    if c is None:
        return "#000000"
    try:
        if isinstance(c, (list, tuple)):
            if len(c) >= 3:
                rgb = [max(0, min(255, int(x * 255))) for x in c[:3]]
                return "#%02x%02x%02x" % tuple(rgb)
        val = int(c)
        return "#%06x" % (val & 0xFFFFFF)
    except Exception:
        return "#000000"


def flags_decomposer(flags):
    """Converte flags de fonte do PyMuPDF em descrição legível."""
    l = []
    if flags & 1:
        l.append("superscript")
    if flags & 2:
        l.append("italic")
    if flags & 4:
        l.append("serifed")
    else:
        l.append("sans")
    if flags & 8:
        l.append("monospaced")
    else:
        l.append("proportional")
    if flags & 16:
        l.append("bold")
    return ", ".join(l)


def extract_pdf_content(doc) -> list[dict]:
    """Extrai todo o conteúdo textual de um documento PDF (pymupdf)."""
    extracted = []
    for i, page in enumerate(doc):
        blocks = page.get_text("dict", flags=11)["blocks"]
        for b in blocks:
            if "lines" not in b:
                continue
            for l in b["lines"]:
                if "spans" not in l:
                    continue
                for s in l["spans"]:
                    text = s["text"].strip()
                    if text:
                        extracted.append({
                            "pagina": i + 1,
                            "texto": text,
                            "fonte": s["font"],
                            "tamanho": round(s["size"], 2),
                            "cor": color_to_hex(s.get("color", 0)),
                            "flags": flags_decomposer(s["flags"])
                        })
    return extracted
