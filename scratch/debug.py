import pymupdf
import re
import json

def test():
    lista_path = "c:/Users/gabriel.sales/Desktop/PROJETOS VSCODE/PROJETO 2 LEITOR DE PDF ONLINE/LISTA 0022602688.pdf"
    doc_lista = pymupdf.open(lista_path)
    extracted = []
    
    for page in doc_lista:
        blocks = page.get_text("dict", flags=11)["blocks"]
        texts = []
        for b in blocks:
            if "lines" not in b: continue
            for l in b["lines"]:
                if "spans" not in l: continue
                for s in l["spans"]:
                    text = s["text"].strip()
                    if text:
                        texts.append(text)
        
        for i, text in enumerate(texts):
            if re.match(r'^\d{6}$', text):
                codigo = text
                requisitar = None
                devolver = None
                desc = ""
                
                print(f"Encontrou código: {codigo}")
                for j in range(i+1, min(i+15, len(texts))):
                    print(f"  texto[{j}]: '{texts[j]}'")
                    if re.match(r'^\d+(,\d+)?$', texts[j]):
                        val = float(texts[j].replace('.', '').replace(',', '.'))
                        if requisitar is None:
                            requisitar = val
                            print(f"    -> Requisitar = {val}")
                        elif devolver is None:
                            devolver = val
                            print(f"    -> Devolver = {val}")
                            if j + 1 < len(texts):
                                desc = texts[j+1]
                                print(f"    -> Descrição = {desc}")
                            break
                if devolver and devolver > 0:
                    extracted.append({
                        "operacao": "R",
                        "mdo": "MATERIAL",
                        "codigo": codigo,
                        "desc_codigo": desc,
                        "total": devolver
                    })
    print("\nRESULTADOS EXTRAIDOS:")
    print(json.dumps(extracted, indent=2))
test()
