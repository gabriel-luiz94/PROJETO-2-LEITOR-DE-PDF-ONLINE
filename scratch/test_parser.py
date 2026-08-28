import pymupdf
import re
import json

def process_pdfs(orcamento_path, lista_path):
    extracted_items = []
    
    # Processa ORÇAMENTO
    print("Processando Orçamento...")
    doc_orc = pymupdf.open(orcamento_path)
    current_mdo = "MÃO-DE-OBRA" # Default
    
    for page in doc_orc:
        # Extrai os textos na ordem em que aparecem
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
        
        # Analisando a sequência de textos
        for i, text in enumerate(texts):
            if text == "MAO-DE-OBRA":
                current_mdo = "MÃO-DE-OBRA"
            elif text == "MATERIAL":
                current_mdo = "MATERIAL"
            
            # Identifica código de 6 dígitos numéricos
            if re.match(r'^\d{6}$', text):
                codigo = text
                # A descrição normalmente vem antes (i-1)
                desc = texts[i-1] if i > 0 else ""
                
                # A quantidade pode estar em i-2 ou i-3 dependendo se há outros elementos.
                qtd = 1.0
                for j in range(i-1, max(-1, i-5), -1):
                    if re.match(r'^\d+,\d+$', texts[j]):
                        try:
                            qtd = float(texts[j].replace('.', '').replace(',', '.'))
                            # The first number found going backwards is the quantity.
                            # Oh wait, the previous number could be 'V.Total' in some cases?
                            # Let's look at the pattern: '232,62' (Total), '232,62' (Unit), '116,31', '139,83', '2,000' (Qtd).
                            # Actually, Qtd is always the *last* number before the description.
                            # So it's exactly i-2. But let's check if i-2 is really a number.
                            break
                        except ValueError:
                            pass
                
                # A operação vem logo depois (i+1)
                op = texts[i+1] if i + 1 < len(texts) else "I"
                if op not in ['I', 'R']:
                    op = 'I'
                
                extracted_items.append({
                    "operacao": op,
                    "mdo": current_mdo,
                    "codigo": codigo,
                    "desc_codigo": desc,
                    "total": qtd,
                    "source": "orcamento"
                })

    doc_orc.close()

    # Processa LISTA
    print("Processando Lista...")
    doc_lista = pymupdf.open(lista_path)
    
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
                # Em LISTA, a estrutura é:
                # Codigo
                # UN
                # Requisitar (qtd)
                # Devolver (qtd)
                # Descrição
                
                if i + 4 < len(texts):
                    devolver_str = texts[i+3]
                    try:
                        devolver_qtd = float(devolver_str.replace('.', '').replace(',', '.'))
                        if devolver_qtd > 0:
                            desc = texts[i+4]
                            extracted_items.append({
                                "operacao": "R",
                                "mdo": "MATERIAL",
                                "codigo": codigo,
                                "desc_codigo": desc,
                                "total": devolver_qtd,
                                "source": "lista"
                            })
                    except ValueError:
                        pass
    doc_lista.close()
    
    return extracted_items

if __name__ == "__main__":
    orc_path = "c:/Users/gabriel.sales/Desktop/PROJETOS VSCODE/PROJETO 2 LEITOR DE PDF ONLINE/ORÇAMENTO 0022602688.pdf"
    lst_path = "c:/Users/gabriel.sales/Desktop/PROJETOS VSCODE/PROJETO 2 LEITOR DE PDF ONLINE/LISTA 0022602688.pdf"
    res = process_pdfs(orc_path, lst_path)
    print(json.dumps(res, indent=2))
