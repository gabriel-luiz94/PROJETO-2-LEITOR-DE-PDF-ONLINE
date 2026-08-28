import pymupdf
import sys
import json

def extract(pdf_path):
    print(f"--- Extracting {pdf_path} ---")
    doc = pymupdf.open(pdf_path)
    for i in range(min(2, len(doc))):
        page = doc[i]
        text = page.get_text("text")
        print(f"Page {i+1}:")
        print(text[:1500]) # Print first 1500 chars to see structure
        print("-" * 40)
    doc.close()

if __name__ == "__main__":
    extract("c:/Users/gabriel.sales/Desktop/PROJETOS VSCODE/PROJETO 2 LEITOR DE PDF ONLINE/ORÇAMENTO 0022602688.pdf")
    extract("c:/Users/gabriel.sales/Desktop/PROJETOS VSCODE/PROJETO 2 LEITOR DE PDF ONLINE/LISTA 0022602688.pdf")
