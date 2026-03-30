
import PyPDF2 as p2
import pymupdf
import pandas
import tkinter
from tkinter import filedialog


""" 
pdf = open('002-23-00637_DM_453501_A1R.pdf', 'rb')

pdf_reader = p2.PdfReader(pdf)

n = len(pdf_reader.pages)

for i in range(0, n):
    print('Página {}'.format(i+1))
    page = pdf_reader.pages[i]
    if page.extract_text() != "":
        conteudo = page.extract_text()
    else:
        print("image")
    with open("teste_de_pdf.txt", 'a', encoding='utf-8') as arq:
        arq.write(conteudo) """


def flags_decomposer(flags):
    """Make font flags human readable."""
    l = []
    if flags & 2 ** 0:
        l.append("superscript")
    if flags & 2 ** 1:
        l.append("italic")
    if flags & 2 ** 2:
        l.append("serifed")
    else:
        l.append("sans")
    if flags & 2 ** 3:
        l.append("monospaced")
    else:
        l.append("proportional")
    if flags & 2 ** 4:
        l.append("bold")
    return ", ".join(l)


FILEOPENOPTIONS = dict(defaultextension=".pdf", initialdir="D://workspace",
                       filetypes=[('pdf file', '*.pdf')])

filename = filedialog.askopenfilename(**FILEOPENOPTIONS)

doc = pymupdf.open(filename)
# page = doc[0]
with open('teste_de_pdf.txt', 'w') as file:
    pass  # This will clear the file
# read page text as a dictionary, suppressing extra spaces in CJK fonts
pagina = 1
for page in doc:
    blocks = page.get_text("dict", flags=11)["blocks"]
    for b in blocks:  # iterate through the text blocks
        for l in b["lines"]:  # iterate through the text lines
            for s in l["spans"]:  # iterate through the text spans
                print("")
                font_properties = "Font: '%s' (%s), size %g, color #%06x" % (
                    s["font"],  # font name
                    flags_decomposer(s["flags"]),  # readable font flags
                    s["size"],  # font size
                    s["color"],  # font color
                )
                texto = " %s ; color #%06x; size %s; (%s)" % (
                    s["text"],
                    s["color"],
                    s["size"],
                    flags_decomposer(s["flags"]),

                )
                print("Text: '%s'" % s["text"])  # simple print of text
                print(font_properties)
                # print(s["color"])
                with open("teste_de_pdf.txt", 'a', encoding='utf-8') as arq:
                    arq.write("\n %s ;" % pagina)
                    # arq.write("Text: '%s'\n" % s["text"])
                    arq.write(texto)
    pagina = pagina + 1

