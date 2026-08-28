import requests

url = "http://localhost:8000/api/importar-rec-pdf"
files = {
    'orcamento': open('c:/Users/gabriel.sales/Desktop/PROJETOS VSCODE/PROJETO 2 LEITOR DE PDF ONLINE/ORÇAMENTO 0022602688.pdf', 'rb'),
    'lista': open('c:/Users/gabriel.sales/Desktop/PROJETOS VSCODE/PROJETO 2 LEITOR DE PDF ONLINE/LISTA 0022602688.pdf', 'rb')
}
response = requests.post(url, files=files)
print(response.status_code)
print(response.json())
