# Leitor de PDF Online Pro

Uma aplicação web moderna para extração e análise de dados de arquivos PDF em tempo real.

## 🚀 Tecnologias Utilizadas

*   **FastAPI**: Backend rápido para processamento de texto.
*   **PyMuPDF (fitz)**: Biblioteca robusta para extração de dados do PDF.
*   **Aesthetics**: Interface premium com micro-animações, suporte a filtros dinâmicos e exportação de dados.
*   **WebSockets**: Sincronização em tempo real (para uso local/avançado).

## 📂 Estrutura do Projeto

*   `app.py`: O servidor FastAPI que processa o PDF.
*   `static/`: Contém a interface HTML/CSS/JS.
*   `requirements.txt`: Dependências do sistema.

## 🛠️ Como rodar localmente

1. Instale o Python.
2. No terminal, execute:
   ```bash
   pip install -r requirements.txt
   python app.py
   ```
3. O app abrirá automaticamente no navegador em `http://127.0.0.1:8000`.

## 🌐 Deploy Automático

Este repositório está configurado para deploy automático no **Render.com** ou **Railway.app**. Siga o [README_DEPLOY.md](README_DEPLOY.md) (opcional) para instruções detalhadas.
