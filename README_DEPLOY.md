# Guia de Deploy Online (Render.com)

Siga estes passos para colocar seu app online gratuitamente:

1.  **Crie o Repositório no GitHub**:
    *   Vá em [github.com/new](https://github.com/new).
    *   Nomeie como `leitor-pdf-online`.
    *   Mantenha como **Public**.
    *   **Não** inicialize com README ou .gitignore (já criamos eles localmente).

2.  **Suba seu código**:
    Abra o terminal na pasta do projeto e rode:
    ```bash
    git init
    git add .
    git commit -m "Arquivos para deploy"
    git branch -M main
    git remote add origin https://github.com/gabriel-luiz94/leitor-pdf-online.git
    git push -u origin main
    ```

3.  **Configure o Render**:
    *   Crie uma conta em [Render.com](https://render.com/).
    *   Clique em **New +** > **Web Service**.
    *   Conecte sua conta do GitHub e selecione o repositório `leitor-pdf-online`.
    *   **Configurações**:
        *   **Runtime**: `Python 3`
        *   **Build Command**: `pip install -r requirements.txt`
        *   **Start Command**: `uvicorn app:app --host 0.0.0.0 --port $PORT`
    *   Aguarde o deploy terminar e pronto!
