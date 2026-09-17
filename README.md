# Leitor de Projetos Pro

Sistema para leitura e análise de projetos de rede elétrica a partir de arquivos PDF e DXF, com orçamento automatizado, autenticação, sincronização offline-first com a nuvem e um assistente de IA integrado.

Funciona tanto como aplicativo desktop nativo (via `pywebview`, empacotável em `.exe`) quanto como serviço web hospedado na nuvem.

## 🚀 Tecnologias Utilizadas

* **FastAPI**: backend do servidor e das APIs.
* **PyMuPDF (fitz)**: extração de texto e metadados de arquivos PDF.
* **ezdxf**: extração de conteúdo de arquivos DXF.
* **SQLite**: banco local (modo offline / desktop).
* **Supabase**: banco na nuvem, autenticação e sincronização em tempo real (modo online).
* **JWT + bcrypt**: autenticação e hashing de senhas.
* **WebSockets**: atualização em tempo real da interface entre clientes.
* **Google Gemini / OpenAI (e compatíveis, ex: Ollama, OpenRouter)**: chat de IA com contexto do projeto carregado.
* **pywebview**: janela de aplicativo desktop nativo.

## 📂 Estrutura do Projeto

* `app.py`: ponto de entrada da aplicação FastAPI (modos `desktop` e `server`).
* `config.py`: configurações centralizadas e resolução de caminhos (dev e `.exe` via PyInstaller).
* `database.py`: inicialização do SQLite, criação de tabelas e seed automático.
* `models.py`: modelos Pydantic usados nas rotas.
* `routers/`: rotas da API (`auth`, `obras`, `orcamento`, `recs`, `projetos`, `regras`, `ai_chat`, `upload`, `admin`, `update`, `health`).
* `services/`: lógica de negócio (extração de PDF/DXF, cálculo de orçamento, sincronização offline-first, cliente Supabase, auto-update).
* `middleware/`: middleware de autenticação JWT.
* `static/`: interface HTML/CSS/JS servida pelo próprio FastAPI.
* `data/`: seed da tabela de orçamento (`tabela_seed.csv`).
* `scripts/`: scripts auxiliares (release, schema do Supabase).

## 🛠️ Como rodar localmente

1. Instale o Python 3.
2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Copie `.env.example` para `.env` e preencha as variáveis (Supabase, JWT, admin inicial, chaves de IA opcionais).
4. Execute:
   ```bash
   python app.py
   ```
5. Em modo desktop, o app abre em uma janela nativa (ou no navegador em `http://127.0.0.1:8000` como fallback).

No Windows também é possível usar `Iniciar_Leitor_PDF.bat` ou criar um atalho na área de trabalho com `Criar_Atalho_Desktop.vbs`.

## ⚙️ Modos de operação

Controlado pela variável `APP_MODE`:

* `desktop` (padrão): uso local/`.exe`, CORS aberto, permite extração de arquivos locais via `/extract-local`.
* `server`: deploy na nuvem, CORS restrito às origens definidas em `CORS_ORIGINS`, sem acesso a arquivos locais.

## 🌐 Deploy Automático

Configurado via GitHub Actions:

* **Backend** → [Fly.io](https://fly.io) (`.github/workflows/deploy-backend.yml`), disparado em mudanças em `app.py`, `requirements.txt`, `Dockerfile`, `fly.toml` ou `static/`.
* **Frontend** → [Cloudflare Pages](https://pages.cloudflare.com) (`.github/workflows/deploy-frontend.yml`), disparado em mudanças em `static/`.
