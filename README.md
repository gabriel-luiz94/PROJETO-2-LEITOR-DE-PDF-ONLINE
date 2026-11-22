# OrçaCore

> Desenvolvido por **Gabriel Sales** · (antes: "Leitor de Projetos Pro", "ReadGrid", "GridReader" e "GridR")

Sistema para leitura e análise de projetos de rede elétrica a partir de arquivos PDF e DXF, com orçamento automatizado, autenticação, sincronização offline-first com a nuvem e um assistente de IA integrado.

Funciona tanto como aplicativo desktop nativo (via `pywebview`, empacotável em `.exe`) quanto como serviço web hospedado na nuvem.

## Instalação no Windows

1. Baixe `OrcaCore-Setup.exe` nos arquivos da [release mais recente](https://github.com/gabriel-luiz94/PROJETO-2-LEITOR-DE-PDF-ONLINE/releases/latest).
2. Execute o instalador na sua conta. O programa fica em `%LOCALAPPDATA%\Programs\OrcaCore` e aparece no menu Iniciar e em **Configurações → Aplicativos → Aplicativos instalados**. A instalação e a desinstalação usam as permissões da conta atual.
3. Para atualizar, abra o OrçaCore e, quando uma nova release estiver disponível, use **Instalar e Reiniciar**.

O arquivo `OrcaCore.exe` continua disponível para uso portátil. Se você já usa uma cópia portátil com dados locais, feche o OrçaCore antes de copiar o banco `banco_resumo.db`, a pasta `skills` e o `.env` (se existirem) para a nova pasta. Faça essa migração antes de começar a trabalhar pela instalação e preserve uma cópia de segurança se a pasta de destino já tiver dados. Guarde os arquivos de obra que você salvou nos locais escolhidos.

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
* `plugin_nanocad/`: subproduto — plugin C# para o nanoCAD que lê o desenho e gera o `.obra.json` (DLL pronta para o APPLOAD em `plugin_nanocad/appload/`; documentação própria em `plugin_nanocad/.ai/`).

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

## 🤖 Chave de IA padrão

O chat de IA funciona sem o usuário digitar chave, usando `GEMINI_API_KEY` (ou `GOOGLE_API_KEY`):

* **Servidor:** defina como variável de ambiente/secret do deploy.
* **Desktop (.exe):** crie um arquivo `.env` na **mesma pasta do executável** com `GEMINI_API_KEY=...`.

Precedência: chave digitada pelo usuário → chave salva no servidor → chave padrão. Com a chave padrão há um limite de mensagens por usuário por minuto (`AI_RATE_LIMIT_POR_MINUTO`, padrão 20). `GET /api/health` mostra a origem ativa em `ai_key_source`, nunca o valor.

Para configurar DeepSeek como reserva, entre no **Resumo → Configurações → Inteligência artificial → Chave API** com perfil administrador. Em **Provedor reserva**, marque **Usar provedor reserva**, selecione **DeepSeek**, confira o modelo `deepseek-flash` e cole a chave do DeepSeek. Clique em **Salvar reserva** e depois em **Testar**. O teste consulta a tabela de Cabos vazia e confirma que o modelo consegue responder após receber o resultado da ferramenta. [Chaves da API do DeepSeek](https://platform.deepseek.com/api_keys)

`OPENAI_API_KEY` é usada exclusivamente no endereço oficial da OpenAI. Chaves pessoais ficam na configuração do navegador. Ao trocar o endereço ou provedor do reserva, informe a chave correspondente. Endereços adicionais podem ser configurados em `AI_OPENAI_ALLOWED_BASE_URLS`, separados por vírgula.

## ⚙️ Modos de operação

Controlado pela variável `APP_MODE`:

* `desktop` (padrão): uso local/`.exe`, CORS aberto, permite extração de arquivos locais via `/extract-local`.
* `server`: deploy na nuvem, CORS restrito às origens definidas em `CORS_ORIGINS`, sem acesso a arquivos locais.

Em servidor ou com `HEADLESS=1`, defina `JWT_SECRET` com um valor aleatório de pelo menos 32 bytes antes de iniciar. Gere um valor com `python -c "import secrets; print(secrets.token_hex(32))"` e configure-o como segredo do ambiente. No desktop, uma instalação sem segredo válido gera e mantém uma chave no perfil local. Sessões assinadas pela antiga chave padrão exigirão novo login.

## 🌐 Deploy Automático

Configurado via GitHub Actions:

* **Backend** → [Fly.io](https://fly.io) (`.github/workflows/deploy-backend.yml`), disparado em mudanças em `app.py`, `requirements.txt`, `Dockerfile`, `fly.toml` ou `static/`.
* **Frontend** → [Cloudflare Pages](https://pages.cloudflare.com) (`.github/workflows/deploy-frontend.yml`), disparado em mudanças em `static/`.
