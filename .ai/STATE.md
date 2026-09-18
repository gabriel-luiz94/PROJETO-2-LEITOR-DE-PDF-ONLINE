# ESTADO DO PROJETO

> Retrato do estado atual, levantado por leitura do código. Itens listados como problema foram
> **verificados no código**, não inferidos. Nenhuma tarefa futura foi inventada: a seção
> "Próximas tarefas conhecidas" contém apenas o que está marcado como pendente no próprio projeto.

---

## Versão

`2.0.0` — `config.py:APP_VERSION`

Branch de trabalho: `claude/beautiful-pasteur-2tdk18`
Último commit analisado: `530d5cc`

---

## Implementado

### Extração
- [x] Extração de PDF com texto, página, fonte, tamanho, **cor** e flags (`pdf_service.py`)
- [x] Extração de DXF com TEXT/MTEXT/ATTRIB e entidades dentro de blocos INSERT (`dxf_service.py`)
- [x] Resolução de cor do DXF por hierarquia (true_color → ACI → layer `TXT_RRGGBB` → layer → preto)
- [x] Quebra de MTEXT por parágrafo com herança de cor entre parágrafos e cor local em `{ }`
- [x] Ordenação espacial da saída do DXF (Y decrescente, depois X)
- [x] Importação de REC a partir de PDFs de Orçamento + Lista (`POST /api/importar-rec-pdf`)
- [x] Leitura de arquivo local por caminho, travada para executável (`/extract-local`)
- [x] Trigger de arquivo via WebSocket (`/trigger-file` + `trigger.py`)

### Classificação e regras
- [x] Classificação automática de entidade/operação/ativo por cor, layer e texto (`computeRowLogic`)
- [x] Normalização de texto livre para fórmula de ativo (`processAtivoFormula`)
- [x] Reclassificação automática ao sair do campo Ativo quando a entidade está em `0`
- [x] Cálculo de `qtdAtivos` com herança de fase para linhas standalone
- [x] Camada unificada com `baseId` e `origem` (`syncTotalizadora`)
- [x] Motor de regras de conversão (ADIÇÃO/SUBST, fator, arredondamento, val_min/val_max)
- [x] Regras de conversão por projeto, com persistência em nuvem + `localStorage`
- [x] Import/export das regras em CSV

### Cálculo
- [x] Cascata de resolução de ativo em 5 passos com filtro de origem
- [x] Expansão ativo → componente → todos os códigos
- [x] Prioridade de projeto (exato > genérico > fallback), com suporte a `PROJETO_A/PROJETO_B`
- [x] Soma `I`/`R` por `(codigo, mdo)` e ordenação MDO → operação → descrição
- [x] Lista de ativos não encontrados devolvida ao frontend e destacada na UI

### Dados e sincronização
- [x] SQLite com criação idempotente de tabelas e migrações por `ALTER TABLE`
- [x] Seed automático da base técnica a partir de `data/tabela_seed.csv` (2.328 linhas)
- [x] Precedência da base: edições do usuário > master do Supabase > seed
- [x] Download paginado da tabela master do Supabase (`sync_tabela_master`)
- [x] Fallback silencioso para SQLite em obras, RECs e projetos quando o Supabase falha
- [x] Preservação de REC de outro usuário via cópia nomeada
- [x] Backup em JSON (`/api/backup/export`) e diagnóstico (`/api/health`), incluindo desde
      2026-09-18 checagem de configuração e alcançabilidade do Supabase (`supabase_configured`,
      `supabase_reachable`, `usuarios_nuvem_has_rows`) — ver TASK-001

### Autenticação e administração
- [x] Login com JWT (24 h) + refresh token (30 d), bcrypt e fallback offline
- [x] Migração automática de senhas SHA-256 legadas para bcrypt
- [x] Middleware global protegendo `/api/*` com whitelist de rotas públicas
- [x] Painel admin: CRUD de usuários, troca de role, redefinição de senha
- [x] Gestão da tabela master: linha avulsa, upload de CSV, sobrescrita completa
- [x] Audit log das ações administrativas

### IA
- [x] Chat com Gemini e OpenAI/compatíveis, com seleção de modelo
- [x] Prompt de domínio de redes elétricas (`prompt_rede_eletrica.txt`)
- [x] Interpretação da resposta como tabela `COMANDO/ID/AÇÃO/ATIVOS` aplicada às tabelas
- [x] Ações de UI por JSON (`ordenar`, `filtrar`, `limpar_filtros`)
- [x] Regras de aprendizado persistidas (`tabela regras`)

### Interface
- [x] Tabelas Cabos e Outros com undo/redo (80 níveis), autocomplete e filtros estilo Excel
- [x] Seleção de células, cópia e colagem em bloco
- [x] Modais geradores: Cabos, Postes e Estruturas, Ramais, Conexões
- [x] Tabela Totalizadora editável com destaque de ativo não encontrado
- [x] Tela de resultado com consolidação, edição manual, salvamento de REC e exportação

### Empacotamento e deploy
- [x] Modo desktop com janela nativa (pywebview) e fallback para navegador
- [x] Modo servidor headless
- [x] Build PyInstaller (`build.bat`)
- [x] Serviço de auto-update para desktop (`auto_updater.py` + `/api/update/*`)
- [x] Workflows do GitHub Actions para Fly.io e Cloudflare Pages

---

## Em desenvolvimento / escrito mas não conectado

- [ ] **Fila de operações offline** — `services/offline_queue.py` está completo (tabela `sync_log`,
      replay a cada 30 s, tratamento de chave duplicada), mas `start_offline_queue_worker()` e
      `enqueue_operation()` **não são chamados em lugar nenhum**. A tabela `sync_log` nunca é
      alimentada.
- [ ] **Sincronização em tempo real** — `services/realtime_sync.py` está escrito
      (listener do Supabase Realtime → SQLite → WebSocket), mas `start_realtime_sync()`
      **nunca é chamado**.
- [ ] **Monitor de conectividade** — `services/connectivity_monitor.py` está escrito, mas
      `start_connectivity_monitor()` **nunca é chamado**. O indicador de status online/offline no
      frontend não recebe eventos.
- [ ] **Embeddings das regras de aprendizado** — a coluna `regras.embedding` existe e é criada por
      migração, mas é sempre gravada como `NULL` (`regras.py:28`). Não há busca semântica.
- [ ] **Botão "LINHA VIVA"** — presente em `resumo.html:592` com `title="Função futura"`,
      sem handler.
- [ ] **`scripts/release.py`** — o `push_to_cloud()` está com o corpo comentado; só imprime
      mensagens. A publicação da versão na nuvem não acontece de fato.

---

## Próximas tarefas conhecidas

Apenas o que está explicitamente marcado como pendente no próprio projeto:

- [ ] Conectar os três workers de background acima (ou decidir removê-los)
- [ ] Implementar o botão "LINHA VIVA" (marcado como "Função futura" na UI)
- [ ] Completar `scripts/release.py:push_to_cloud` (código de referência já está comentado no arquivo)
- [ ] Adicionar `Dockerfile` e `fly.toml`, exigidos pelo workflow de deploy do backend

Nenhuma outra tarefa futura foi inferida. O que o usuário quiser fazer além disso deve virar um
arquivo em `.ai/tasks/`.

---

## Problemas conhecidos

> Verificados no código. **Nenhum foi corrigido** — a inicialização IA-First não alterou código.

### Deploy
1. **O workflow de backend não pode funcionar como está.**
   `.github/workflows/deploy-backend.yml` roda `flyctl deploy --remote-only`, mas **não existem
   `Dockerfile` nem `fly.toml`** no repositório. O workflow é disparado por mudanças em `app.py`,
   `requirements.txt`, `Dockerfile`, `fly.toml` e `static/`.
2. **`static/**` dispara os dois workflows** (backend e frontend) simultaneamente.
3. `build.bat` tem o caminho absoluto do PyInstaller da máquina de um desenvolvedor específico
   (`C:\Users\gabriel.sales\...`) — não funciona em outra máquina.
4. `scripts/release.py` invoca `pyinstaller --clean Leitor_PDF_Pro_v35.spec`, um arquivo `.spec`
   que não existe no repositório (`*.spec` está no `.gitignore`), e que também não corresponde ao
   nome usado no `build.bat` (`Leitor_Projetos`).

### Divergências de schema (SQLite ↔ Supabase)
5. `tabela_orcamento_master` **não tem a coluna `origem`** em `scripts/schema_supabase.sql`, mas
   `admin.py` (`add-row`, `upload-master`) a envia e `sync_service.py:57` a lê. Em um Supabase criado
   a partir do schema versionado, essas escritas falham ou perdem o campo.
   ✅ **Corrigido — TASK-002 (2026-09-18):** confirmado num caso real (mesmo padrão da TASK-001):
   o `INSERT` do Supabase em `upload_master_csv` falhava com o erro engolido em silêncio
   (`except Exception: logger.warning(...)`, resposta continuava "ok"), e a próxima chamada a
   `GET /api/orcamento/dados` sobrescrevia o local com a master antiga da nuvem, apagando o
   `origem` que tinha acabado de ser importado. `scripts/schema_supabase.sql` ganhou a coluna na
   `CREATE TABLE` e uma migração idempotente para instalações existentes. Além disso,
   `admin.py:sync_master_all` (o botão "Sincronizar Tudo com a Nuvem") **nunca lia nem gravava
   `origem`** — bug de código independente do schema, também corrigido (agora segue o mesmo
   padrão de `upload_master_csv`/`add_master_row`). E `upload_master_csv` passou a repassar ao
   admin, via `HTTPException` 500, quando a sincronização com o Supabase falha, em vez de mascarar
   como sucesso — mudança espelhada no frontend (`orcamento.html`) para mostrar a mensagem real.
   Achados relacionados, fora do escopo desta tarefa: o botão "Importar CSV Local" chama
   `/api/upload/csv-orcamento`, endpoint que **não existe** no backend (404); e a tabela pessoal
   do usuário (`routers/orcamento.py` `/upload` e `/salvar`) também não trata `origem`, mas nunca
   sincroniza com a nuvem.
   ✅ **Segunda causa raiz encontrada e corrigida (2026-09-18):** usuário rodou a migração no
   Supabase real e reimportou a base, mas `origem` continuava `null`, sem nenhum erro. Achado um
   commit anterior a esta tarefa (`08cd044`, feito manualmente pelo usuário em 17/09, quando a
   coluna `origem` de fato ainda não existia no Supabase) que **filtrava `origem` para fora do
   payload** antes de todo `INSERT` no Supabase, em `upload_master_csv` e `sync_master_all`
   (`rows_supabase = [{k: v for k, v in r.items() if k != 'origem'} for r in rows_to_insert]`).
   Esse trecho sobreviveu a um merge de `main` para a branch de trabalho e continuou removendo
   `origem` mesmo depois do schema já ter a coluna — por isso o insert nunca falhava (sem erro
   visível) e `origem` nunca chegava na nuvem. Removido nos dois handlers, mesclado em `main`
   (PR #3, commit `36de3aa`) e **confirmado funcionando pelo usuário em produção** após reimportar
   a base. Problema encerrado. Ver `.ai/tasks/TASK-002-18-09-2026.md` (status: CONCLUÍDA).
6. `usuarios_nuvem` **não tem a coluna `is_admin`** no schema versionado, mas `auth.py:57` e
   `admin.py:102,106` leem e escrevem `is_admin`.
   ✅ **Confirmado e corrigido — TASK-001 (2026-09-18):** o inverso também ocorre e foi verificado
   num caso real: o Supabase de um usuário tinha `is_admin` mas **não tinha `role`**, quebrando
   `POST /api/admin/users` (insere `role`) e fazendo `PUT /api/admin/users/{id}/role` falhar em
   silêncio — e, por consequência, o login também falhava para credenciais válidas.
   `scripts/schema_supabase.sql` ganhou uma migração idempotente
   (`ALTER TABLE ... ADD COLUMN IF NOT EXISTS role`). O usuário rodou a migração no Supabase real
   e confirmou: `GET /api/health` retornou `supabase_configured: true`, `supabase_reachable: true`,
   `usuarios_nuvem_has_rows: true`, e o **login voltou a funcionar**. Cadastro de novo usuário
   ainda não testado na prática — ver `.ai/tasks/TASK-001-18-09-2026.md`.
7. `admin.py:sync_master_all` grava a master **sem** a coluna `origem`, enquanto `upload_master_csv`
   grava **com**. Os dois caminhos produzem resultados diferentes.

### Papéis (roles) inconsistentes
8. Há **três conjuntos divergentes** de roles no código:
   - `database.py` cria colunas com `DEFAULT 'viewer'`;
   - `admin.py:20` define `VALID_ROLES = {"admin", "operador"}` e o docstring do módulo diz
     "três roles: admin, editor, viewer";
   - `auth_middleware.py` usa `"operador"` como fallback e `auth.py:58` usa
     `"admin" if is_admin else "operador"`.
   Um usuário criado pela migração com role `viewer` não casa com nenhuma verificação de permissão.
   ⚠️ **Manifestação real observada — TASK-001 (2026-09-18):** a migração que adicionou a coluna
   `role` a `usuarios_nuvem` (problema 6) fez o Postgres preencher **todas** as linhas existentes
   com o default `'operador'`, inclusive contas com `is_admin = true`. Como `auth.py:58` prioriza
   `role` sobre `is_admin`, isso derrubou o acesso admin de uma conta real até uma correção manual
   via SQL (`UPDATE ... SET role = 'admin' WHERE is_admin = true AND role <> 'admin'`). Qualquer
   `ALTER TABLE ... ADD COLUMN ... DEFAULT` futuro que crie uma coluna já lida em conjunto com
   outra (aqui, `role` vs. `is_admin`) tem esse mesmo risco de backfill — differenciar "coluna
   nova, valor desconhecido" de "coluna nova, valor default real" não é possível só com `DEFAULT`.
   Ver `.ai/tasks/TASK-001-18-09-2026.md`.
9. `admin.py:toggle_admin` (endpoint de compatibilidade `PUT /api/admin/users/{id}/admin`)
   **ignora o corpo da requisição e sempre define `role="admin"`**, mesmo quando a intenção seria
   remover o privilégio. Também executa uma consulta cujo resultado é descartado.

### Autenticação
10. `JWT_SECRET` é lido **duas vezes de forma independente**: `config.py:87` e
    `auth_middleware.py:16`. Ambos usam o mesmo default inseguro
    (`"dev-secret-change-in-production"`), mas a duplicação permite divergência futura.
    `config.JWT_SECRET` não é usado por ninguém.
11. `POST /api/auth/refresh` **não está** em `PUBLIC_ROUTES`. Como o middleware exige um JWT válido
    para todo `/api/*` não listado, o endpoint de renovação parece exigir o token que ele deveria
    renovar. Comportamento efetivo **precisa de verificação**.
12. `POST /upload` é público (`PUBLIC_PREFIXES` inclui `/upload`) e não tem limite de tamanho de
    arquivo — o conteúdo é lido inteiro em memória (`await file.read()`).
13. `GET /api/orcamento/dados` é público **e dispara `sync_tabela_master()`**, que faz uma
    sincronização completa com o Supabase. Qualquer chamada anônima aciona esse trabalho.
14. `GET /api/recs` lista os RECs de **todos** os usuários quando o Supabase responde
    (`recs.py:23` — sem filtro por `user_id`), enquanto o fallback SQLite também lista todos.
    Só a **exclusão** valida o dono.

### Configuração
15. `config.py:84` — `APP_MODE = os.environ.get("APP_MODE", "desktop" if IS_FROZEN else "desktop")`:
    os dois ramos do ternário são idênticos. O default é sempre `"desktop"`, inclusive no servidor,
    a menos que `APP_MODE=server` seja definido no ambiente.
16. `app.py:43` — em modo servidor, se `CORS_ORIGINS` não estiver definido, o CORS cai para `["*"]`
    **com `allow_credentials=True`**. O comentário no código chama isso de "fallback temporário".
17. `database.py:224-225` — sem `ADMIN_EMAIL`/`ADMIN_PASSWORD` no ambiente, o admin inicial é criado
    como `admin@local.com` / `admin123`.

### Qualidade de código
18. **Lógica de negócio duplicada** entre `static/script.js` e `static/resumo.js`
    (`isGray`, `processAtivoFormula`, cascata de classificação). Alterar só um lado faz as duas abas
    divergirem. Ver RULES.md, Regra 4.
19. `resumo.js:225` — `entAuto = '0'` é reatribuído logo após o bloco dos elos fusíveis, descartando
    a entidade `CHAVE` definida em `:209`. O ativo montado (`<qtd>-EF…`) é preservado, mas a
    classificação como `CHAVE` só volta a acontecer mais adiante, pela regra genérica de `-EF`
    (`:287`), que exige `opAuto !== 'M'`. **Comportamento precisa de confirmação:** é intencional ou
    resíduo de refatoração?
20. `orcamento_calc.py:187` — no passo 4 (busca parcial em `DESC_ATIVO`), quando há vários matches o
    código usa `matches[0]`, cuja ordem depende da iteração do dicionário. Resultado potencialmente
    não determinístico entre execuções com bases diferentes.
21. Capturas de exceção silenciosas (`except Exception: pass`) em todos os fallbacks de Supabase
    (`obras.py`, `recs.py`, `projetos.py`, `admin.py`). Falhas de nuvem são invisíveis em log.
22. `orcamento.py:18` reimporta `APIRouter, UploadFile, File, Request` já importados na linha 7;
    `projetos.py:57,61` importa `HTTPException` duas vezes.
23. `connectivity_monitor.py:39-56` manipula event loop do asyncio a partir de uma thread
    (`get_event_loop` / `new_event_loop` / `run_until_complete` / `loop.close()`) — padrão frágil.
    Sem efeito prático hoje, já que o worker nunca é iniciado.

### Dados
24. O repositório continha PDFs, DXFs e o `banco_resumo.db` com dados reais de obra, versionados
    antes das regras do `.gitignore`. **Removidos no commit `530d5cc`**; o histórico do Git ainda os
    contém. Mencionado aqui para que ninguém os re-adicione.
25. `routers/regras.py` (regras de conversão CABOS/OUTROS → totalizadora, botão "Salvar na Nuvem"
    em `resumo.html`) gravava e lia **somente no SQLite local**, apesar do rótulo "nuvem" —
    desktop e Render tinham cada um sua própria cópia, sem nenhum compartilhamento, e o dado se
    perdia ao trocar de ambiente ou reiniciar o servidor.
    ✅ **Corrigido — TASK-003 (2026-09-18):** criada a tabela `regras_conversao` no Supabase
    (`scripts/schema_supabase.sql`); `GET/POST /api/regras/conversao` passaram a ler da nuvem
    primeiro (fallback local se o Supabase não estiver configurado/alcançável) e a gravar em
    ambos, no mesmo padrão de `obras.py`. Diferente de `obras.py` (item 21 acima), a falha de
    sincronização aqui **não é engolida em silêncio** — vira `HTTPException` 500 visível ao
    usuário. Mesclado em `main` (PR #5, commit `cac4b85`) e **confirmado funcionando pelo usuário
    em produção** (desktop e online compartilhando as mesmas regras). Ver
    `.ai/tasks/TASK-003-18-09-2026.md` (status: CONCLUÍDA).

---

## Limitações atuais

- **Sem testes automatizados.** Nenhum arquivo de teste, nenhum framework, nenhum CI de teste.
- **Sem validação das regras técnicas do domínio.** As regras de engenharia em
  `prompt_rede_eletrica.txt` (poste de 10 m proibido em MT, CFU exige SUPL, trafo exige PR15 +
  PR220, mínimos de P50, estruturas tipo 3 não isoladas) são instruções ao modelo de linguagem —
  **o sistema não as verifica**.
- **Estado entre telas depende de `localStorage`.** Limpar o armazenamento do navegador perde o
  trabalho não salvo; não há recuperação.
- **Frontend sem build e sem modularização.** `resumo.js` tem 3.481 linhas e `resultado_orcamento.html`
  1.700 linhas com JS embutido.
- **Sincronização apenas sob demanda.** Sem os workers conectados, a nuvem só é consultada nas
  chamadas de rota; não há replay de operações feitas offline.
- **Sem paginação** em `/api/orcamento/dados`: a base técnica inteira é serializada a cada chamada.
- **Auto-update só no Windows** (script `.bat`, `subprocess.CREATE_NEW_CONSOLE`).
- **Operação `M` soma com `fator_r`** — no backend só `I`/`*I` contam como instalação; todo o
  resto cai no ramo de remoção. Ver GLOSSARY.md › OPERAÇÃO.

---

## Decisões recentes

Registradas como ADRs retroativos (reconstruídos a partir do código, não de registro histórico):

- **ADR-001** — Arquitetura offline-first com SQLite local e Supabase como nuvem
- **ADR-002** — Modo único de código para desktop e servidor, chaveado por `APP_MODE`
- **ADR-003** — Motor de regras de conversão no frontend, entre a classificação e o cálculo

Ver `.ai/decisions/`.

---

## Testes

**Status: inexistentes.**

- Arquivos de teste no repositório: **0**
- Framework configurado: nenhum (`requirements.txt` não inclui pytest ou equivalente)
- CI de testes: nenhum (os dois workflows só fazem deploy)

Melhor ponto de partida, se o usuário quiser cobertura: `services/orcamento_calc.py` —
função pura, sem I/O, sem dependências internas, concentrando as regras RN-03 a RN-10.

---

## Última atualização

**Data:** 2026-09-18
**Motivo:** Criação da estrutura de contexto IA-First (`.ai/`). Levantamento inicial do estado do
projeto por leitura completa do backend, dos serviços, do middleware e dos módulos de negócio do
frontend (`script.js`, `resumo.js`, `resultado_orcamento.html`).
**Alterações de código:** nenhuma.
