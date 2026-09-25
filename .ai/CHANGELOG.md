# CHANGELOG

> Histórico de alterações **relevantes** do projeto, do ponto de vista de contexto para IA:
> arquitetura, regras de negócio, fluxo de dados, estrutura de banco e comportamento importante.
>
> Não registre aqui alterações triviais (formatação, renomeação local, ajuste de texto).
> Formato: mais recente primeiro.

---

## 2026-09-25 — TASK-006: classificação do leitor vira motor de regras editável por projeto

**Tipo:** nova arquitetura (motor de regras) · **Tarefa:** `.ai/tasks/TASK-006-25-09-2026.md`

A classificação automática do leitor (entidade/operação/ativo a partir de texto/cor/layer),
antes embutida em `processAtivoFormula`/`computeRowLogic`/`updateRowLogic`/`autoClassifyEntidade`
(duplicada e já divergente entre `script.js` e `resumo.js`), virou duas tabelas de regras por
projeto — **Processamento** (texto/cor/layer → operação/ativo, em 3 fases: seleção de ramo,
pós-processamento, overrides por cor/layer) e **Classificação** (operação/ativo/cor/layer/texto →
operação/entidade, reutilizada também em `resumo.js`) — interpretadas por um motor genérico
(`static/regras_leitor_engine.js`).

**Decisão de unificação:** o comportamento das duas abas foi unificado usando `script.js` (o mais
completo) como canônico. Dois bugs reais do código original foram corrigidos por decisão do
usuário em vez de replicados: elo fusível nunca classificava como `CHAVE` (variável `uAtivo`
desatualizada) e a detecção de cabo multiplexado `M3x1` nunca funcionava (regex case-sensitive
contra texto já maiúsculo).

**Alterado/criado:**
- `static/regras_leitor_engine.js` (novo) — motor de regras, roda em navegador e Node
- `data/regras_leitor_{processamento,classificacao}_seed.json` (novos) — transcrição fiel de
  `script.js` (45 + 22 regras), semeada automaticamente para os projetos `027`/`229`
- `database.py`, `scripts/schema_supabase.sql` — duas tabelas novas por projeto
- `routers/regras_leitor.py` (novo) — CRUD nuvem-primeiro, mesmo padrão de `regras_conversao`
- `static/script.js`, `static/resumo.js` — cascata antiga removida, religada ao motor; botão
  "Regras do Leitor" (visualização) + aviso/escolha de projeto alternativo quando o projeto
  selecionado não tem regras cadastradas

**Validado:** porta fiel do código original para Node.js ("oráculo"), comparada contra o motor
novo em 66 casos sintéticos (63 idênticos, 3 divergências aprovadas — os bugs corrigidos, 0
falhas não explicadas); servidor real + navegador real via Playwright confirmando que o motor
carrega as regras pela API de verdade e produz os mesmos resultados dentro do navegador.

**Não corrigido nesta etapa:** `.ai/CONTEXT.md`/`.ai/ARCHITECTURE.md` ainda descrevem a
classificação como lógica fixa em código (desatualizado); nenhuma UI de edição das regras foi
construída (só visualização, que era o que o critério de aceite pedia).

---

## 2026-09-25 — TASK-005: obras, RECs e regras de IA isolados por projeto

**Tipo:** schema + backend + frontend · **Tarefa:** `.ai/tasks/TASK-005-25-09-2026.md`

Obras salvas, RECs e regras de IA (aprendizado do chat) apareciam em todos os projetos,
independentemente de onde foram salvos. Adicionada a coluna/campo `projeto`/`projeto_codigo` em
`obras`, `historico_rec` e `regras` (SQLite + Supabase), com filtro em todas as rotas de
leitura/escrita relevantes. Regras de IA, que nunca tinham sincronização com o Supabase (só
SQLite local de cada instância), ganharam esse suporte pela primeira vez, no mesmo padrão de
`regras_conversao` (TASK-003).

**Correção de convenção:** o valor de backfill/`DEFAULT` usado é o **código** do projeto (`"229"`
para RONDÔNIA), não o nome — corrigido depois de descoberto, durante o teste end-to-end da
TASK-006, que o nome não é a chave técnica usada em nenhum outro lugar do app para agrupar por
projeto (`regras_conversao.projeto_codigo`, `localStorage['projeto_selecionado_codigo']`).

**Validado:** simulado um banco "antigo" sem as colunas novas e confirmado o backfill via
`ALTER TABLE ... DEFAULT`; testado via HTTP real que salvar/listar obras, RECs e regras de IA em
projetos diferentes fica isolado, e que a listagem sem filtro preserva o comportamento anterior.

---

## 2026-09-19 — TASK-004-19-09-2026: correção de cache de JS no navegador (Revisão)

**Tipo:** correção de infraestrutura · **Tarefa:** `.ai/tasks/TASK-004-19-09-2026.md`

Alterações nos arquivos JS só carregavam em aba privativa (incognito). Causa raiz: o navegador
armazenava um "hard cache" dos arquivos `.html` (ex: `/static/resumo.html`). Por causa
do cache local, o navegador nunca pedia o arquivo ao servidor, e assim nunca recebia os
cabeçalhos de `NO_CACHE` criados. Para invalidar o hard cache local, era preciso alterar 
a URL de navegação (ex: de `/static/resumo.html` para `/resumo`).

**Alterado:**
- `app.py` — Novas rotas `@app.get("/resultado_orcamento")` e `@app.get("/orcamento")` servindo as páginas sem cache.
- `static/index.html`, `static/resumo.html`, `static/script.js`, `static/resumo.js` — Alteradas todas as navegações `window.location` e `window.open` que apontavam para `/static/*.html`, alterando para as rotas limpas do `app.py` (ex: `/resumo`, `/resultado_orcamento`).
- `middleware/nocache_middleware.py` — Injeta `NO_CACHE_HEADERS` em respostas `/static/*.html` (implementado no passo anterior, mantido por precaução).

**Validado:** aguardando confirmação do usuário em aba normal após reinício do servidor.

---


**Tipo:** melhoria de UX + segurança de processo · **Tarefa:** `.ai/tasks/TASK-002-19-09-2026.md`

Ao iniciar o programa, o navegador abria diretamente na página principal com dados em cache,
sem exigir novo login. Além disso, o processo Python na porta 8000 ficava ativo mesmo após
fechar todas as abas do navegador, sem forma simples de encerrá-lo pela UI.

**Alterado:**
- `app.py` — URL de abertura mudada de `/` para `/login?new_session=1`; novo endpoint
  `GET /api/shutdown` que envia `SIGTERM` ao processo após 500 ms (bloqueado com HTTP 403
  em `APP_MODE == "server"`).
- `middleware/auth_middleware.py` — `/api/shutdown` adicionado a `PUBLIC_ROUTES`.
- `static/login.html` — script IIFE que, ao detectar `?new_session=1`, limpa `auth_token`,
  `user_id`, `user_email` e `is_admin` do `localStorage` e pré-preenche o campo email com
  o último email usado (conveniência — senha nunca armazenada).
- `static/index.html` — botão "Encerrar Servidor" adicionado ao header; ao confirmar o
  diálogo, chama `GET /api/shutdown` e exibe mensagem "Servidor encerrado. Pode fechar esta aba."

**Validado:** servidor real reiniciado; navegador abriu em `/login?new_session=1` ✅;
login realizado com dados da nuvem carregados corretamente ✅; botão "Encerrar Servidor"
acionado com confirmação — log: `Encerramento solicitado pelo usuário via /api/shutdown.` ✅

---


**Tipo:** melhoria de UI · **Tarefa:** `.ai/tasks/TASK-001-19-09-2026.md`

A página inicial (`index.html`) exibia apenas dois caminhos de entrada: "Procurar Arquivo"
e "Montar Projeto". O usuário não tinha como navegar diretamente para a tela de resultado
de orçamento (`resultado_orcamento.html`) sem antes passar por outra tela.

**Alterado:**
- `static/index.html` — adicionados `<p>ou acesse os orçamentos</p>` e
  `<button id="btn-manipular-orcamentos">Manipular Orçamentos</button>` no `upload-card`,
  após o botão "Montar Projeto". O botão navega para `/static/resultado_orcamento.html`
  via `window.location.href`, no mesmo padrão já usado pelo botão "Montar Projeto".

**Não alterado:** nenhuma funcionalidade das abas, nenhuma rota de API, nenhum banco
de dados, nenhum contrato de `localStorage`.

**Validado:** servidor já em execução; alteração visível ao recarregar a página inicial.

---


**Tipo:** nova tabela + correção de código · **Tarefa:** `.ai/tasks/TASK-003-18-09-2026.md`

O botão "Salvar na Nuvem" das regras de conversão (CABOS/OUTROS → totalizadora, em
`resumo.html`/`resumo.js`) nunca escrevia no Supabase — `routers/regras.py` lia e gravava apenas
na tabela `configuracoes` do SQLite local. Cada instância do backend (desktop, Render) tem seu
próprio arquivo de banco, sem nenhum compartilhamento; por isso as regras "somem" ao trocar de
ambiente ou reiniciar o servidor.

**Alterado:**
- `scripts/schema_supabase.sql` — nova tabela `regras_conversao` (`projeto_codigo` PK,
  `regras_json`), bloco idempotente no mesmo padrão das migrações anteriores.
- `routers/regras.py:get_regras_conversao` — passa a consultar o Supabase primeiro (fonte de
  verdade quando configurado/alcançável), com fallback para o SQLite local, no mesmo padrão de
  `routers/obras.py:get_obras`.
- `routers/regras.py:save_regras_conversao` — passa a gravar local **e** fazer upsert no
  Supabase; diferente de `obras.py` (que engole erro de Supabase em silêncio — problema 21 do
  `STATE.md`, não corrigido aqui por estar fora do escopo pedido), aqui a falha de sincronização
  vira `HTTPException` 500 visível, no padrão já aprovado na TASK-002.
- `static/resumo.js:salvarRegrasNuvem` — passa a mostrar `data.detail` (mensagem real de erro) em
  vez de um texto genérico, no mesmo padrão já usado em `orcamento.html`.

**Validado:** servidor real subido a partir de uma cópia isolada do projeto (stubs para `ezdxf`,
`pymupdf` e `supabase`, técnica já usada nas tarefas anteriores), autenticado como admin:
- `POST` sem Supabase configurado → `200`, grava só local (comportamento preservado)
- `POST` com Supabase configurado porém falhando → `500` com mensagem clara, e SQLite local
  confirmado gravado mesmo assim
- `GET` de um `projeto_codigo` que só existia no Supabase "fake" (nunca gravado localmente) →
  retornou o dado da nuvem corretamente, confirmando que a integração é real e não um fallback
  disfarçado
- `POST` seguido de `GET` do mesmo `projeto_codigo`, com Supabase "fake" funcionando → o `GET`
  retornou exatamente o que foi salvo, vindo da nuvem

**Mesclado em `main`** via PR #5 (commit de merge `cac4b85`). Usuário rodou a migração no
Supabase real, redeployou e **confirmou**: regras salvas na nuvem persistem tanto no desktop
quanto na versão online. **TASK-003 encerrada.**

---

## 2026-09-18 — TASK-002 (continuação): `origem` era removida do payload antes do envio ao Supabase

**Tipo:** correção de código · **Tarefa:** `.ai/tasks/TASK-002-18-09-2026.md`

Depois que o usuário rodou a migração `ALTER TABLE ... ADD COLUMN origem` no Supabase real e
reimportou a base master completa, `origem` continuava `null` em todas as linhas, sem nenhum erro
sendo reportado. Achado um trecho em `routers/admin.py` (`upload_master_csv` e `sync_master_all`)
que filtrava `origem` para fora do dicionário antes de cada `INSERT`/`DELETE` no Supabase
(`rows_supabase = [{k: v for k, v in r.items() if k != 'origem'} for r in rows_to_insert]`), com o
comentário "A coluna 'origem' não existe no Supabase - removemos antes do insert". Esse trecho
veio de um commit anterior à TASK-002 (`08cd044`, 17/09), feito quando a coluna de fato ainda não
existia — e sobreviveu a um merge de `main` de volta para a branch de trabalho sem gerar conflito,
continuando a remover `origem` mesmo depois do schema já ter a coluna. Por isso o `INSERT` nunca
falhava (sem erro visível) e `origem` nunca chegava na nuvem, independentemente de qualquer outro
fix já aplicado.

**Alterado:** removida a filtragem `rows_supabase = [...]` nos dois handlers; ambos voltam a
enviar `origem` ao Supabase como qualquer outra coluna, no mesmo padrão já usado por
`add_master_row`.

**Mesclado em `main`** via PR #3 (commit de merge `36de3aa`). Usuário redeployou/reconstruiu o
ambiente, reimportou a base master e **confirmou**: `ORIGEM` aparece corretamente preenchida em
`GET /api/orcamento/dados` após recarregar a tela. **TASK-002 encerrada.**

---

## 2026-09-18 — TASK-002: coluna `origem` restabelecida na tabela master de orçamento

**Tipo:** correção de schema + bug de código + diagnóstico · **Tarefa:**
`.ai/tasks/TASK-002-18-09-2026.md`

Investigado por que a coluna `ORIGEM`, ao ser importada num CSV para a tabela master, não ficava
disponível depois de recarregar a tela. Dois problemas distintos, mesma raiz de família da
TASK-001: schema do Supabase desatualizado em relação ao código.

**Causa raiz:** `tabela_orcamento_master` no Supabase real nunca teve a coluna `origem`
(`scripts/schema_supabase.sql` não a declarava). `admin.py:upload_master_csv` lia e gravava
`origem` certo no CSV e no SQLite local, mas o `INSERT` espelhado no Supabase falhava com o erro
**engolido em silêncio** — a resposta ao admin continuava dizendo sucesso. Na sequência, o próximo
`GET /api/orcamento/dados` puxava a master antiga da nuvem e sobrescrevia o local, apagando o dado
que tinha acabado de ser importado.

**Bug de código independente:** `admin.py:sync_master_all` (botão "Sincronizar Tudo com a Nuvem")
nunca lia nem gravava `origem` em lugar nenhum, apesar do frontend já enviar o campo — não
dependia do schema do Supabase, era um bug puro no Python.

**Alterado:**
- `scripts/schema_supabase.sql` — coluna `origem` adicionada ao `CREATE TABLE
  tabela_orcamento_master` + migração idempotente para instalações existentes
- `routers/admin.py:sync_master_all` — passou a ler/gravar `origem`, alinhado com
  `upload_master_csv`/`add_master_row`
- `routers/admin.py:upload_master_csv` — falha de sincronização com o Supabase agora vira
  `HTTPException` 500 visível ao admin (`"Tabela local atualizada, mas falhou ao sincronizar com a
  nuvem (Supabase): {erro}"`), em vez de log silencioso; adicionado `except HTTPException: raise`
  para essa exceção não ser reembrulhada pelo handler genérico da função (que trocaria o status
  500 por 400 e duplicaria a mensagem) — **melhoria aprovada explicitamente pelo usuário**
- `static/orcamento.html` — o botão "Sobrescreve a tabela mestre" passou a mostrar a mensagem real
  de erro (`data.detail`), no mesmo padrão já usado pelo botão "Sincronizar Tudo com a Nuvem"

**Validado:** servidor real subido localmente; `POST /api/admin/sync-master-all` com `origem`
preenchida confirmado persistindo no SQLite; `POST /api/admin/upload-master` com Supabase
configurado mas inválido confirmado retornando `500` com mensagem clara (não mais `400`
reembrulhado) e, mesmo assim, salvando `origem` corretamente no SQLite local.

**Não corrigido nesta etapa** (depende de ação do usuário, fora do alcance deste ambiente): rodar
o `ALTER TABLE` no Supabase real.

**Achados registrados, não corrigidos** (fora do escopo pedido): o botão "Importar CSV Local"
chama um endpoint (`/api/upload/csv-orcamento`) que não existe no backend; a tabela pessoal do
usuário (`routers/orcamento.py`) também não trata `origem`, mas nunca sincroniza com a nuvem.

---

## 2026-09-18 — TASK-001: login/cadastro via Supabase restabelecidos (desktop + Render.com)

**Tipo:** correção de schema (dados) + diagnóstico (código) · **Tarefa:**
`.ai/tasks/TASK-001-18-09-2026.md` · **Status final: CONCLUÍDA**

Investigado com o usuário por que login e cadastro de usuário (Supabase) haviam parado de
funcionar. RLS descartado (estava desabilitado). Confirmado que a tabela real `usuarios_nuvem`
do usuário tinha `is_admin` mas não tinha `role` — coluna que `routers/admin.py` grava em
`POST /api/admin/users` (falha visível, com mensagem enganosa de "e-mail duplicado") e em
`PUT /api/admin/users/{id}/role` (falha silenciosa).

**Alterado no código:**
- `scripts/schema_supabase.sql` — migração idempotente adicionando `role` a `usuarios_nuvem`
- `routers/health.py` — `GET /api/health` ganhou `supabase_configured`, `supabase_reachable` e
  `usuarios_nuvem_has_rows`, sem expor dados de usuário (rota é pública)

**Validado:** servidor real subido localmente (dependências pesadas substituídas por stubs, já
que não influenciam este diagnóstico); `/api/health` testado com Supabase não configurado e com
credenciais configuradas porém inválidas — os dois casos respondem corretamente, sem quebrar a
rota.

**Feito pelo usuário, fora deste ambiente (sem acesso a Supabase/Render a partir daqui):**
1. `ALTER TABLE ... ADD COLUMN role` no Supabase real → destravou login e cadastro
2. Efeito colateral descoberto e corrigido: o `DEFAULT 'operador'` do passo 1 fez backfill em
   **todas** as linhas existentes, inclusive contas admin — derrubando o acesso ao painel admin
   até um `UPDATE` reconciliando `role` com `is_admin` (ver detalhe em STATE.md, problema 8, e no
   arquivo da tarefa)
3. Configuradas `SUPABASE_URL`/`SUPABASE_KEY` nas env vars do serviço Render.com (nunca haviam
   sido definidas lá — ambiente de deploy adicional, não documentado antes desta tarefa) e forçado
   um redeploy manual para pegar o código já mesclado em `main`

**Resultado confirmado pelo usuário:** login, painel admin e cadastro de usuário funcionando tanto
no desktop quanto na aplicação online (Render.com).

**Não alterado** (fora do escopo pedido): qualquer lógica de `orcamento_calc.py` ou das tabelas de
orçamento.

---

## 2026-09-18 — Estrutura de contexto IA-First criada

**Tipo:** documentação · **Alterações de código: nenhuma**

Criada a pasta `.ai/` como fonte única de contexto do projeto para desenvolvimento assistido por IA.

### Estado do projeto no momento do levantamento
- Versão `2.0.0`, branch `claude/beautiful-pasteur-2tdk18`, commit base `530d5cc`
- Backend FastAPI com 11 routers e 9 services; frontend HTML/CSS/JS vanilla
- SQLite local (11 tabelas) + Supabase (5 tabelas) em arquitetura offline-first
- Sem nenhum teste automatizado

### Arquivos analisados para o levantamento
**Raiz:** `app.py`, `config.py`, `database.py`, `models.py`, `websocket_manager.py`, `trigger.py`,
`requirements.txt`, `.env.example`, `.gitignore`, `build.bat`, `README.md`

**Routers:** `auth.py`, `upload.py`, `orcamento.py`, `regras.py`, `recs.py`, `obras.py`,
`projetos.py`, `ai_chat.py`, `admin.py`, `health.py`, `update.py`

**Services:** `pdf_service.py`, `dxf_service.py`, `orcamento_calc.py`, `supabase_client.py`,
`sync_service.py`, `offline_queue.py`, `realtime_sync.py`, `connectivity_monitor.py`,
`auto_updater.py`

**Middleware:** `auth_middleware.py`

**Frontend:** `script.js`, `resumo.js` (seções de lógica de negócio, cálculo de `qtdAtivos`, motor de
regras e camada unificada), `resultado_orcamento.html` (cálculo e REC), `resumo.html`,
`orcamento.html`, `index.html`, `login.js`, `admin.js`, `auth_fetch.js`

**Dados e infra:** `data/tabela_seed.csv`, `prompt_rede_eletrica.txt`,
`scripts/schema_supabase.sql`, `scripts/release.py`,
`.github/workflows/deploy-backend.yml`, `.github/workflows/deploy-frontend.yml`

### Estrutura criada
```
.ai/
├── CONTEXT.md        identificação, tecnologias, fluxo, entradas/saídas, 17 regras de negócio
├── ARCHITECTURE.md   camadas, pipeline, módulos, banco, API, exemplos entrada→regra→saída
├── GLOSSARY.md       termos do sistema e ativos de rede elétrica, com origem de cada definição
├── RULES.md          12 regras de trabalho para agentes + protocolo de 6 fases
├── STATE.md          implementado, não conectado, 24 problemas conhecidos, limitações
├── CHANGELOG.md      este arquivo
├── tasks/TEMPLATE.md modelo de tarefa
└── decisions/        TEMPLATE.md + ADR-001, ADR-002, ADR-003 (retroativos)
```

### Documentado a partir do código
- Pipeline completo: extração → classificação → separação → camada unificada → motor de regras →
  cálculo → resultado
- 17 regras de negócio (RN-01 a RN-17) com referência a arquivo e linha
- Cascata de 5 passos de resolução de ativo e expansão ativo → componente → códigos
- Precedência da base técnica (usuário > master > seed) e prioridade de projeto
- Contratos de `localStorage` entre as telas

### Registrado como pendente ou divergente (não corrigido)
- Três workers de background escritos e nunca iniciados (`offline_queue`, `realtime_sync`,
  `connectivity_monitor`)
- `Dockerfile` e `fly.toml` ausentes, embora exigidos pelo workflow de deploy
- Divergências de schema entre SQLite e Supabase (`origem`, `is_admin`)
- Três conjuntos inconsistentes de roles (`viewer` vs. `operador` vs. `editor`)
- Lógica de classificação duplicada entre `script.js` e `resumo.js`
- Ausência total de testes

### Pontos que necessitam confirmação do usuário
Registrados em GLOSSARY.md §3 e na seção final do relatório de inicialização: `K1`, `L1`, `L2`,
`TR+1/+2/+3`, `U23`, `U32`, significado de `REC`, `RS M`/`RS T`, `FLY`, `CV`, e o propósito da
coluna `filtro`.

---

## Antes de 2026-09-18

Histórico anterior não registrado neste formato. As mensagens de commit existentes
(`"comit changes"`, `"melhorias"`, `"bug fixes"`, `"mudanças"`) não permitem reconstruir com
segurança quais alterações de arquitetura ou de regra de negócio ocorreram em cada ponto.

`[NÃO DETERMINADO PELO CÓDIGO]` — para reconstruir o histórico anterior seria necessário o
relato do usuário.

Um marco é identificável pelo próprio código: o salto para a versão `2.0.0`, descrito em
`app.py:2` como *"Ponto de entrada da aplicação FastAPI (v2.0 - Profissionalizado)"*, introduziu a
separação em `routers/`, `services/` e `middleware/`, os dois modos de operação e a autenticação JWT.
