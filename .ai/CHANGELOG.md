# CHANGELOG

> Histórico de alterações **relevantes** do projeto, do ponto de vista de contexto para IA:
> arquitetura, regras de negócio, fluxo de dados, estrutura de banco e comportamento importante.
>
> Não registre aqui alterações triviais (formatação, renomeação local, ajuste de texto).
> Formato: mais recente primeiro.

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
