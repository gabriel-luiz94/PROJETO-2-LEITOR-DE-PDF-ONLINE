# ARQUITETURA

> Documenta a arquitetura **real**, como implementada. Cada afirmação aponta para o arquivo e a
> linha de origem. O que não pôde ser determinado pelo código está marcado como
> `[NÃO DETERMINADO PELO CÓDIGO]`.

---

## 1. Visão geral em camadas

```
┌─────────────────────────────────────────────────────────────────────┐
│  FRONTEND (HTML/CSS/JS vanilla, sem build)                          │
│  index.html + script.js     → extração e classificação inicial      │
│  resumo.html + resumo.js    → Cabos/Outros, regras, totalizadora, IA│
│  resultado_orcamento.html   → cálculo final, REC, exportação        │
│  orcamento.html             → edição da base técnica                │
│  admin.html + admin.js      → usuários, tabela master, audit log    │
│  login.html + login.js      → autenticação                          │
└──────────────────────┬──────────────────────────────────────────────┘
                       │ HTTP /api/*  +  WebSocket /ws  +  localStorage
┌──────────────────────┴──────────────────────────────────────────────┐
│  app.py — FastAPI                                                   │
│  CORS → AuthMiddleware (JWT) → routers                              │
├─────────────────────────────────────────────────────────────────────┤
│  routers/    auth · obras · regras · recs · projetos · orcamento     │
│              ai_chat · upload · health · admin · update             │
├─────────────────────────────────────────────────────────────────────┤
│  services/   pdf_service · dxf_service · orcamento_calc             │
│              supabase_client · sync_service · offline_queue         │
│              realtime_sync · connectivity_monitor · auto_updater    │
├─────────────────────────────────────────────────────────────────────┤
│  database.py (SQLite)   ◄── sync ──►   Supabase (Postgres)          │
└─────────────────────────────────────────────────────────────────────┘
```

---

## 2. Fluxo de dados completo (pipeline do orçamento)

```
   ARQUIVO (PDF / DXF)
        │
        ▼
┌───────────────────┐
│ EXTRAÇÃO          │  POST /upload  →  pdf_service | dxf_service
│                   │  saída: {pagina, texto, fonte, tamanho, cor, flags, layer}
└────────┬──────────┘
         ▼
┌───────────────────┐
│ CLASSIFICAÇÃO     │  script.js (aba principal) → usuário revisa/edita
│ (CONTEXTO)        │  cor + layer + texto  →  {entidade, operacao, ativo}
└────────┬──────────┘
         │  localStorage['processar_dados']
         ▼
┌───────────────────┐
│ SEPARAÇÃO         │  resumo.js: computeRowLogic() por linha
│                   │
│   entidade=CABO ──────────► TABELA CABOS
│   entidade∈{RAMAIS,IP,APOIO*} ──► modal RAMAIS (_ramaisData)
│   demais (≠0) ────────────► TABELA OUTROS
└────────┬──────────┘
         ▼
┌───────────────────┐
│ CAMADA UNIFICADA  │  syncTotalizadora() → rawItems[]
│                   │  { baseId:"TOT-n", obs:entidade, operacao,
│                   │    ativo, qtd, desc, origem:"CABOS"|"OUTROS" }
│                   │  ▸ baseId preserva a rastreabilidade da linha de origem
│                   │  ▸ origem preserva a procedência do registro
└────────┬──────────┘
         ▼
┌───────────────────┐
│ MOTOR DE REGRAS   │  Tabela de Regras de Conversão (por projeto)
│                   │  match: origem + op_de + ativo_de
│                   │  aplica: op_para, ativo_para, fator,
│                   │          arredondamento, val_min, val_max
│                   │  ADIÇÃO → mantém original + gera novo
│                   │  SUBST  → substitui o original
└────────┬──────────┘
         ▼
┌───────────────────┐
│ TOTALIZADORA      │  editável manualmente pelo usuário
└────────┬──────────┘
         │  localStorage['orcamentoPayload'] = {cabos[], outros[], projeto}
         ▼
┌───────────────────┐
│ CÁLCULO           │  POST /api/orcamento/calcular
│                   │  processar_calculo(cabos, outros, projeto, base_técnica)
│                   │    1. índice ATIVO → linhas (com prioridade de projeto)
│                   │    2. parse dos formatos CABOS e OUTROS → ativos_qtd[]
│                   │    3. cascata de resolução (5 passos) → componente → códigos
│                   │    4. soma_i / soma_r por (codigo, mdo)
│                   │    5. uma linha por operação com soma > 0
└────────┬──────────┘
         ▼
┌───────────────────┐
│ RESULTADO         │  {operacao, mdo, codigo, desc_codigo, filtro, total}
│                   │  + nao_encontrados[]
└────────┬──────────┘
         ▼
   REC (historico_rec)  ·  export .rec/CSV  ·  impressão
```

---

## 3. Módulos e responsabilidades

### 3.1 Raiz

| Arquivo | Responsabilidade |
|---|---|
| `app.py` | Cria o FastAPI, configura CORS por modo, registra `AuthMiddleware`, chama `init_db()`, monta `/static`, serve as páginas HTML, expõe `/ws` e `/api/version`, inclui todos os routers, e faz o bootstrap desktop (pywebview) ou servidor (uvicorn headless) |
| `config.py` | Versão, logging, resolução de caminhos (dev vs. PyInstaller `_MEIPASS`), `DB_PATH`, `STATIC_DIR`, `PROMPT_PATH`, `SEED_CSV_PATH`, `APP_MODE`, `JWT_SECRET`, `SERVER_URL` |
| `database.py` | Conexões SQLite (WAL), hash/verificação de senha, migração SHA-256→bcrypt, `init_db()` com criação de tabelas, `ALTER TABLE` idempotentes, admin inicial e seed do CSV |
| `models.py` | Modelos Pydantic de request |
| `websocket_manager.py` | `ConnectionManager` com conexões por `user_id` e anônimas; `broadcast` e `send_personal_message` |
| `trigger.py` | Utilitário Windows/Tkinter: escolhe um PDF e o envia ao app via `/trigger-file`, subindo o app se necessário |

### 3.2 `routers/`

| Router | Prefixo | Responsabilidade |
|---|---|---|
| `auth.py` | `/api/auth` | Login (Supabase → SQLite), refresh de token, JWT + bcrypt |
| `upload.py` | — | `POST /upload` (PDF/DXF), `GET /extract-local` (só `.exe`), `POST /api/importar-rec-pdf`, `GET /trigger-file` |
| `orcamento.py` | `/api/orcamento` | `upload` e `salvar` da base (admin), `dados` (merge), `search`, `detalhes`, **`calcular`** |
| `regras.py` | `/api/regras` | Regras de aprendizado da IA (`GET`/`POST`) e **regras de conversão por projeto** (`/conversao`) |
| `recs.py` | — | `/api/recs` e `/api/rec/*`: histórico de RECs com preservação de REC de terceiros |
| `obras.py` | `/api/obras` | CRUD de obras por usuário (Supabase com fallback SQLite) |
| `projetos.py` | `/api/projetos` | Lista mesclada local+nuvem; cadastro só admin |
| `ai_chat.py` | `/api/gemini` | Listagem de modelos e chat com Gemini/OpenAI, injetando `prompt_rede_eletrica.txt` e as regras aprendidas |
| `admin.py` | `/api/admin` | Usuários (CRUD, role, senha), tabela master (add/upload CSV/sync completo), audit log |
| `health.py` | — | `/api/health`, `/api/backup/export`, `/api/health/sync-master` |
| `update.py` | `/api/update` | `check` (público, consulta `configuracoes`) e `apply` (só desktop) |

### 3.3 `services/`

| Service | Responsabilidade | Wired? |
|---|---|---|
| `pdf_service.py` | Extração de spans do PDF com cor e flags | ✅ usado por `upload.py` |
| `dxf_service.py` | Extração de DXF com resolução de cor e quebra de MTEXT | ✅ usado por `upload.py` |
| `orcamento_calc.py` | **Núcleo de cálculo do orçamento** (puro, sem I/O) | ✅ usado por `orcamento.py` |
| `supabase_client.py` | Singleton thread-safe do cliente Supabase, credenciais só via env | ✅ usado por vários routers |
| `sync_service.py` | `sync_tabela_master()` (baixa master paginada do Supabase) e `get_merged_orcamento()` (precedência da base) | ✅ usado por `orcamento.py` e `health.py` |
| `offline_queue.py` | Fila persistente `sync_log` + worker de replay a cada 30 s | ⚠️ **definido, nunca chamado** |
| `realtime_sync.py` | Listener do Supabase Realtime → atualiza SQLite e emite WebSocket | ⚠️ **definido, nunca chamado** |
| `connectivity_monitor.py` | Testa o Supabase a cada 15 s e emite status via WebSocket | ⚠️ **definido, nunca chamado** |
| `auto_updater.py` | Verifica versão no servidor, baixa `.exe` e reinicia via `.bat` | ✅ usado por `update.py` |

⚠️ `start_offline_queue_worker()`, `start_realtime_sync()`, `start_connectivity_monitor()` e
`enqueue_operation()` **não são chamados em lugar nenhum** do código. A sincronização em produção
acontece apenas de forma síncrona, sob demanda (`sync_tabela_master()` dentro de
`GET /api/orcamento/dados` e de `POST /api/health/sync-master`). Ver STATE.md.

### 3.4 `middleware/auth_middleware.py`

Middleware global que protege `/api/*`. Extrai o JWT de `Authorization: Bearer` ou do
query param `?token=`, valida, e injeta `request.state.user = {user_id, email, role}`.

- **`PUBLIC_ROUTES`** (sem auth): `/api/auth/login`, `/api/health`, `/api/health/sync-master`,
  `/api/update/check`, `/api/orcamento/dados`, `/api/orcamento/calcular`, `/api/orcamento/search`,
  `/api/orcamento/detalhes`, `/trigger-file`
- **`OPTIONAL_AUTH_ROUTES`**: `/api/projetos`
- **`PUBLIC_PREFIXES`**: `/static`, `/`, `/login`, `/resumo`, `/admin`, `/upload`, `/extract-local`,
  `/ws`, `/docs`, `/openapi.json`

Helpers: `create_jwt_token` (24 h), `create_refresh_token` (30 d), `decode_jwt_token`,
`require_role(*roles)` (dependency, **não usado** — os routers chamam `_require_admin` /
comparam `role` manualmente), `get_current_user_from_state`.

---

## 4. Banco de dados

### 4.1 SQLite local (`banco_resumo.db`) — `database.py:init_db`

| Tabela | Chave | Campos principais | Observação |
|---|---|---|---|
| `obras` | `id` TEXT | `nome`, `data`, `dados_json`, `user_id`, `updated_at` | Projeto em edição |
| `regras` | `id` AUTOINC | `conteudo`, `embedding`, `updated_at` | Regras de aprendizado da IA. `embedding` existe mas **nunca é preenchido** |
| `configuracoes` | `chave` | `valor` | Chave/valor: API key da IA, `regras_conversao_<projeto>`, `latest_desktop_version`, `latest_desktop_url` |
| `tabela_orcamento` | `id` AUTOINC | `user_id`, `ativo`, `desc_ativo`, `componente`, `projeto`, `mdo`, `codigo`, `desc_codigo`, `fator_i`, `fator_r`, `filtro`, `origem` | **Base técnica** do usuário; `user_id IS NULL` = seed |
| `tabela_orcamento_master` | `id` AUTOINC | idem, sem `user_id` | Base oficial, espelho do Supabase |
| `sessoes` | `user_id` | `email`, `access_token`, `refresh_token`, `is_admin`, `role`, `ultimo_login` | Sessão local |
| `usuarios_locais` | `id` AUTOINC | `email` UNIQUE, `senha_hash`, `is_admin`, `role`, `ativo`, timestamps | Login offline e painel admin |
| `historico_rec` | `numero_obra` | `dados_json`, `data_criacao`, `user_id`, `updated_at` | REC da obra |
| `projetos` | `nome` | `codigo`, `updated_at` | Seed: PARAIBA/027, RONDONIA/229 |
| `sync_log` | `id` AUTOINC | `tabela`, `operacao`, `registro_id`, `dados_json`, `timestamp`, `sincronizado`, `tentativas`, `erro` | Fila offline — **nunca alimentada** (`enqueue_operation` não é chamado) |
| `audit_log` | `id` AUTOINC | `user_id`, `email`, `action`, `table_name`, `record_id`, `details`, `created_at` | Alimentado apenas por `admin.py:_audit` |

Migrações: `init_db()` roda `ALTER TABLE … ADD COLUMN` dentro de `try/except
sqlite3.OperationalError`, tornando a evolução do schema idempotente. Não há ferramenta de migração.

### 4.2 Supabase (`scripts/schema_supabase.sql`)

`historico_rec`, `obras`, `projetos`, `tabela_orcamento_master`, `usuarios_nuvem` + índices.

⚠️ Divergências entre o schema versionado e o que o código escreve:
- `tabela_orcamento_master` no schema **não tem a coluna `origem`**, mas `admin.py:upload-master` e
  `admin.py:add-row` a enviam e `sync_service` a lê.
- `usuarios_nuvem` no schema **não tem a coluna `is_admin`**, mas `auth.py` e `admin.py` leem e
  escrevem `is_admin`.

### 4.3 Estratégia offline-first

```
LEITURA da base técnica (get_merged_orcamento):
   tabela_orcamento WHERE user_id = <atual>     ← edições do usuário (maior prioridade)
        ↓ vazio
   tabela_orcamento_master                      ← master sincronizada do Supabase
        ↓ vazio
   tabela_orcamento WHERE user_id IS NULL       ← seed do CSV

LOGIN (auth.py):
   Supabase usuarios_nuvem  →  fallback SQLite usuarios_locais

DEMAIS ENTIDADES (obras, recs, projetos):
   tenta Supabase; em qualquer exceção, cai para o SQLite silenciosamente (except: pass)
```

---

## 5. API — mapa de rotas

| Método | Rota | Auth | Descrição |
|---|---|---|---|
| GET | `/`, `/login`, `/resumo`, `/admin` | pública | Páginas HTML (no-cache) |
| GET | `/api/version` | JWT | Versão e modo |
| WS | `/ws` | pública | Eventos `load_file`, `connectivity` |
| POST | `/api/auth/login` | pública | Login → access + refresh token |
| POST | `/api/auth/refresh` | pública¹ | Renova o access token |
| POST | `/upload` | pública | Extrai PDF/DXF |
| GET | `/extract-local?path=` | pública | Só quando `IS_FROZEN` |
| POST | `/api/importar-rec-pdf` | JWT | Importa Orçamento+Lista |
| GET | `/trigger-file?path=` | pública | Broadcast de `load_file` |
| GET | `/api/orcamento/dados` | **pública** | Base técnica mesclada (dispara `sync_tabela_master`) |
| POST | `/api/orcamento/calcular` | **pública** | **Cálculo do orçamento** |
| GET | `/api/orcamento/search` | **pública** | Busca na base |
| POST | `/api/orcamento/detalhes` | **pública** | Descrição/MDO por código |
| POST | `/api/orcamento/upload` | JWT + admin | Substitui a base via CSV |
| POST | `/api/orcamento/salvar` | JWT + admin | Substitui a base do usuário |
| GET/POST | `/api/regras` | JWT | Regras de aprendizado da IA |
| GET/POST | `/api/regras/conversao` | JWT | Regras de conversão por projeto |
| GET/POST/DELETE | `/api/recs`, `/api/recs/{n}` | JWT | Histórico de RECs |
| POST/GET | `/api/rec/salvar`, `/api/rec/{n}` | JWT | Rotas alternativas de REC |
| GET/POST/DELETE | `/api/obras` | JWT | Obras do usuário |
| GET | `/api/projetos` | opcional | Lista mesclada |
| POST | `/api/projetos` | JWT + admin | Cadastra projeto |
| GET | `/api/gemini/models` | JWT | Modelos disponíveis |
| POST | `/api/gemini/chat` | JWT | Chat com contexto da tabela |
| GET/POST/PUT/DELETE | `/api/admin/*` | JWT + admin | Usuários, master, audit log |
| GET | `/api/health`, `/api/backup/export` | pública / JWT | Diagnóstico e backup |
| GET | `/api/update/check` | pública | Versão mais recente |
| POST | `/api/update/apply` | JWT | Auto-update (só desktop) |

¹ `/api/auth/refresh` não está em `PUBLIC_ROUTES`; como todo `/api/*` não listado exige JWT válido,
o comportamento efetivo desse endpoint precisa de verificação. Ver STATE.md.

---

## 6. Modos de operação (`APP_MODE`)

| | `desktop` | `server` |
|---|---|---|
| CORS | `["*"]` | `CORS_ORIGINS` (fallback `["*"]`) |
| Janela | pywebview, fallback navegador | headless (uvicorn `0.0.0.0:$PORT`) |
| `/extract-local` | permitido se `IS_FROZEN` | bloqueado |
| `connectivity_monitor` | iniciaria (função não é chamada) | não inicia |
| `/api/update/check` | consulta o servidor central | lê `configuracoes` |

⚠️ `config.py:84` — `APP_MODE = os.environ.get("APP_MODE", "desktop" if IS_FROZEN else "desktop")`:
os dois ramos do ternário são iguais, então o default é sempre `"desktop"`. O modo servidor só é
ativado definindo explicitamente `APP_MODE=server` no ambiente.

---

## 7. Frontend — páginas e estado compartilhado

| Página | Papel |
|---|---|
| `index.html` + `script.js` | Upload, tabela de extração, edição de entidade/operação/ativo, filtros estilo Excel, seleção de células, cópia, modal de RECs |
| `resumo.html` + `resumo.js` | Tabelas Cabos e Outros, undo/redo, autocomplete, chat de IA, modais geradores (Cabos, Postes e Estruturas, Ramais, Conexões), Tabela de Regras, Tabela Totalizadora |
| `resultado_orcamento.html` | Chama `/api/orcamento/calcular`, consolida por `operação|mdo|código`, permite edição manual, salva REC, exporta |
| `orcamento.html` | Visualiza e edita a base técnica com filtros por coluna |
| `admin.html` + `admin.js` | Painel administrativo |
| `login.html` + `login.js` | Login; grava `auth_token` no `localStorage` |
| `auth_fetch.js` | Wrapper de `fetch` que injeta o `Authorization: Bearer` |

### Estado em `localStorage` (contrato entre páginas)

| Chave | Escrita em | Leitura em |
|---|---|---|
| `processar_dados` | `script.js:836` | `resumo.js:57` |
| `orcamentoPayload` | `resumo.js:1885` | `resultado_orcamento.html:1155` |
| `projeto_selecionado`, `projeto_selecionado_codigo` | `resumo.html`, `resumo.js` | `resumo.js`, cálculo |
| `regras_orcamento_<projCode>` | `resumo.js:salvarRegras` | `resumo.js:carregarRegras` |
| `numero_obra` | `resumo.html`, `resultado_orcamento.html` | `resultado_orcamento.html` |
| `auth_token`, `user_id`, `user_email`, `is_admin` | `login.js` | `auth_fetch.js` e páginas |
| `ai_provider`, `gemini_api_key`, `gemini_model` | `resumo.js` | `resumo.js` |

⚠️ `localStorage` é o mecanismo de transporte entre telas. Qualquer mudança no formato de
`processar_dados` ou `orcamentoPayload` quebra a página seguinte — são **contratos**, trate-os como tal.

---

## 8. Detalhamento das regras principais (entrada → regra → saída)

> Exemplos construídos a partir do código citado. Onde o exemplo depende de dados reais da base
> técnica, está marcado.

### 8.1 Operação pela cor — RN-01
```
ENTRADA   texto="DT11/300", cor="#FF0000"
REGRA     computeRowLogic: cor #FF0000 → I
SAÍDA     { operacao: "I" }

ENTRADA   texto="DT11/300", cor="#808080"   (cinza: |R-G|<5, |G-B|<5, 20<R<230)
REGRA     computeRowLogic: isGray → R
SAÍDA     { operacao: "R" }
```

### 8.2 Linha viva pelo layer — RN-02
```
ENTRADA   texto="1-CFU", cor="#FF0000", layer="01_LV"
REGRA     cor vermelha → I ; layer 01_LV → prefixa "*"
SAÍDA     { operacao: "*I" }
```

### 8.3 Retensionamento — RN-02
```
ENTRADA   texto="TR-2", cor preta, layer="01_RETENS"
REGRA     preto + layer de retensionamento + /TR\s*-\s*([123])/
SAÍDA     { entidade: "TRAFO", ativo: "1-RTR2", operacao: "I" }

ENTRADA   texto="3-100A", cor preta, layer="01_RETENS_LV"
REGRA     preto + layer + /^([123])\s*-\s*100\s*A/
SAÍDA     { entidade: "CHAVE", ativo: "3-RCFU", operacao: "*I" }
```

### 8.4 Normalização de texto livre — `processAtivoFormula`
```
"INSTALAR 01 U3"        → "1-U3"
"ROÇADO DE 35 METROS"   → "35-ROCO"
"REC. CALÇADA (2)"      → "2-RECAL"
"CONC. BASE 3X"         → "3-BASE"
"AFASTADOR"             → "1-AF"
"TR-1-45kVA"            → "1-TR1"
"1-100A-3H"             → "1-CFU 1-EF3H"
```

### 8.5 Quantidade de ativos do cabo — RN-04
```
"CAA 2 ABC 35 m"  → fase="ABC" (3 chars)            → qtdAtivos = 3
"CAA 2 A 35 m"    → prefixo CAA2 + fase 1 char      → qtdAtivos = 2   (len+1)
"CA 4 AB 20 m"    → prefixo CA4                     → qtdAtivos = 3   (len+1)
"M1x1x16+16 3 40 m" → prefixo começa com M          → qtdAtivos = 1
"CAZ 9,5 ABC 10 m"  → prefixo CAZ                   → qtdAtivos = 1
"P50"  (standalone, linha seguinte "CAA 2 ABC 35 m") → herda fase "ABC" → qtdAtivos = 3
```

### 8.6 Motor de regras de conversão — RN-05
```
REGRA     origem=CABOS, op_de=(vazio), ativo_de=(vazio), acao=ADICAO,
          fator=1.05, arredondamento=NORMAL
ENTRADA   { origem:"CABOS", operacao:"I", ativo:"CAA2", qtd:100 }
APLICAÇÃO qtd × 1.05 = 105 → arredonda 2 casas
SAÍDA     mantém o original { ativo:"CAA2", qtd:100 }   (ADIÇÃO não remove)
          + gera             { ativo:"CAA2", qtd:105 }

REGRA     origem=OUTROS, ativo_de=TERRA1, acao=SUBST, ativo_para=TERRA3, fator=1
ENTRADA   { origem:"OUTROS", operacao:"I", ativo:"TERRA1", qtd:2 }
SAÍDA     apenas { ativo:"TERRA3", qtd:2 }   (SUBST remove o original)
```
Padrões aceitos em `ativo_de`/`op_de`/`origem`: vazio = qualquer; `CA%` = LIKE;
`CAA\d+` = regex (ancorada automaticamente com `^…$`).

### 8.7 Cascata de resolução na base técnica — RN-06 / RN-07
```
ENTRADA   ativo = "DT9/150", origem = "OUTROS", projeto = "RONDONIA"

PASSO 1   ATIVO exato "DT9/150"  → encontrado (linha do seed)
          componente = "9150"
PASSO 7   expande TODOS os códigos do componente "9150":
            1015  MAO-DE-OBRA  Poste Limpo …     fator_i=1.0  fator_r=1.0
            22758 MATERIAL     POSTE CONCRETO …  fator_i=1.0  fator_r=1.0
SAÍDA     2 linhas no orçamento, cada uma com total = qtd × fator
```
Não achando por `ATIVO`, tenta `COMPONENTE` exato → `DESC_ATIVO` exato → `DESC_ATIVO` parcial →
`CÓDIGO` direto (fatores forçados a `1.0`). Falhando tudo, o ativo entra em `nao_encontrados`
e a linha fica destacada em vermelho na Totalizadora.

### 8.8 Soma por código — RN-09
```
ENTRADA   3 itens resolvidos para o código 1015 (MAO-DE-OBRA):
            qtd=2 op=I fator_i=1.0
            qtd=1 op=I fator_i=1.0
            qtd=4 op=R fator_r=1.0
REGRA     chave (1015, "MAO-DE-OBRA"); I/*I → soma_i; demais → soma_r
SAÍDA     { operacao:"I", codigo:"1015", total:3.0 }
          { operacao:"R", codigo:"1015", total:4.0 }
```

### 8.9 Prioridade de projeto — RN-08
```
BASE      linha A: ativo=CFU projeto="RONDONIA"
          linha B: ativo=CFU projeto=""            (genérica)
          linha C: ativo=CFU projeto="PARAIBA/BAHIA"

req_projeto = "RONDONIA"  → usa apenas a linha A
req_projeto = "CEARA"     → nenhuma exata, nenhuma… usa a genérica B
req_projeto = "" (vazio)  → usa a genérica B
```

---

## 9. Arquitetura conceitual (CABOS/POSTES → CONTEXTO → UNIFICADA → CÁLCULO)

Correspondência entre o modelo conceitual e o que existe no código:

| Etapa conceitual | Implementação real | Situação |
|---|---|---|
| Entrada **CABOS** | `tableStates.cabos` (entidade `CABO`), origem `"CABOS"` | ✅ existe |
| Entrada **POSTES** | **não é uma origem separada.** `POSTE` é uma das 10 *entidades*, e postes entram pela tabela **OUTROS**. Existe um modal "Gerador de Códigos — POSTES E ESTRUTURAS" (`resumo.js:2085`) que **alimenta a tabela OUTROS** | ⚠️ diverge do modelo conceitual |
| **CONTEXTO / REGRAS** | duas camadas distintas: (a) classificação automática por cor/layer/texto (`computeRowLogic`); (b) Tabela de Regras de Conversão (motor `ADIÇÃO`/`SUBST` com fator, arredondamento e limites) | ✅ existe, em duas partes |
| **CAMADA UNIFICADA** | `rawItems[]` em `syncTotalizadora` → Tabela Totalizadora. Mantém `origem` (procedência) e `baseId` `TOT-n` (agrupamento/rastreabilidade) | ✅ existe, sem esse nome no código |
| **CÁLCULO** | `processar_calculo()` no backend | ✅ existe |
| **RESULTADO** | `{operacao, mdo, codigo, desc_codigo, filtro, total}` | ⚠️ ver abaixo |

Sobre os campos esperados no resultado final: o cálculo devolve `codigo`, `desc_codigo`, `mdo`,
`filtro`, `operacao` e `total`. **`ativo`, `desc_ativo`, `fator`, `fator_i` e `fator_r` não são
devolvidos** — eles são *insumos* do cálculo (colunas da base técnica) e se perdem na agregação por
`(codigo, mdo)`. Se o resultado final precisar carregar o ativo de origem, isso é uma mudança de
comportamento, não uma correção.

### Regras de ativos U1–U4, U23, U32, TR
`U1`, `U2`, `U3`, `U4`, `U23`, `U32` aparecem **apenas** em `prompt_rede_eletrica.txt` (lista de
códigos válidos para a IA) e, como dados, nas linhas da base técnica. **Não existe lógica de código
que trate esses ativos de forma especial** — eles são resolvidos pela cascata genérica da seção 8.7.
Sobre `TR`: o código trata `TR-1/TR-2/TR-3` → `1-TR1/1-TR2/1-TR3` e `RTR1/RTR2/RTR3`
(retensionamento). A notação `TR+1`, `TR+2`, `TR+3` **não existe no código**.

---

## 10. Dependências entre módulos

```
app.py
 ├── config ─────────────► (todos)
 ├── database ──────────► routers/*, services/sync_service, offline_queue
 ├── middleware.auth_middleware ──► routers/auth, orcamento, obras, recs, projetos, admin
 ├── websocket_manager ─► routers/upload, services/connectivity_monitor, realtime_sync
 └── routers/
      ├── upload ──────► services/pdf_service, dxf_service, websocket_manager
      ├── orcamento ───► services/orcamento_calc, sync_service
      ├── auth ────────► services/supabase_client, database, middleware
      ├── admin ───────► services/supabase_client, database
      ├── ai_chat ─────► database, config.PROMPT_PATH
      ├── update ──────► services/auto_updater
      └── health ──────► services/sync_service

services/sync_service ──► supabase_client, offline_queue, database
services/offline_queue ─► supabase_client, database
services/realtime_sync ─► supabase_client, database, websocket_manager
services/orcamento_calc ─► (nenhuma — função pura, só stdlib `re`)
```

`orcamento_calc.py` não importa nada do projeto: é testável isoladamente e é o melhor ponto
para cobertura de testes.

---

## 11. Duplicação conhecida de lógica

`static/script.js` e `static/resumo.js` contêm **cópias independentes** de:
- `isGray()` (`script.js:278`, `resumo.js:129`)
- `processAtivoFormula()` (`script.js:295`, `resumo.js:137`)
- a cascata de classificação de entidade (`script.js:updateRowLogic:360`,
  `resumo.js:computeRowLogic:181`, `resumo.js:autoClassifyEntidade:1006`)

Alterar uma regra de classificação exige alterar **todas** as cópias, senão a aba principal e a aba
de resumo passam a discordar. Ver RULES.md, Regra 4.
