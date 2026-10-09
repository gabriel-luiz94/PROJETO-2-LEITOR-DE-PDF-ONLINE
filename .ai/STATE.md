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
- [x] Classificação automática de entidade/operação/ativo por cor, layer e texto, via **motor de
      regras do leitor** orientado a dados (`static/regras_leitor_engine.js`,
      `RegrasLeitorEngine.processarEClassificar`) — substituiu a lógica antes fixa em
      `computeRowLogic`/`updateRowLogic`/`processAtivoFormula`/`autoClassifyEntidade`
      (TASK-006, 2026-09-25). Duas tabelas por projeto (Processamento, Classificação), persistidas
      via Supabase/SQLite. Botão "Regras do Leitor" visualiza as regras do projeto selecionado.
      Ver `.ai/tasks/TASK-006-25-09-2026.md`.
- [x] Edição das regras do leitor pela UI, restrita a usuários `admin` (TASK-007, 2026-09-29):
      formulário estruturado por regra (campos condicionais por modo/fase), reordenação por
      drag-and-drop (reescreve o campo `ordem`, que é o que o motor de fato usa para ordenar — não
      a posição no array/JSON), painel de teste que roda o motor no navegador contra o conjunto de
      regras em edição (ainda não salvas), validação de schema no backend antes de salvar (regex
      compila, campos obrigatórios por modo), histórico de versões com reversão. Rotas de escrita
      (`POST`) exigem role `admin`; leitura continua aberta a qualquer usuário autenticado. Ver
      `.ai/tasks/TASK-007-29-09-2026.md`.
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
- [x] Obras, RECs e regras de IA (aprendizado do chat) isolados por projeto (coluna/campo
      `projeto`/`projeto_codigo`) — TASK-005, 2026-09-25; regras de IA passaram a sincronizar com
      o Supabase pela primeira vez (antes só existiam no SQLite local de cada instância)
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
- [x] Validação de contrato das planilhas Cabos e Outros (TASK-011, 2026-09-30): `POST
      /api/validacao/planilhas` (formato, operação, quantidade, duplicidade e, sobre o payload de
      cálculo, ativo ausente da base). Só backend — sem botão/painel ainda (TASK-014). Catálogo em
      `.ai/VALIDACOES.md`.
- [x] Correção assistida por IA (TASK-015, 2026-09-30): propostas antes → depois com aceite por linha, aplicar
      como um passo de histórico; a IA nunca aplica nem altera a operação; proposta que ainda reprova na
      camada 1 é descartada com o motivo.
- [x] Validação na tela e revisão por IA (TASK-014, 2026-09-30): botão **Validar**, opção **Validar ao montar
      orçamento** (radio local, padrão Não), painel de achados (contrato, domínio, base técnica e IA) e diálogo
      "continuar mesmo assim?" quando há erro/aviso. IA opcional, com o prompt salvo do projeto; sua falha nunca
      esconde os achados determinísticos. Falta a correção assistida (TASK-015).
- [x] Regras de domínio da validação (TASK-013, 2026-09-30): 11 regras (CFU/SUPL, CFU-CFUR/EF, trafo/PR15,
      trafo/PR220 com dobro por duas descidas, trafo sem chave não leva elo, P50 total, poste de 10 m em MT,
      estrutura isolada, formato do poste, estrutura SI exige RA2), por projeto com fallback `DEFAULT`, **todas desligadas até o admin ligar**. Editor, histórico e
      teste do rascunho no painel admin; a rota `/api/validacao/planilhas` já as aplica. Sem botão/painel na tela
      de trabalho ainda (TASK-014). Regras novas da semente chegam a quem já tem regras salvas pelo botão "Adicionar regras novas da semente".
- [x] Prompts de validação editáveis pelo admin (TASK-012, 2026-09-30): dois prompts semeados
      (`validar-planilhas`, `corrigir-planilhas`), versão por projeto com fallback `DEFAULT`, histórico e
      reversão, editor no painel admin. Ainda **não são usados** por nenhuma chamada de IA (TASK-014).
- [x] Chave de IA padrão do sistema (TASK-009, 2026-09-30): sem chave digitada, o chat usa
      `GEMINI_API_KEY`/`GOOGLE_API_KEY` (servidor: variável de ambiente; desktop: `.env` ao lado do
      `.exe`). Precedência usuário > salva > padrão; rate-limit por usuário só com a chave padrão;
      `/api/health` expõe `ai_key_source`. Corrigido o bug em que o sentinela `SAVED_IN_BACKEND`
      impedia o uso da variável de ambiente.

### Interface
- [x] Tabelas Cabos e Outros com undo/redo (80 níveis), autocomplete e filtros estilo Excel
- [x] Seleção de células, cópia e colagem em bloco
- [x] Modais geradores: Cabos, Postes e Estruturas, Ramais, Conexões
- [x] Tabela Totalizadora editável com destaque de ativo não encontrado
- [x] Tela de resultado com consolidação, edição manual, salvamento de REC e exportação
- [x] Painel "Resumo da rede" (TASK-008) — 4 contagens em tempo real (postes, rede de média, rede
      de baixa e equipamentos instalando) calculadas a partir das tabelas Cabos/Outros já
      carregadas na tela, sem persistência em banco (`static/resumo.js:calcularResumoRede`)
- [x] Vínculo cabo↔estrutura/poste (TASK-032, 2026-10-03): modal "VINCULAR CABOS" organizado por
      cabo (1 estrutura se operação M/`*M`, 2 se instalando), regras de vinculação por projeto
      (`regras_vinculacao`: tipo de estrutura pelo primeiro dígito do ativo, qtd. de cabos exigida,
      compatibilidade LIVRE/MESMO_TIPO_FASE_OPERACAO — mesmo padrão aberto de `regras_conversao`,
      sem histórico/admin-gate), algoritmo opcional de vinculação automática por proximidade de
      coordenadas (`_x`/`_y`, já existentes no DXF e agora também extraídas do PDF em
      `pdf_service.py`), que nunca sobrescreve vínculo manual. Persiste junto da obra/REC (vínculos
      viajam dentro de `tableStates.cabos.data`/`outros.data`, preservados por `deepClone()` —
      correção necessária, pois havia duas funções `deepClone` com o mesmo nome no arquivo e a
      última sobrescrevia a primeira silenciosamente para todo o escopo). Ver
      `.ai/tasks/TASK-032-02-10-2026.md`.
- [x] Validação de vínculo ausente/incompatível (TASK-033, 2026-10-03): reaproveita o motor de
      validação já existente (`executarValidacao()`/modal `#modal-validacao`) em vez de um painel
      próprio — aviso não bloqueante quando uma estrutura não tem os cabos exigidos por
      `regras_vinculacao`, com uma exceção: vínculo zero é perdoado se existir, em qualquer lugar do
      projeto (não necessariamente vinculado àquela estrutura), algum cabo `M`/`*M` ou ativo
      `RETCA`. Ver `.ai/tasks/TASK-033-02-10-2026.md`.
- [x] Ativo composto de vínculo + Regras de Conversão (TASK-034, 2026-10-03): cada vínculo
      cabo↔estrutura gera um ativo composto `<prefixo_cabo>_<ativo_estrutura>` (ex.: `CAA2_N4`),
      existente só dentro do cálculo da Totalizadora — nunca grava nas tabelas Cabos/Outros.
      Consumível pelas Regras de Conversão (TASK-003) através de uma nova origem `VINCULO` (além de
      `CABOS`/`OUTROS`), com uma diferença de comportamento: composto sem regra correspondente nunca
      aparece na Totalizadora (diferente de `CABOS`/`OUTROS`, que sempre mantêm o item original).
      Ver `.ai/tasks/TASK-034-02-10-2026.md`.

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
- [ ] **Embeddings das regras de aprendizado** — a coluna `regras.embedding` existe e é criada por
      migração, mas é sempre gravada como `NULL` (`regras.py:28`). Não há busca semântica.
- [ ] **Botão "LINHA VIVA"** — presente em `resumo.html:592` com `title="Função futura"`,
      sem handler.
- [ ] **`scripts/release.py`** — o `push_to_cloud()` está com o corpo comentado; só imprime
      mensagens. A publicação da versão na nuvem não acontece de fato.

---

## Próximas tarefas conhecidas

Apenas o que está explicitamente marcado como pendente no próprio projeto:

- [ ] Conectar os dois workers de background restantes acima (fila offline, sync em tempo real) ou decidir removê-los
- [ ] Implementar o botão "LINHA VIVA" (marcado como "Função futura" na UI)
- [ ] Completar `scripts/release.py:push_to_cloud` (código de referência já está comentado no arquivo)
- [ ] Adicionar `Dockerfile` e `fly.toml`, exigidos pelo workflow de deploy do backend
- [ ] Atualizar `.ai/CONTEXT.md` §7 (RN-01 a RN-05) e `.ai/ARCHITECTURE.md` para descrever o motor
      de regras do leitor (TASK-006) em vez da lógica fixa em código, agora obsoleta nesses
      documentos
- [x] Sistema de validação das planilhas Cabos/Outros em camadas (ADR-004): TASK-009 a TASK-015 concluídas.
      Todas as regras do catálogo estão implementadas (desligadas na semente). Catálogo em `.ai/VALIDACOES.md`.
- [x] Melhorias das regras de domínio pedidas pelo usuário (2026-09-30): **TASK-016** (motor v2: condição por
      quantidade, grupos, E/OU/NÃO, explicação), **TASK-018** (regras em camadas por projeto: padrão + ajustes, selos,
      −/+) e **TASK-017** (editor visual legível, grupos, assistente, teste explicativo) — concluídas.
- [x] **Validação + Ajuste** (pedido de 2026-10-01; TASK-019 a 025 concluídas), propostas em PLANEJAMENTO aguardando decisões do usuário, ordem sugerida:
      ~~TASK-021 (undo/redo)~~ ✔ → ~~TASK-020 (modos)~~ ✔ (modos determinística/IA) → TASK-019 (painel de regras no Resumo) →
      ~~TASK-022 (motor de ajustes)~~ ✔ → ~~TASK-023 (cadastro de ajustes)~~ ✔ → ~~TASK-024 (fluxo Validar+Ajustar)~~ ✔ → ~~TASK-025 (IA estruturada)~~ ✔.
- [x] **TASK-026** (pedido de 2026-10-01; concluída): embutir na gaveta lateral os controles de modo e "Validar ao montar", deixando na
      barra só `Regras de validação`, `Validar` e `Ajustar ▾`. Em PLANEJAMENTO, com decisões a confirmar.
- [x] **TASK-027** (pedido de 2026-10-01; concluída no navegador, falta a verificação humana no app de desktop): botão de microfone no chat de IA do Resumo (ditado por voz do navegador, conferir antes
      de enviar; leitura da resposta em voz alta opcional e desligada). Em PLANEJAMENTO, com 3 confirmações pendentes; risco
      principal: suporte à Web Speech API na janela desktop (pywebview).
- [x] **TASK-028** (pedido de 2026-10-01; concluída): IA do chat consulta as obras salvas e executa "adicione a obra 2 ao projeto" etc. (contexto
      sob demanda, ação `acao_ui: obra` com popup de confirmação, sem regressão e sem lentidão). CONCLUÍDA (`data/manual_regras_e_ajustes.md`, `/manual`, `GET /api/manual`).
- [x] **TASK-029** (pedido de 2026-10-02): o botão `Ajustar` roda todos os ajustes habilitados e lista cada correção com **Executar correção**,
      com **Executar todas as correções** acima. CONCLUÍDA (preview-lote + modal de cartões).
- [x] **TASK-030** (pedido de 2026-10-02): manual de uso para cadastrar regras de validação e ajustes. CONCLUÍDA (`data/manual_regras_e_ajustes.md`, `/manual`, `GET /api/manual`).
- [x] **TASK-031** (pedido de 2026-10-02): modo autônomo (ler → processar → ajustar → orçamento → salvar em pasta e banco, sem IA). Em PLANEJAMENTO, decisões confirmadas; **Fases A, B, C e D concluídas** (leitor via QuickJS, montagem, aplicador de diff, Totalizadora, pipeline com confirmação de exclusões, pasta monitorada + API de controle admin); tela `/autonomo` entregue; falta verificar no .exe/Windows e com arquivos reais; 4 fases (A leitor via QuickJS + montagem em Python, B pipeline, C pasta monitorada, D tela).
- [x] **TASK-035** (pedido de 2026-10-03): filtro por projeto no histórico da tela `/autonomo`, combinável com o
      filtro de status já existente (`GET /api/autonomo/execucoes?status=&projeto=`) — viabiliza conferir em lote
      vários projetos rodados pelo modo autônomo de uma vez. CONCLUÍDA.
- [x] **TASK-036** (pedido de 2026-10-03): incorpora o vínculo cabo↔estrutura/poste (TASK-032/033/034) ao modo
      autônomo — nova etapa `vincular` no pipeline, módulo `services/autonomo/vinculacao.py` (porte fiel do bloco
      de `static/resumo.js`), `montar_totalizadora()` reconhecendo a origem `VINCULO`. Decisão do usuário:
      auto-link sem combinação válida só gera achado de aviso, nunca bloqueia. **Achou e corrigiu 3 bugs
      pré-existentes** (confirmados com o leitor real): `_tipoEstruturaVinculo` e o ativo composto pegavam o
      dígito/prefixo de quantidade do ativo de Outros em vez do código da estrutura; `services/dxf_service.py`
      descartava `_x`/`_y` antes de devolver (nenhum DXF tinha coordenada disponível para o vínculo, só PDF).
      CONCLUÍDA. Ver `.ai/tasks/TASK-036-03-10-2026.md`.
- [x] **TASK-037** (pedido de 2026-10-03): `static/script.js` (botão "Processar" da tela Leitor) agora repassa
      `_x`/`_y` para `localStorage['processar_dados']` — fecha o último dos 4 bugs que impediam o vínculo
      automático por coordenada (TASK-032) de funcionar na tela manual. Confirmado fim a fim com DXF real, sem
      atalho sintético. CONCLUÍDA. Ver `.ai/tasks/TASK-037-03-10-2026.md`.
- [x] **TASK-038** (pedido de 2026-10-05): `adicionar_ativo` (motor de ajustes, TASK-022) aceita `qtd` negativo,
      gerando o prefixo `*` (ex.: `*1-PR`) em vez de rejeitar ou produzir um `-` literal que o cálculo de Outros
      não lê como sinal. `_num`/`_juntar`/`_fmt_qtd` são privados de `services/ajustes_planilhas.py`, sem efeito
      em outro módulo. CONCLUÍDA. Ver `.ai/tasks/TASK-038-05-10-2026.md`.
- [x] **TASK-039/040/041** (pedido de 2026-10-05): usuário relatou "a conexão com o Supabase está se perdendo
      após um tempo" e pediu garantia de disponibilidade de dados. Causa raiz confirmada: `services/supabase_client.py`
      era um singleton nunca recriado (`reset_supabase_client()` existia mas não era chamado em lugar nenhum), e
      toda falha de rede era engolida (`except Exception: pass`, sem log) em
      `obras.py`/`recs.py`/`projetos.py`/`admin.py` — uma vez que o socket caía (comportamento normal do
      Supabase/PostgREST após ociosidade), o processo nunca mais reconectava sozinho até reiniciar.
      **TASK-039** (reconexão reativa): `get_supabase()` ganhou TTL proativo de 180s + reset reativo via
      `registrar_falha()` quando a exceção é classificada como erro de rede. **TASK-040** (monitor de
      conectividade): `services/connectivity_monitor.py` corrigido (asyncio via `run_coroutine_threadsafe`) e
      ligado no startup de `app.py`, em desktop **e** servidor; transição offline→online força reset do
      cliente; novo indicador visual em `static/resumo.html`/`resumo.js`. **TASK-041** (visibilidade de falha):
      os 16 pontos silenciosos agora chamam `registrar_falha()` — decisão tomada de manter "sempre grava local,
      sempre responde sucesso", só adicionando o log (opção que não muda a experiência do usuário). CONCLUÍDAS.
      Ver `.ai/tasks/TASK-039-05-10-2026.md`, `TASK-040-05-10-2026.md`, `TASK-041-05-10-2026.md`.
- [x] **TASK-042** (pedido de 2026-10-06): usuário pediu para tornar as `regras_vinculacao`
      editáveis numa UI dentro do modal de vinculação, e para acessar o ativo composto
      (cabo+estrutura vinculados, ex. `CAA2_U4`) a partir de regras com formato amigável, também
      dentro do modal. Decisão: não estender o motor de Ajustes (roda antes do vínculo existir) —
      em vez disso, duas seções accordion novas dentro de `#modal-vinculacao`: (1) editor de
      `regras_vinculacao` sobre o endpoint já existente; (2) formulário amigável (tipo de cabo +
      tipo de estrutura → ativo) sobre a MESMA tabela de Regras de Conversão da Totalizadora
      (`origem: VINCULO`, já existente desde TASK-011/032/034), traduzindo `ativo_de` automático.
      Nenhum endpoint/tabela novo. CONCLUÍDA. Ver `.ai/tasks/TASK-042-06-10-2026.md`.
- [x] **TASK-043** (pedido de 2026-10-06): usuário perguntou se o programa lia `.dwg` — não
      diretamente (`ezdxf` só lê `.dxf`) — e pediu para o próprio programa converter internamente
      usando uma ferramenta online. Decisões do usuário: serviço **CloudConvert** (API v2, chave
      simples) e credencial no mesmo padrão da chave do Gemini (env padrão + chave do usuário,
      salva no servidor só após uso com sucesso). Novo `services/cloudconvert_service.py`
      (`converter_dwg_para_dxf`, sempre com I/O real — mockado em todos os testes); `.dwg` agora
      aceito em `POST /upload`, `GET /extract-local` e na pasta monitorada do modo autônomo,
      convertido para `.dxf` antes de seguir pelo mesmo caminho de extração. `.dxf`/`.pdf`
      continuam 100% offline. Primeira funcionalidade de LEITURA de arquivo do projeto a depender
      de internet (só para `.dwg`) — aceito como exceção pontual, mesmo padrão já aplicado à IA e
      ao Supabase. CONCLUÍDA. Ver `.ai/tasks/TASK-043-06-10-2026.md`.
- [x] **TASK-044** (pedido de 2026-10-06): usuário pediu para cadastrar "contextos" dentro de um
      mesmo projeto (ex.: "obras de 34,5kV"), cada um com seu próprio conjunto de ajustes ligados/
      desligados/sobrescritos, selecionáveis por um seletor na linha de totais (postes/cabos/
      equipamentos). Desenho: terceira camada de overlay (`padrão → projeto → contexto`),
      reaproveitando `services/ajustes_camadas.py::efetivo()` em CADEIA (chamada 2x), sem alterar a
      função. Tabelas novas `contextos` (metadados, pensada para um dia servir Regras de Domínio
      também) e `ajustes_contextos`/histórico (overlay por contexto). Toda rota de Ajustes ganhou
      `contexto` opcional — sem ele, comportamento idêntico a antes (testado). Escopo desta rodada:
      só Ajustes; modo autônomo e Regras de Domínio não usam contexto ainda. Histórico/reverter por
      contexto ficaram fora desta rodada (botões desabilitados na gaveta enquanto editando um
      contexto, para não arriscar reverter a camada errada). CONCLUÍDA. Ver
      `.ai/tasks/TASK-044-06-10-2026.md`.
- [x] **TASK-045** (pedido de 2026-10-06): dentro de um contexto (TASK-044), usuário pediu uma regra
      de ajuste "se tem U3 adicione em mesma quantidade 90277 negativo" — confirmado que U3 varia
      de quantidade (não é sempre 1), então um `qtd` fixo não resolvia. `adicionar_ativo.qtd` passou
      a aceitar `{"soma": SELETOR, "fator"?: número}` — soma dinâmica das quantidades de outro(s)
      ativo(s) da MESMA linha, reaproveitando `_contexto()`/`_casa()` já existentes (nenhuma mudança
      em `regras_dominio.py`). `fator: -1` gera o prefixo `"*"` automaticamente. UI da gaveta
      ganhou o seletor "quantidade: fixa/dinâmica". Achado colateral não corrigido: guard de
      contrato pré-existente trata mal código hifenado + token negativo (item 31 acima). CONCLUÍDA.
      Ver `.ai/tasks/TASK-045-06-10-2026.md`.
- [x] **TASK-046** (pedido de 2026-10-06): logo após a TASK-045, usuário pediu "caso tenha um trafo
      mono TR1* adicione com a mesma quantidade negativa o texto TR1* acrescido com VP no fim desse
      texto" — o nome do ativo adicionado precisa ser o código especificamente encontrado na linha
      (ex.: `TR110`), não um literal fixo. `adicionar_ativo.ativo` passou a aceitar também
      `{"igual_a": SELETOR, "prefixo"?: texto, "sufixo"?: texto}` — um token novo por item da
      própria linha que casa com o seletor, nomeado `<prefixo><código><sufixo>`, com quantidade =
      quantidade daquele item × `qtd` (fator). Reaproveita `_contexto()`/`_casa()` da TASK-045,
      nenhuma mudança em `regras_dominio.py`. Guard de idempotência adicionado: item cujo nome já
      carrega o prefixo/sufixo configurado é excluído dos candidatos (evita `"TR110VPVPVP..."` ao
      reaplicar sobre um curinga amplo como `"TR1*"`). UI da gaveta ganhou o seletor "ativo:
      fixo/dinâmico". CONCLUÍDA. Ver `.ai/tasks/TASK-046-06-10-2026.md`.
- [x] **TASK-047** (pedido de 2026-10-06): logo após a TASK-046, usuário perguntou se dava pra
      adicionar uma lista de ativos fixos (ex.: `90525, 90542, 92540`) numa única regra — já era
      possível com várias ações `adicionar_ativo` no mesmo ajuste, mas o usuário queria algo mais
      compacto. Confirmado "quantidade igual pra todos" antes de implementar. `adicionar_ativo.ativo`
      passou a aceitar também uma LISTA de códigos fixos — um token por código, todos com a mesma
      `qtd` (fixa ou dinâmica via `{"soma": ...}` da TASK-045, resolvida uma única vez). Extraída
      `_validar_qtd_adicionar()` pra eliminar duplicação entre o caso de ativo único e o de lista. UI
      da gaveta ganhou a opção "ativo: lista de códigos". CONCLUÍDA. Ver
      `.ai/tasks/TASK-047-06-10-2026.md`.
- [x] **TASK-048** (pedido de 2026-10-06): logo após a TASK-047, usuário perguntou se a lista aceita
      mais de um fator — confirmado o pedido real: "quantidades diferentes por código pra diminuir o
      número de regras". Um item da lista de `ativo` passou a aceitar `{"ativo": código, "qtd"?:
      número}` — quantidade PRÓPRIA, substituindo a `qtd` compartilhada só para aquele código; itens
      string continuam usando a compartilhada, os dois formatos podem ser misturados na mesma lista.
      UI: campo "códigos" ganhou a sintaxe `código:qtd` por item, sem seletor de modo novo.
      CONCLUÍDA. Ver `.ai/tasks/TASK-048-06-10-2026.md`.
- [x] **TASK-049** (pedido de 2026-10-06): usuário tentou combinar `código:qtd` (TASK-048) com
      "mesma quantidade de outro ativo" esperando "duas vezes a quantidade de U4, um código negativo
      e outro positivo" — não funcionou, pois `qtd` própria substitui a compartilhada por um valor
      fixo. Confirmado "não podemos adicionar a lógica do fator?". Um item da lista passou a aceitar
      também `{"ativo": código, "fator"?: número}` — MULTIPLICA a mesma base da `qtd` compartilhada
      (a soma, se dinâmica) em vez de substituí-la; `"qtd"` e `"fator"` no mesmo item são mutuamente
      exclusivos. Nova `_qtd_base()` extrai a magnitude não escalada, reaproveitada por
      `_qtd_dinamica()` (TASK-045, refatorada, comportamento idêntico). UI: sintaxe `código:xN` no
      mesmo campo "códigos". CONCLUÍDA. Ver `.ai/tasks/TASK-049-06-10-2026.md`.
- [x] **TASK-050** (pedido de 2026-10-07): usuário relatou que quantidade decimal com vírgula em
      Outros (ex.: `"0,7-M335"`) não funciona. Investigado e confirmado: dos quatro parsers
      independentes de Outros no projeto, só o da Totalizadora (`syncTotalizadora` em
      `static/resumo.js` e seu porte fiel `services/autonomo/totalizadora.py`) não aceitava vírgula
      — regex só reconhecia `.`, e `parseFloat`/`js_parse_float` também não leem vírgula sem
      `.replace(',', '.')`. Corrigido nos dois lugares em paralelo, mantendo a paridade testada
      contra o JS real. CONCLUÍDA. Ver `.ai/tasks/TASK-050-07-10-2026.md`.
- [x] **TASK-051** (pedido de 2026-10-07/08): usuário pediu que erros na tabela Outros fossem
      sinalizados pra evitar erros em grandes obras; depois pediu que a sinalização disparasse ao
      clicar em Ajustar, Validar ou Montar Orçamento, bloqueando mas permitindo "continuar mesmo
      assim", sempre ativa e automática. Achado: **Validar** já cobria os dois sinais (contrato +
      ativo não encontrado) por padrão, sem mudança; faltava ligar a mesma checagem em **Ajustar**
      (não validava nada) e tornar obrigatória em **Montar Orçamento** (só validava com a
      preferência opcional ligada). `executarValidacao` ganhou o parâmetro `somenteContrato`, que
      força Camada 1 + ativo não encontrado ignorando a preferência salva, sem duplicar nenhuma
      lógica; nova `checarInconsistenciasBasicas()` reaproveita o mesmo `cicloValidacao(res, true)`
      que o Montar Orçamento opcional já usava. CONCLUÍDA. Ver `.ai/tasks/TASK-051-07-10-2026.md`.
- [x] **TASK-052** (pedido de 2026-10-08): usuário pediu análise e estratégia (depois autorizou
      implementar) pra editar a base técnica (`tabela_orcamento_master`) escopada por projeto —
      hoje uma tabela única com todos os projetos misturados, campo Projeto como texto livre.
      Análise sobre o CSV real de produção (2.646 linhas, sem acesso de rede ao Supabase neste
      sandbox) revelou que `projeto` já suporta mais de um valor por linha via `/` (ex.
      `"PARAIBA/PARAIBANOVO"`, 108 linhas reais), lido por `processar_calculo`. Pedido final do
      usuário: "projetos totalmente manipuláveis de forma separada... editando inclusive as linhas
      compartilhadas". Resolvido com migração única (dado, sem alteração de schema): toda linha
      com `/` vira uma cópia exclusiva por projeto, valores idênticos no momento da migração, local
      e Supabase. `GET /api/orcamento/dados` ganhou `projeto`/`genericas`/`nao_reconhecido`
      opcionais; `upload-master`/`sync-master-all`/`salvar` (achado durante a implementação: o
      mesmo risco de DELETE sem filtro de projeto também existia na cópia pessoal do usuário, não
      só na master) ganharam escopo por projeto no DELETE, com validação das linhas enviadas.
      `static/orcamento.html` ganhou seletor de projeto + abas (Projeto selecionado/Comuns/Não
      reconhecido), campo Projeto travado. CONCLUÍDA. Ver `.ai/tasks/TASK-052-08-10-2026.md`.
- [x] **TASK-053** (pedido de 2026-10-08): usuário pediu pra ajustar a regra "AJUSTE DO (2)" (Cabos:
      "(2)" depois da bitola = "tem neutro, bitola 2") pra, quando a bitola da linha for diferente
      da do neutro, criar uma linha nova com o neutro usando o MESMO comprimento capturado na linha
      original. Investigação: juntar N na fase e limpar o "(2)" sem criar linha já eram possíveis
      com `substituir` + regex/backreferences existentes; criar linha com conteúdo dinâmico não
      era — a única ação que insere linha (`adicionar_linha`) usa texto fixo, sem regex, sem
      condição. Nova ação `criar_linha_derivada` (motor de ajustes): pra cada linha que casa com um
      regex, insere uma linha nova logo depois/antes dela, com o texto resolvido via
      `match.expand()` — backreferences do match DAQUELA linha. Reaproveita `_novo_erro_c1`/
      `_alvos`/`_quando` já existentes; itera um snapshot pra nunca reavaliar a linha recém-criada
      na mesma execução. Editor da gaveta e manual atualizados. CONCLUÍDA. Ver
      `.ai/tasks/TASK-053-08-10-2026.md`.
- [x] **TASK-054** (pedido de 2026-10-08): usuário recebeu `[ERRO] 400 INVALID_ARGUMENT ...
      'Multiturn chat is not enabled for this model'` no chat. Causa raiz: `routers/ai_chat.py`
      usava a API de sessão do Gemini (`chats.create`/`send_message_stream`), que alguns modelos
      recusam, e o fallback não reconhecia esse erro. Corrigido trocando para
      `client.aio.models.generate_content_stream` (chamada única em streaming, histórico embutido
      em `contents`, sem API de sessão) — elimina a causa raiz. Fallback ampliado com
      `invalid_argument`/`multiturn` (defesa em profundidade). Listas de modelo limpas
      (`gemini-1.5-flash`/`gemini-1.5-pro`/`gemini-2.5-flash` removidas das duas listas,
      `gemini-3.1-flash-lite` mantido como padrão). CONCLUÍDA. Ver
      `.ai/tasks/TASK-054-08-10-2026.md`.
- [x] **TASK-055** (pedido de 2026-10-08, logo após a TASK-054): usuário pediu suporte real a
      OpenAI e Claude como provedores de IA (hoje só Gemini funcionava), com o Claude Haiku
      (`claude-haiku-5-5`) como modelo padrão rápido/barato. `services/validacao_ia.py` ganhou
      `chamar_openai`/`chamar_claude` (mesmo contrato de `chamar_gemini`);
      `routers/validacao.py:_preparar_ia` ficou provider-aware (lê `ai_provider` salvo, chave/header
      próprios por provedor); `routers/ai_chat.py:gemini_chat` ganhou branches reais de streaming
      para os dois provedores novos. UI: seletor de provedor no modal único de config de IA
      (`static/resumo.html`/`resumo.js`), com o bug já identificado corrigido de passagem — o botão
      Salvar gravava `ai_provider: 'gemini'` fixo, agora grava o provedor escolhido. `anthropic`
      adicionado ao `requirements.txt`; `OPENAI_API_KEY`/`ANTHROPIC_API_KEY` documentadas no
      `.env.example`. Sem chave de teste real de OpenAI/Anthropic disponível na implementação —
      validado com fakes de SDK; usuário precisa confirmar com a própria chave. CONCLUÍDA. Ver
      `.ai/tasks/TASK-055-08-10-2026.md`.
- [x] **TASK-056** (pedido de 2026-10-08): exportar obra para arquivo `.obra.json` e importar por qualquer
      usuário (vira cópia particular de quem importa, só do projeto selecionado). CONCLUÍDA. Ver
      `.ai/tasks/TASK-056-08-10-2026.md`.
- [x] **TASK-057** (pedido de 2026-10-08): obras **públicas** (só visíveis no mesmo projeto, para leitura e
      cópia) ou **particulares** (padrão), com filtros Visibilidade/Origem e selos; a IA enxerga as do usuário e
      as públicas do projeto; obras do modo autônomo sempre particulares; obras antigas sem dono só para
      administradores (ação "Assumir"). **Ação do usuário:** rodar o SQL novo de `scripts/schema_supabase.sql`
      (colunas `publica`, `dono_nome`) no Supabase — até lá o programa funciona como antes. CONCLUÍDA. Ver
      `.ai/tasks/TASK-057-08-10-2026.md`.
- [x] **TASK-058** (pedido de 2026-10-08): **modelos** de obra (públicos no projeto, editáveis só pelo
      criador, filtráveis) com quantidades `V` (variável); carregar/adicionar/subtrair (e o comando da IA) abre
      a janela de parâmetros e gera uma obra padrão (não salva). **Ação do usuário:** rodar o SQL novo
      (`obras.tipo`) de `scripts/schema_supabase.sql`. CONCLUÍDA. Ver `.ai/tasks/TASK-058-08-10-2026.md`.
- [x] **TASK-059** (pedido de 2026-10-08): aprendizado supervisionado. **Concluída (etapas 1, 2 e 3):** registra as correções
      (autônomo **e trabalho manual**: botão Processar do Leitor + obras salvas), propõe **ajustes** (estatística + IA opcional
      após 10 obras) e **regras do leitor** (simuladas no motor real, sem regressões), aprovação sempre manual, só do próprio
      usuário; **nível de confiança** do autônomo por tipo de arquivo, com confirmação automática de exclusões opcional (desligada por padrão). Ver `.ai/tasks/TASK-059-08-10-2026.md`.
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

### IA (achados na TASK-009, não corrigidos)
26. A chave digitada pelo usuário é salva em `configuracoes` (global, sem `user_id`, texto puro): no
    servidor, a chave de um usuário passa a valer para os demais.
27. ~~`POST /api/gemini/chat` só implementa `provider="gemini"`; `openai` responde "não suportado"
    embora `GET /api/gemini/models` liste modelos da OpenAI.~~ **Resolvido na TASK-055
    (2026-10-08):** branches reais de streaming para `openai` e `claude`/`anthropic` (SDKs próprios,
    histórico e `system_instruction` aplicados do mesmo jeito que no Gemini); `_preparar_ia`
    (Camada 3) também ficou provider-aware. Sem chave de teste real de OpenAI/Anthropic disponível
    na implementação — validado com fakes de SDK; usuário precisa confirmar com a própria chave.
28. O rate-limit do chat é em memória e por processo.

29. ~~**Desfazer/refazer com desvio de uma posição**~~ **Resolvido na TASK-021 (2026-10-01):** `undo()` grava o estado
    ao vivo se ele estiver à frente do ponteiro antes de voltar, e `pushHistory()` ignora estado repetido; cada
    edição/lote é exatamente um passo (verificado no navegador, incluindo Limpar Tudo).

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
21. ~~Capturas de exceção silenciosas (`except Exception: pass`) em todos os fallbacks de Supabase
    (`obras.py`, `recs.py`, `projetos.py`, `admin.py`). Falhas de nuvem são invisíveis em log.~~
    **Corrigido — TASK-041 (2026-10-05):** os 16 pontos silenciosos agora chamam
    `services/supabase_client.registrar_falha(contexto, exc)`, que loga em `warning` sempre.
22. `orcamento.py:18` reimporta `APIRouter, UploadFile, File, Request` já importados na linha 7;
    `projetos.py:57,61` importa `HTTPException` duas vezes.
23. ~~`connectivity_monitor.py:39-56` manipula event loop do asyncio a partir de uma thread
    (`get_event_loop` / `new_event_loop` / `run_until_complete` / `loop.close()`) — padrão frágil.
    Sem efeito prático hoje, já que o worker nunca é iniciado.~~ **Corrigido — TASK-040
    (2026-10-05):** o loop principal é capturado uma vez no startup e usado com
    `asyncio.run_coroutine_threadsafe`; o monitor agora roda de fato (desktop e servidor).

### Dados
24. O repositório continha PDFs, DXFs e o `banco_resumo.db` com dados reais de obra, versionados
    antes das regras do `.gitignore`. **Removidos no commit `530d5cc`**; o histórico do Git ainda os
    contém. Mencionado aqui para que ninguém os re-adicione.
    De novo em 2026-09-30: `banco_resumo.db-shm` e `banco_resumo.db-wal` (SQLite em modo WAL, ~4 MB, dados
    recentes do banco do usuário) foram versionados por engano na `main` (commit `625cb13`, `.gitignore` só cobria
    `*.db` e `*.db-journal`). **Removidos do repositório e `*.db-shm`/`*.db-wal` adicionados ao `.gitignore`**; o
    histórico do Git ainda os contém.
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
30. `static/script.js` (botão "Processar" da tela Leitor) não repassava `_x`/`_y` do item extraído para
    `localStorage['processar_dados']` (achado na TASK-036, 2026-10-03) — o vínculo automático por coordenada
    (TASK-032) ficava não funcional na tela manual, mesmo com os outros bugs já corrigidos naquela tarefa.
    ✅ **Corrigido — TASK-037 (2026-10-03):** `_x`/`_y` incluídos condicionalmente no objeto exportado.
    Confirmado fim a fim com DXF real (servidor + Playwright, sem leitor falso nem dado sintético): extração →
    Processar → Resumo → vínculo automático funcionando. Ver `.ai/tasks/TASK-037-03-10-2026.md`.
31. **Não corrigido.** Achado na TASK-045 (2026-10-06): o guard de contrato `C1-OUT-NUM`
    (`services/validacao_planilhas.py`, usado pela camada 1 de `services/ajustes_planilhas.py`
    como guarda) trata mal uma linha de Outros que tem um código de ativo **hifenado** (ex.:
    `SUP-L`) junto de um token com prefixo negativo `"*"` (ex.: resultado de `adicionar_ativo` com
    `qtd` negativo) — descarta a mudança com um erro de "quantidade não numérica" que não reflete
    o problema real. Reproduz com `qtd` fixo negativo, não é algo introduzido pela TASK-045 (só
    descoberto ao testar quantidade dinâmica com `@GRUPO` contendo `SUP-L`). Sem correção ainda —
    fora do escopo do pedido que o achou.
32. ~~Quantidade decimal com vírgula em Outros não aparece certa na Totalizadora.~~
    ✅ **Corrigido — TASK-050 (2026-10-07):** dos quatro parsers independentes de Outros no
    projeto, três (`orcamento_calc.py`, `resumo.js::extrairParesQtdAtivoOutros`,
    `ajustes_planilhas.py`) já aceitavam vírgula; só o da Totalizadora não —
    `static/resumo.js:4601` (`syncTotalizadora`) e seu porte fiel
    `services/autonomo/totalizadora.py:19` (`_RE_OUTROS`) só reconheciam `.` como decimal, e
    `parseFloat`/`js_parse_float` também não liam vírgula sem `.replace(',', '.')` antes. Corrigido
    nos dois lugares em paralelo (regex `[.,]` + troca de vírgula por ponto antes de converter pra
    número), mantendo a paridade testada contra o JS real. Ver `.ai/tasks/TASK-050-07-10-2026.md`.

---

## Limitações atuais

- **Testes automatizados mínimos.** Só `tests/` (pytest): validação das planilhas e não-regressão do parser de `orcamento_calc.py`. Sem CI de teste e sem testes de frontend.
- **Validação das regras técnicas do domínio só sob demanda e só com regra ligada.** As regras de engenharia do
  `prompt_rede_eletrica.txt` (poste de 10 m em MT, CFU/SUPL, trafo/PR15/PR220, P50, estruturas isoladas, elo,
  SI/RA2) agora são verificáveis (ADR-004, TASK-013), mas vêm **desligadas** na semente e só rodam pelo botão
  Validar ou com "Validar ao montar" = Sim. Cada regra precisa ser ligada pelo admin depois de testada com dados
  reais; enquanto estiver desligada, o sistema **não** a aplica. O `prompt_rede_eletrica.txt` continua sendo
  apenas instrução ao modelo do chat.
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

**Status: focado na validação das planilhas (TASK-011 a TASK-013, 2026-09-30).**

- Suíte `pytest` em `tests/` (234 testes): `test_validacao_planilhas.py` e `test_rota_validacao.py` (camada 1,
  `C1-BASE`, paridade de tokenização, não-regressão de `processar_calculo`), `test_prompts_validacao.py`
  (TASK-012) `test_regras_dominio.py` (TASK-013) e `test_validacao_ia.py` (TASK-014) e `test_correcao_ia.py` (TASK-015), IA sempre simulada. Testes de rota usam banco temporário, nunca o de desenvolvimento.
- Dependências só de desenvolvimento: `requirements-dev.txt` (`pytest`, `httpx`); `pytest.ini` na raiz.
- CI de testes: nenhum (os dois workflows só fazem deploy). Frontend: nenhum teste.

Próximo passo natural, se o usuário quiser mais cobertura: as regras RN-03 a RN-10 de
`services/orcamento_calc.py` além do que a TASK-011 já toca.

---

## Última atualização

**Data:** 2026-10-08
**Motivo:** TASK-054/TASK-055 — usuário mandou `"substitua o texto S3 por S3T na tabela outros"` no
chat e recebeu `[ERRO] 400 INVALID_ARGUMENT ... 'Multiturn chat is not enabled for this model'`.
Investigado: o chat usava a API de sessão do Gemini (`chats.create`/`send_message_stream`), que
alguns modelos recusam, e o fallback não reconhecia esse erro pra tentar o próximo modelo. No mesmo
pedido, o usuário apontou modelos mortos na lista de fallback (`gemini-1.5-flash`/`gemini-1.5-pro`
desligados desde 2025; `gemini-2.5-flash` com desligamento anunciado para 16/10/2026) e, depois de
uma investigação de custo/arquitetura, pediu suporte real a OpenAI e Claude como provedores de IA
(hoje só Gemini funcionava), com o Claude Haiku (`claude-haiku-5-5`) como modelo padrão rápido/barato.
**TASK-054 — Alterações de código:** `routers/ai_chat.py` (`_stream_gemini` reescrita: troca de
`client.aio.chats.create`/`chat.send_message_stream` para `client.aio.models.generate_content_stream`
— chamada única em streaming com o histórico embutido manualmente em `contents`, eliminando a
dependência da API de sessão; gatilhos de fallback ampliados com `invalid_argument`/`multiturn`;
listas de modelo limpas — `gemini-1.5-flash`/`gemini-1.5-pro`/`gemini-2.5-flash` removidas, ficando
`["gemini-3.1-flash-lite", "gemini-3.6-flash", "gemini-3.5-flash"]`), `services/validacao_ia.py`
(`MODELOS_RESERVA` e `_ERROS_DE_FALLBACK` com a mesma limpeza/ampliação — `chamar_gemini` já usava
chamada única, sem o bug do multiturn, só tinha os modelos mortos na lista).
**TASK-055 — Alterações de código:** `services/validacao_ia.py` (`chamar_openai`/`chamar_claude`,
mesmo contrato de `chamar_gemini`; `MODELOS_RESERVA_OPENAI`/`MODELOS_RESERVA_CLAUDE`, Claude com
Haiku primeiro e Sonnet como reserva de qualidade); `routers/validacao.py` (`_preparar_ia` ganhou
dispatch por provedor — lê `ai_provider` salvo, resolve a chave certa por header/config próprios
`X-OpenAI-Key`/`openai_api_key` e `X-Anthropic-Key`/`anthropic_api_key`, e escolhe a função de
chamada pelo NOME via `globals()` pra preservar a testabilidade existente — "modelo" do cabeçalho do
prompt salvo só é usado pro Gemini, nunca pra OpenAI/Claude); `routers/ai_chat.py` (`resolver_credencial`
e `_chave_padrao` ganharam o parâmetro `provider`; `gemini_chat` ganhou branches reais de streaming
pra `openai` e `claude`/`anthropic`, com histórico e `system_instruction` aplicados do mesmo jeito
que no Gemini; `GET /models` ganhou listagem fixa para Claude e passou a usar `X-OpenAI-Key` próprio
em vez de reaproveitar a chave do Gemini); `models.py` (comentário do `ChatRequest.provider`
documentando o terceiro valor); `requirements.txt` (`anthropic` adicionado); `.env.example`
(`OPENAI_API_KEY`/`ANTHROPIC_API_KEY` documentadas); `static/resumo.html` (seletor de provedor com
3 botões, seções `#section-openai`/`#section-claude` novas, modelo padrão do Claude pré-selecionado);
`static/resumo.js` (`PROVIDER_CAMPOS`, troca de seção ao clicar no provedor, `sendToGemini` provider-
aware — chave/modelo/headers conforme o provedor salvo — e **bug corrigido**: o botão Salvar gravava
`ai_provider: 'gemini'` fixo, agora grava o provedor realmente escolhido).
Testes: `tests/test_ai_chat.py` (23, novo — chat completo: Gemini sem sessão, histórico no `contents`,
fallback em erro de multiturn/cota, listas sem modelos mortos, branches OpenAI/Claude, `/models` da
Claude, precedência de chave, UI estática do seletor e do bug corrigido) e 14 testes novos em
`tests/test_validacao_ia.py` (`chamar_openai`/`chamar_claude` unitários com fake de SDK, dispatch de
provedor em `_preparar_ia`, listas sem modelos mortos). Nenhum teste faz chamada de rede real —
SDKs (`google.genai`, `openai`, `anthropic`) sempre mockados. Suíte completa (843 testes) sem
regressão. Smoke test real-server + Playwright (uvicorn real numa thread, SDKs de IA mockados em
processo, banco SQLite temporário, Chromium via `executable_path` explícito — o download da revisão
esperada pelo Playwright foi bloqueado pelo proxy deste sandbox, usada uma revisão já presente no
ambiente): login real pela UI, chat respondendo à mensagem simples e à mensagem exata do bug
reportado COM histórico prévio (confirmado sem "[ERRO]"/"multiturn"), troca de provedor no modal
mostrando a seção certa a cada clique, modelo padrão do Claude pré-selecionado, `ai_provider` salvo
corretamente como "claude" (bug confirmado corrigido pelo navegador, não só por teste estático) e
provedor salvo restaurado ao reabrir o modal.
**Pendências explícitas para o usuário** (sem chave de API real de Gemini/OpenAI/Anthropic
disponível neste ambiente de implementação): (1) confirmar os IDs de modelo Gemini
(`gemini-3.1-flash-lite`/`gemini-3.6-flash`/`gemini-3.5-flash`) contra `GET /api/gemini/models` com
uma chave real; (2) confirmar que `claude-haiku-5-5`/`claude-sonnet-5-5` existem e respondem na
conta Anthropic do usuário (uma fonte de terceiros consultada nesta implementação dava o Haiku 5.5
como ainda não lançado em 08/10/2026 — tratado como ruído de pesquisa, não como motivo para desviar
do ID pedido explicitamente, mas vale confirmar); (3) os modelos padrão/reserva da OpenAI
(`gpt-4o-mini`/`gpt-4o`) foram uma escolha razoável desta implementação, não um pedido explícito do
usuário — vale confirmar ou trocar. Ver `.ai/tasks/TASK-054-08-10-2026.md`,
`.ai/tasks/TASK-055-08-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-053 — usuário trouxe o ajuste existente "AJUSTE DO (2)" (um `substituir` simples que
removia "(2)" do Cabos) e pediu pra ele passar a: quando a bitola do cabo bate com a bitola do
neutro "(2)", juntar o N na fase da mesma linha; quando a bitola diverge, tirar o "(2)" da linha
original e criar uma linha nova separada só com o neutro, reusando o comprimento capturado; quando a
fase já tem "N", só tirar o "(2)" sem duplicar. Investigação: 2 das 3 transformações já eram
alcançáveis com `substituir` + backreferences de regex; criar a linha nova dinâmica não era possível
com nada existente (`adicionar_linha` só aceita texto fixo, sem regex, insere uma vez só por
execução, sem suporte a `quando`) — gap explicado ao usuário, que autorizou implementar e mesclar.
**Alterações de código:** `services/ajustes_planilhas.py` — nova ação `criar_linha_derivada`
(`{"de": {"regex"}, "ativo_novo", "operacao_nova"?, "entidade_nova"?, "posicao"?, "quando"?,
"apenas_se_nao_existir"?}`): pra cada linha que casar com `de`, insere uma linha NOVA logo
depois/antes dela com `ativo_novo` resolvido via backreferences do match daquela linha (nunca de
outra); reusa `_novo_erro_c1` (mesmo guarda de contrato Camada 1 do `adicionar_linha` — linha
inválida é descartada, não inserida) e `_ops_filtro`/`_alvos`/`_quando` já existentes; itera um
snapshot pra nunca reavaliar a linha recém-criada na mesma execução (sem risco de crescimento em
cadeia). Campo `operacao_nova` reaproveitado de propósito (mesmo nome usado por outras ações) pra
ganhar validação genérica já existente sem código duplicado. Editor da gaveta
(`static/painel_ajustes.js`) e manual (`data/manual_regras_e_ajustes.md`) atualizados com a nova
ação, usando o exemplo real do "(2)" do usuário. Decisão: NÃO adicionada ao vocabulário do prompt de
correção por IA (`data/validacoes/ajustar-planilhas.md`) — fora do escopo deliberadamente limitado
desse prompt (máx. 5 ações, correções simples).
Testes: 18 novos em `tests/test_ajustes_planilhas.py` (ação isolada, `posicao`, filtro por `quando`,
`apenas_se_nao_existir`, descarte de linha inválida, não-reprocessamento da linha recém-criada,
`operacao_nova`/`entidade_nova`, validação de schema, `descrever_acao`, e um teste de integração com
a regra completa de 4 ações reproduzindo os três exemplos exatos do pedido). Suíte completa (806
testes) sem regressão. Smoke test real-server (uvicorn em processo, DB SQLite temporário via
`database.DB_PATH` patcheado, nunca `banco_resumo.db`) + Playwright: 3 linhas de Cabos com os
exemplos do usuário, editor da gaveta renderizado com os campos da nova ação, pré-visualização
confirmando os três resultados exatos ("CA 2 ABCN 35 m", "CA 1/0 ABC 35 m" + linha nova "CA 2 N 35
m", "CA 2 AN 35 m"). Ver `.ai/tasks/TASK-053-08-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-052 — usuário pediu análise e estratégia (depois autorizou implementar) pra editar
a base técnica (`tabela_orcamento_master`) escopada por projeto — hoje misturada numa tabela única,
campo Projeto como texto livre. Análise sobre o CSV real de produção (2.646 linhas) revelou que
`projeto` já suporta mais de um valor por linha via `/` (108 linhas reais `"PARAIBA/PARAIBANOVO"`).
Pedido final: "projetos totalmente manipuláveis de forma separada... editando inclusive as linhas
compartilhadas".
**Alterações de código:** `services/sync_service.py` — `migrar_projetos_compartilhados()` (divide
toda linha com `/` em cópias exclusivas por projeto, local e Supabase, cada um a partir da própria
consulta); `normalizar_projeto`/`validar_linhas_do_projeto`/`filtrar_linhas_por_categoria` (regra de
comparação de projeto extraída uma única vez). `routers/orcamento.py` — `GET /dados` ganhou
`projeto`/`genericas`/`nao_reconhecido` opcionais; `POST /salvar` ganhou `projeto` opcional que
escopa o DELETE (achado durante a implementação: tinha o mesmo risco sem filtro de projeto da
master, só que na cópia pessoal do usuário). `routers/admin.py` — `upload-master`/`sync-master-all`
ganharam `projeto` opcional (DELETE escopado por igualdade exata, local e Supabase, valida que toda
linha pertence ao projeto); novo `POST /admin/migrar-projetos-compartilhados`. `models.py` —
`SalvarOrcamentoRequest.projeto` opcional. `static/orcamento.html` — seletor de projeto + abas
(Projeto selecionado/Comuns a todos/Não reconhecido), campo Projeto travado (`disabled`) exceto na
aba de corrigir typo, botão de migração. `scripts/schema_supabase.sql` — índice opcional em
`projeto` (não obrigatório, sem mudança de schema).
Testes: 18 novos em `tests/test_orcamento_projeto.py` (migração, idempotência, filtro por
categoria, DELETE escopado em upload-master/sync-master-all/salvar, Supabase simulado). Suíte
completa (788 testes) sem regressão. Smoke test real-server + Playwright: populou as 3 categorias
+ uma linha compartilhada, confirmou a visão por projeto e o campo travado, rodou a migração ao
vivo (linha compartilhada virou duas cópias, cada uma só na visão do seu projeto) e confirmou que
importar um CSV escopado em PARAIBA não afeta RONDONIA nem a cópia de PARAIBANOVO. Ver
`.ai/tasks/TASK-052-08-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-051 — usuário pediu que erros na tabela Outros fossem sinalizados pra evitar erros
em grandes obras; depois refinou para disparar ao clicar em Ajustar, Validar ou Montar Orçamento,
bloqueando mas permitindo "continuar mesmo assim", sempre ativo e automático. Achado: Validar já
cobria os dois sinais (Camada 1 de contrato + ativo não encontrado) por padrão; faltava ligar a
mesma checagem em Ajustar (não validava nada) e tornar obrigatória em Montar Orçamento (só validava
com a preferência opcional ligada).
**Alterações de código:** `static/resumo.js` — `executarValidacao` ganhou o parâmetro opcional
`somenteContrato` (quando `true`, força `{det:true, dominio:false, ia:false, pularIa:true}` em vez
de `lerModos()`, resto da função inalterado); nova `checarInconsistenciasBasicas()` (chama
`executarValidacao(true)`, mostra o painel via `cicloValidacao(res, true)` só se houver achado de
erro/aviso, devolve se o usuário confirmou "continuar mesmo assim"); `btnAjustar` passou a chamar
essa checagem antes de abrir a gaveta; `btnMontarOrcamento` passou a chamá-la sempre, antes da
checagem opcional mais completa já existente (inalterada). `btnValidar` sem mudança.
Testes: suíte completa (770 testes) sem regressão — nenhuma mudança de backend. Smoke test
real-server (uvicorn em processo, DB SQLite temporário, nunca `banco_resumo.db`) + Playwright:
linha de Outros com quantidade não numérica bloqueou Ajustar e Montar Orçamento (mesmo com a
preferência desligada) mostrando a mensagem certa no painel; "continuar mesmo assim" seguiu
normalmente; linha com ativo não encontrado na base técnica também bloqueou; linha sem
inconsistência (ativo real da base técnica, sem pendência de vínculo) abriu a gaveta/montou o
orçamento direto, sem painel. Ver `.ai/tasks/TASK-051-07-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-050 — usuário relatou que quantidade decimal com
vírgula em Outros (ex.: `"0,7-M335"`) não funciona. Investigado: dos quatro parsers independentes
de Outros no projeto, só o da Totalizadora não aceitava vírgula.
**Alterações de código:** `static/resumo.js:4601,4609` (`syncTotalizadora`: regex `[.,]` em vez de
só `\.`, `.replace(',', '.')` antes do `parseFloat`); `services/autonomo/totalizadora.py:19,181`
(`_RE_OUTROS` com `[.,]`; `.replace(",", ".")` antes de `js_parse_float`, porte fiel — os dois
precisam mudar juntos pra manter a paridade testada contra o JS real).
Testes: 2 novos em `tests/test_autonomo_totalizadora.py` (decimal com vírgula, determinístico e
comparado contra o JS real via oráculo Node), mais um caso decimal adicionado ao conjunto de dados
do teste de fuzz existente. Suíte completa (770 testes) sem regressão. Smoke test real-server
(uvicorn em processo, DB SQLite temporário, nunca `banco_resumo.db`) + Playwright na tela de
verdade: digitou `"0,7-M335 *0,5-TR3 2-U4"` na célula Ativo da tabela Outros, confirmou via DOM que
a Totalizadora mostra `M335` qtd `0.7`, `TR3` qtd `-0.5`, `U4` qtd `2`. Ver
`.ai/tasks/TASK-050-07-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-049 — usuário tentou combinar `código:qtd` (TASK-048) com "mesma quantidade de
outro ativo" esperando "duas vezes a quantidade de U4, um código negativo e outro positivo"; não
funcionou, pois `qtd` própria substitui a compartilhada por um valor fixo. Confirmado "não podemos
adicionar a lógica do fator?".
**Alterações de código:** `services/ajustes_planilhas.py` (item da lista aceita também `{"ativo":
código, "fator"?: número}` — MULTIPLICA a mesma base da `qtd` compartilhada em vez de substituí-la;
`"qtd"` e `"fator"` no mesmo item são mutuamente exclusivos; nova `_qtd_base()` extrai a magnitude
não escalada, reaproveitada por `_qtd_dinamica()` [TASK-045, refatorada, comportamento idêntico];
`descrever_acao()` com frase para item com fator próprio); `static/painel_ajustes.js` (sintaxe
`código:xN` no mesmo campo "códigos", prefixo `x` distingue fator de quantidade fixa).
Testes: 6 novos em `tests/test_ajustes_planilhas.py` (fator próprio com base dinâmica — pedido real
do usuário —, fator com base literal, lista mista, base zero, validação de schema, frase em
português). Suíte completa sem regressão. Smoke test real-server (uvicorn em processo, DB SQLite
temporário via `database.DB_PATH` patcheado, nunca `banco_resumo.db`) + Playwright confirma o fluxo
completo pela UI (sintaxe `código:xN`, modelo em memória e frase ao vivo corretos). Ver
`.ai/tasks/TASK-049-06-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-048 — logo após a TASK-047, usuário perguntou se a
lista de códigos aceita mais de um fator; confirmado o pedido real: "quantidades diferentes por
código pra diminuir o número de regras".
**Alterações de código:** `services/ajustes_planilhas.py` (um item da lista de `ativo` aceita
`{"ativo": código, "qtd"?: número}` — quantidade PRÓPRIA, substituindo a `qtd` compartilhada só
para aquele código; itens string continuam usando a compartilhada, os dois formatos podem ser
misturados na mesma lista; `descrever_acao()` com frase para lista mista); `static/painel_ajustes.js`
(novos helpers `ajListaAtivoParaTexto`/`ajTextoParaListaAtivo`, sintaxe `código:qtd` no mesmo campo
"códigos", sem seletor de modo novo).
Testes: 7 novos em `tests/test_ajustes_planilhas.py` (quantidade própria por código — pedido real do
usuário —, lista mista com item simples, item dict sem `qtd`, interação com `se_ja_existe`,
validação de schema, frase em português). Suíte completa (762 testes) sem regressão. Smoke test
real-server (uvicorn em processo, DB SQLite temporário via `database.DB_PATH` patcheado, nunca
`banco_resumo.db`) + Playwright confirma o fluxo completo pela UI (sintaxe `código:qtd`, modelo em
memória e frase ao vivo corretos). Ver `.ai/tasks/TASK-048-06-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-047 — logo após a TASK-046, usuário perguntou se
dava pra adicionar uma lista de ativos fixos (ex.: `90525, 90542, 92540`) numa única regra. Já era
possível com várias ações `adicionar_ativo` no mesmo ajuste (o array `acoes` já suportava isso desde
a TASK-023), mas o usuário queria algo mais compacto. Confirmado "quantidade igual pra todos" antes
de implementar.
**Alterações de código:** `services/ajustes_planilhas.py` (`ativo` aceita também uma LISTA de
códigos fixos — um token por código, todos com a mesma `qtd`, resolvida uma única vez; nova
`_validar_qtd_adicionar()` extraída pra eliminar duplicação de validação entre o caso de ativo único
e o de lista; `_t_adicionar_ativo` com novo ramo para lista; `descrever_acao()` com frase em
português); `static/painel_ajustes.js` (seletor "ativo" ganhou a opção "lista de códigos", campo de
texto separado por vírgula via `ajCsv`; `qtd` reaproveita sem mudança o seletor fixa/dinâmica da
TASK-045).
Testes: 7 novos em `tests/test_ajustes_planilhas.py` (lista com quantidade fixa negativa — pedido
real do usuário —, sem qtd, com qtd dinâmica, soma dinâmica zero, interação com `se_ja_existe`,
validação de schema, frase em português). Suíte completa (755 testes) sem regressão. Smoke test
real-server (uvicorn em processo, DB SQLite temporário via `database.DB_PATH` patcheado, nunca
`banco_resumo.db`) + Playwright confirma o fluxo completo pela UI (alternar "ativo" para "lista de
códigos", preencher códigos separados por vírgula, modelo em memória e frase ao vivo corretos). Ver
`.ai/tasks/TASK-047-06-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-046 — usuário pediu "caso tenha um trafo mono TR1*
adicione com a mesma quantidade negativa o texto TR1* acrescido com VP no fim desse texto". O nome
do ativo adicionado precisa ser o código especificamente encontrado na linha (ex.: `TR110`), não um
texto fixo `"TR1*VP"` (curinga de busca não é um código válido).
**Alterações de código:** `services/ajustes_planilhas.py` (`ativo` aceita também `{"igual_a":
SELETOR, "prefixo"?: texto, "sufixo"?: texto}`; `_aplicar_um_ativo()` extraída de
`_t_adicionar_ativo` para reaproveitar em loop, um token por item casado; guard de idempotência
excluindo da busca itens que já carregam o prefixo/sufixo configurado; `descrever_acao()` com frase
em português para o caso dinâmico); `static/painel_ajustes.js` (seletor "ativo: fixo/dinâmico" com
campos "igual a"/"prefixo"/"sufixo"; com ativo dinâmico, `qtd` vira só o campo "fator").
Testes: 8 novos em `tests/test_ajustes_planilhas.py` (ativo dinâmico com quantidade negativa —
pedido real do usuário —, sem fator, com prefixo+sufixo, sem match, vários matches na mesma linha,
interação com `se_ja_existe` incluindo idempotência, validação de schema, frase em português).
Suíte completa (748 testes) sem regressão. Smoke test real-server (uvicorn em processo, DB SQLite
temporário via `database.DB_PATH` patcheado, nunca `banco_resumo.db`) + Playwright confirma o fluxo
completo pela UI (alternar "ativo" para dinâmico, preencher seletor/sufixo/fator, modelo em memória
e frase ao vivo corretos). Mesmo achado colateral da TASK-045 reencontrado (guard de contrato
pré-existente trata mal código hifenado + token negativo — item 31 em "Problemas conhecidos"), não
corrigido, fora do escopo. Ver `.ai/tasks/TASK-046-06-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-045 — dentro de um contexto (TASK-044), usuário
pediu uma regra de ajuste "se tem U3 adicione em mesma quantidade 90277 negativo". Confirmado com o
usuário que U3 varia de quantidade (`1-U3`, `2-U3`, `3-U3`...) antes de implementar, já que
`adicionar_ativo.qtd` só aceitava número fixo.
**Alterações de código:** `services/ajustes_planilhas.py` (`qtd` aceita
`{"soma": SELETOR, "fator"?: número}`; nova `_qtd_dinamica()`, reaproveitando `_contexto()`/`_casa()`
já existentes; `_t_adicionar_ativo` resolve a soma antes de aplicar, no-op se soma=0;
`descrever_acao()` com frase em português para o caso dinâmico); `static/painel_ajustes.js`
(seletor "quantidade: fixa/dinâmica" com campos "soma de"/"fator").
Testes: 8 novos em `tests/test_ajustes_planilhas.py` (quantidade dinâmica negativa — pedido real do
usuário —, sem fator, fator≠1, soma zero, soma de `@GRUPO`, interação com `se_ja_existe`, validação
de schema, frase em português). Suíte completa (740 testes) sem regressão. Smoke test real-server +
Playwright confirma o fluxo completo pela UI (trocar pra quantidade dinâmica, preencher seletor e
fator, modelo em memória correto). Achado colateral não corrigido (fora do escopo): guard de
contrato pré-existente trata mal código hifenado + token negativo — ver item 31 em "Problemas
conhecidos". Ver `.ai/tasks/TASK-045-06-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-044 — usuário pediu para cadastrar "contextos" dentro de um mesmo projeto (ex.:
"obras de 34,5kV"), cada um com seu próprio conjunto de ajustes ligados/desligados/sobrescritos,
selecionáveis por um seletor na linha de totais. Decisões do usuário: terceira camada de overlay
(poder completo, não só liga/desliga), desenho genérico (pensado para um dia servir Regras de
Domínio também), e trocar de contexto não recalcula nada na hora — só define o que roda na próxima
vez que clicar em "Ajustar".
**Alterações de código:** tabelas novas `contextos`/`ajustes_contextos`/
`ajustes_contextos_historico` (`database.py`, `scripts/schema_supabase.sql`); novo
`services/contextos.py` (CRUD de metadados); `services/repo_json.py` (nova `RepoJsonContexto`,
chave composta, sem alterar `RepoJson`); `routers/validacao_ajustes.py` (`resolver()` em cadeia —
chama `services/ajustes_camadas.py::efetivo()` duas vezes sem alterá-la —, rotas de CRUD de
contexto, campo `contexto` opcional em `obter`/`salvar`/`preview`/`preview-lote`); `static/resumo.html`
(seletor "Contexto" na linha de totais); `static/resumo.js` (`contextoSelecionado()`,
`carregarContextosAjustes`, criar/excluir contexto, persistência por projeto); `static/painel_ajustes.js`
(`ajContexto()`, editor da gaveta context-aware, Histórico/Restaurar desabilitados em modo contexto).
Escopo: só Ajustes; modo autônomo e histórico/reverter por contexto ficaram fora desta rodada.
Testes: `tests/test_ajustes_contextos.py` (11, novo) — CRUD, cadeia, independência entre contextos,
regressão byte a byte sem `contexto`, `preview-lote` respeitando o contexto; 2 mocks ajustados em
`tests/test_ajustes_lote.py` para a nova assinatura opcional. Suíte completa (732 testes) sem
regressão. Smoke test real-server + Playwright confirma o fluxo completo pela UI (criar contexto,
editar na gaveta, salvar, projeto permanece intacto). Ver `.ai/tasks/TASK-044-06-10-2026.md` e
`.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-043 — usuário perguntou se o programa conseguia ler `.dwg`; confirmado que não
diretamente (`ezdxf` só lê `.dxf`, formato é binário proprietário) e pediu para o próprio programa
converter internamente usando uma ferramenta online. Decisões do usuário: serviço CloudConvert
(API v2) e credencial no mesmo padrão da chave do Gemini (env padrão do sistema + chave do usuário,
persistida no servidor só depois de confirmada por um uso com sucesso).
**Alterações de código:** novo `services/cloudconvert_service.py` (`converter_dwg_para_dxf`, I/O
real via `httpx`, sempre mockado em teste); `routers/upload.py` (resolução de credencial —
header `X-CloudConvert-Key` > `configuracoes` > env — e dispatch de `.dwg` em `/upload` e
`/extract-local`); `routers/health.py` (`cloudconvert_key_source`); `services/autonomo/pipeline.py`
(`.dwg` em `EXTENSOES`, convertido antes de `ler_arquivo`); `static/index.html`/`script.js` (`.dwg`
aceito no input, botão "Chave DWG" + modal); `requirements.txt` (`httpx` como dependência direta).
Testes: `tests/test_cloudconvert_service.py` (9, novo), `tests/test_upload_dwg.py` (10, novo),
`tests/test_autonomo_pipeline.py` (+3) — `httpx` sempre mockado, nenhuma chamada de rede real;
`tests/conftest.py` passou a zerar `CLOUDCONVERT_API_KEY` também. Suíte completa (721 testes) sem
regressão. Smoke test real-server + Playwright confirma a UI nova (botão, modal, salvar, `accept`)
sem erros de JS. Ver `.ai/tasks/TASK-043-06-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-042 — pedido do usuário para tornar as
`regras_vinculacao` editáveis numa UI dentro
do modal de vinculação, e para acessar o ativo composto (cabo+estrutura vinculados, ex. `CAA2_U4`)
a partir de regras com formato amigável, também dentro do modal. Decisão: não estender o motor de
Ajustes (roda antes do vínculo/Totalizadora existirem) — em vez disso, dar UI aos dois mecanismos
já existentes no lugar certo.
**Alterações de código:** `static/resumo.html` (duas seções accordion novas dentro de
`#modal-vinculacao`: `#table-vinc-regras` e `#table-vinc-conectores`), `static/resumo.js`
(`toggleVincRegras`/`renderRegrasVinculacaoTable`/`updateRegraVinculacaoRow`/
`adicionarRegraVinculacao`/`excluirRegraVinculacao`/`salvarRegrasVinculacaoNuvem` sobre o endpoint
já existente `/api/regras-vinculacao`; `toggleVincConectores`/`renderRegrasConectoresTable`/
`updateRegraConectorRow`/`adicionarRegraConector`/`excluirRegraConector`/`_splitAtivoVinculo` sobre
a mesma tabela de Regras de Conversão da Totalizadora, `origem: VINCULO`). Nenhum endpoint/tabela
novo, nenhum arquivo de backend tocado.
Testes: smoke test real-server + Playwright (DB temporário, nunca o `banco_resumo.db`) — adicionar/
editar/salvar nas duas seções persiste corretamente (confirmado via `GET` de volta nos dois
endpoints), e uma linha inválida é rejeitada pelo backend com erro visível via toast, sem poluir o
array salvo. Ver `.ai/tasks/TASK-042-06-10-2026.md` e `.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-039/040/041 — pedido do usuário "garantir que os dados estejam sempre disponíveis.
E em caso de perda de conexão os dados se mantenham com a última versão carregada para trabalho",
motivado pelo relato "a conexão com o Supabase está se perdendo após um tempo". Causa raiz
confirmada: cliente Supabase era um singleton nunca recriado, e toda falha de rede era engolida
sem log em 4 routers.
**Alterações de código:** `services/supabase_client.py` (TTL proativo de 180s em `get_supabase()`,
nova `registrar_falha(contexto, exc)` com reset reativo classificado por tipo de exceção),
`services/connectivity_monitor.py` (asyncio corrigido via `run_coroutine_threadsafe`, reset do
cliente na transição offline→online, sem restrição de `APP_MODE`), `app.py` (novo startup hook
`_iniciar_monitor_conectividade`), `routers/obras.py`/`recs.py`/`projetos.py`/`admin.py` (16 pontos
`except Exception: pass` → `registrar_falha`), `static/resumo.html`/`resumo.js` (novo indicador
visual de conectividade), `tests/test_supabase_client.py` (novo, 14 testes).
Testes: `pytest tests/test_supabase_client.py` (14/14) + `pytest tests/` completo (699 testes, sem
regressão) + smoke test real-server (startup sem crash, `/api/health` normal) + verificação
Playwright do indicador (estado inicial oculto, sem erros de JS). Ver
`.ai/tasks/TASK-039-05-10-2026.md`, `TASK-040-05-10-2026.md`, `TASK-041-05-10-2026.md` e
`.ai/CHANGELOG.md`.

Entrada anterior (mantida para histórico): TASK-038 — pedido do usuário por um ajuste que, achando
TR110 (trafo mono) na linha, adiciona vários ativos, um deles negativo (`*1-PR`, "linha viva"). A
ação `adicionar_ativo` exigia `qtd > 0` e, mesmo aceitando, geraria um `-` literal que o cálculo de
Outros não lê como sinal. Alterações: `services/ajustes_planilhas.py` (`validar_acoes` aceita `qtd`
negativo; `_num` lê um token existente com `*`; nova `_fmt_qtd()` gera o prefixo `*` corretamente),
`data/manual_regras_e_ajustes.md`. Ver `.ai/tasks/TASK-038-05-10-2026.md`.

Entrada anterior (mantida para histórico): TASK-037 — última correção da cadeia de 4 bugs que
impediam o vínculo automático por coordenada (TASK-032) de funcionar de verdade (dígito do tipo,
ativo composto, DXF descartando coordenada, e o `static/script.js` não repassando `_x`/`_y` para
o Resumo). Com os 4 corrigidos, confirmado fim a fim com um DXF real. A funcionalidade de vínculo
(TASK-032/033/034/036/037) está completa e funcional nos dois modos (manual e autônomo). Ver
`.ai/tasks/TASK-037-03-10-2026.md`.

Entrada anterior (mantida para histórico): TASK-036 — incorpora o vínculo cabo↔estrutura/poste ao
modo autônomo; achou e corrigiu 3 dos 4 bugs acima. Alterações: `services/autonomo/montagem.py`,
novo `services/autonomo/vinculacao.py`, `services/autonomo/pipeline.py`,
`services/autonomo/totalizadora.py`, `services/dxf_service.py`, `static/resumo.js`. Ver
`.ai/tasks/TASK-036-03-10-2026.md`.

Entrada anterior (mantida para histórico): TASK-035 — filtro por projeto no histórico do modo
autônomo. Alterações: `services/autonomo/execucoes.py` (`listar()` ganhou o parâmetro `projeto`),
`routers/autonomo.py`, `static/autonomo.html`, `static/autonomo.js`. Ver
`.ai/tasks/TASK-035-03-10-2026.md`.

Entrada anterior (mantida para histórico): TASK-032/033/034 — vínculo cabo↔estrutura/poste, validação
do vínculo e ativo composto consumido pelas Regras de Conversão, a pedido do usuário ("Implemente. E
refine."). Renumeradas de TASK-009/010/011 por colisão com tarefas já mescladas em `main` por outra
sessão. Alterações: `services/pdf_service.py` (coordenadas `_x`/`_y`), `config.py`,
`data/regras_vinculacao_seed.json`, `database.py` (tabela `regras_vinculacao` + seed),
`routers/regras_vinculacao.py`, `app.py`, `scripts/schema_supabase.sql`, `static/resumo.html`,
`static/resumo.js`. Ver `.ai/tasks/TASK-032-02-10-2026.md`, `TASK-033-02-10-2026.md`,
`TASK-034-02-10-2026.md`.
