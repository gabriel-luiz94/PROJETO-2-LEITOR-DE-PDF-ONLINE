# CHANGELOG

> Histórico de alterações **relevantes** do projeto, do ponto de vista de contexto para IA:
> arquitetura, regras de negócio, fluxo de dados, estrutura de banco e comportamento importante.
>
> Não registre aqui alterações triviais (formatação, renomeação local, ajuste de texto).
> Formato: mais recente primeiro.

---

## 2026-10-08 — TASK-051: checagem obrigatória (contrato + ativo não encontrado) ao Ajustar/Montar Orçamento

**Tipo:** feature (validação) · `static/resumo.js` · `.ai/tasks/TASK-051-07-10-2026.md`

Usuário pediu que erros na tabela Outros fossem sinalizados pra evitar erros em grandes obras.
Depois de descartar a opção de unificar os 4 parsers de Outros (risco alto, muitos pontos de
regressão), desenhamos destaque ao vivo na tabela — mas o usuário mudou o modelo de disparo: "quero
que esse tipo de sinalização seja disparada quando clicar em ajustar, validar ou montar orçamento.
Exiba uma mensagem e não prossiga devido às inconsistências." Confirmado: bloqueia mas permite
"continuar mesmo assim" (mesmo padrão que Montar Orçamento já tinha, opcionalmente); roda sempre,
automaticamente, sem precisar ligar nenhuma preferência.

Investigação mostrou que **Validar** já cobria os dois sinais pedidos por padrão (Camada 1 de
contrato + `validar_base`/ativo não encontrado, via `MODOS_PADRAO.det = true`) — zero mudança
necessária ali. Faltava: (1) Ajustar não validava nada antes de abrir a gaveta; (2) Montar Orçamento
só validava se a preferência "Validar ao montar orçamento" (desligada por padrão) estivesse ligada.

Implementado com reaproveitamento total da lógica existente, sem duplicar nada: `executarValidacao`
ganhou um parâmetro opcional `somenteContrato` que, quando `true`, ignora a preferência salva do
usuário e força `{det:true, dominio:false, ia:false, pularIa:true}` — mesma montagem de payload,
mesma chamada a `/api/validacao/planilhas`, mesmo formato de retorno de sempre, só sem regras de
domínio nem IA. Nova função `checarInconsistenciasBasicas()` chama `executarValidacao(true)` e, se
houver achado de erro/aviso, usa o mesmo `cicloValidacao(res, true)` que o Montar Orçamento opcional
já usava pra mostrar o painel e perguntar "continuar mesmo assim" — nenhuma UI nova. `btnAjustar`
agora chama essa checagem antes de abrir a gaveta; `btnMontarOrcamento` passou a chamá-la sempre,
antes da checagem opcional mais completa (que continua intacta, pra quem já liga "Validar ao montar
orçamento" e quer domínio/IA também). `btnValidar` não mudou.

Suíte completa (770 testes) sem regressão — nenhuma mudança de backend. Smoke test real-server +
Playwright: linha de Outros com quantidade não numérica bloqueou Ajustar e Montar Orçamento (mesmo
com a preferência desligada) com a mensagem certa no painel; "continuar mesmo assim" seguiu
normalmente; linha com ativo não encontrado na base técnica também bloqueou; linha sem nenhuma
inconsistência abriu a gaveta/montou o orçamento direto, sem painel.

---

## 2026-10-07 — TASK-050: quantidade decimal com vírgula em Outros não aparecia na Totalizadora

**Tipo:** correção de bug · `static/resumo.js`, `services/autonomo/totalizadora.py` ·
`.ai/tasks/TASK-050-07-10-2026.md`

Usuário relatou: "na tabela outros não é possível utilizar uma quantidade quebrada como, por
exemplo 0,7-M335". Investigado: o projeto tem 4 implementações independentes do parser de Outros
(`<qtd>-<ativo>`). Três já aceitavam vírgula decimal (`orcamento_calc.py`,
`resumo.js::extrairParesQtdAtivoOutros`, o motor de Ajustes). A quarta — `syncTotalizadora`
(`static/resumo.js`, monta a Totalizadora/payload de cálculo) e seu porte fiel documentado no modo
autônomo (`services/autonomo/totalizadora.py`) — não: o regex só reconhecia `.` como separador
decimal, e mesmo reconhecendo, `parseFloat`/`js_parse_float` também não leem vírgula sem
`.replace(',', '.')` antes. Efeito: `"0,7-M335"` não casava o regex, caía no fallback que trata a
linha inteira como nome do ativo com quantidade 1 — sem erro visível, silenciosamente errado.

Corrigido nos dois lugares em paralelo (JS e seu porte Python, que precisam mudar juntos pra manter
a paridade testada em `tests/test_autonomo_totalizadora.py` contra o JS real via Node/QuickJS):
regex passou a aceitar `[.,]`, e o valor captado troca vírgula por ponto antes de virar número.
Mudança 100% aditiva — nenhum valor que já funcionava (inteiro, com ponto) muda de comportamento.

2 testes novos (`tests/test_autonomo_totalizadora.py`), mais um caso decimal adicionado ao conjunto
de dados do teste de fuzz existente (roda contra o JS real). Suíte completa (770 testes) sem
regressão. Smoke test real-server + Playwright na tela de verdade: digitou `"0,7-M335 *0,5-TR3
2-U4"` na célula Ativo da tabela Outros, confirmou via DOM que a Totalizadora mostra `M335` qtd
`0.7`, `TR3` qtd `-0.5`, `U4` qtd `2`.

---

## 2026-10-06 — TASK-049: `adicionar_ativo` — item da lista com fator próprio sobre a qtd dinâmica

**Tipo:** feature (motor de Ajustes) · `services/ajustes_planilhas.py`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-049-06-10-2026.md`

Usuário tentou combinar a sintaxe `código:qtd` (TASK-048) com "mesma quantidade de outro ativo"
(U4, fator 2), esperando "duas vezes a quantidade de U4, um código negativo e outro positivo" — não
funcionou, porque `qtd` própria SUBSTITUI a compartilhada por um valor literal fixo (comportamento
correto da TASK-048, mas não servia aqui). Perguntado "não podemos adicionar a lógica do fator?" —
confirmado.

Implementado: um item da lista agora também pode ser `{"ativo": código, "fator"?: número}` —
MULTIPLICA a mesma base da `qtd` compartilhada (a soma, se `qtd` for dinâmico; o próprio número, se
for literal) em vez de substituí-la por um valor fixo. `"qtd"` e `"fator"` no mesmo item são
mutuamente exclusivos. Nova `_qtd_base()` extrai a magnitude não escalada (sem fator), reaproveitada
tanto por `_qtd_dinamica()` (TASK-045, refatorada para usá-la, comportamento idêntico) quanto pelo
novo caso de item com fator próprio. UI: sintaxe `código:xN` no mesmo campo "códigos".

Exemplo exato do pedido: `{"acao": "adicionar_ativo", "tabela": "outros", "ativo": [{"ativo":
"90277", "fator": -2}, {"ativo": "90279", "fator": 2}], "qtd": {"soma": "U4"}, "quando": {"tem":
"U4"}}` — com `2-U4` na linha, gera `*4-90277` e `4-90279`.

6 testes novos (`tests/test_ajustes_planilhas.py`). Suíte completa sem regressão. Smoke test
real-server + Playwright confirma o fluxo completo pela UI (sintaxe `código:xN`, modelo em memória
e frase ao vivo corretos).

---

## 2026-10-06 — TASK-048: `adicionar_ativo` — item da lista com quantidade própria

**Tipo:** feature (motor de Ajustes) · `services/ajustes_planilhas.py`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-048-06-10-2026.md`

Logo após a TASK-047, usuário perguntou se a lista aceitava mais de um fator — respondido que, como
implementada, a quantidade era compartilhada por toda a lista. Usuário confirmou o pedido real:
"quantidades diferentes por código pra diminuir o número de regras".

Implementado: um item da lista de `ativo` agora pode ser `{"ativo": código, "qtd"?: número}` em vez
de uma string simples — tem a PRÓPRIA quantidade, substituindo a `qtd` compartilhada da ação só para
aquele código. Itens string continuam usando a `qtd` compartilhada (comportamento da TASK-047,
inalterado) — os dois formatos podem ser MISTURADOS na mesma lista. UI: o campo "códigos" ganhou a
sintaxe `código:qtd` (ex.: `90525:-1, 90542:-2, 92540`), sem precisar de um seletor de modo novo.

Exemplo exato do pedido: `{"acao": "adicionar_ativo", "tabela": "outros", "ativo": [{"ativo":
"90525", "qtd": -1}, {"ativo": "90542", "qtd": -2}]}`.

7 testes novos (`tests/test_ajustes_planilhas.py`). Suíte completa sem regressão. Smoke test
real-server + Playwright confirma o fluxo completo pela UI (sintaxe `código:qtd`, modelo em memória
e frase ao vivo corretos).

---

## 2026-10-06 — TASK-047: `adicionar_ativo` com lista de ativos fixos (mesma quantidade pra todos)

**Tipo:** feature (motor de Ajustes) · `services/ajustes_planilhas.py`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-047-06-10-2026.md`

Usuário perguntou se dava pra passar uma lista (ex.: `90525, 90542, 92540`) numa única regra. Já era
possível adicionar vários ativos no mesmo ajuste usando várias ações `adicionar_ativo` (o array
`acoes` já suportava isso desde a TASK-023), mas o usuário queria algo mais compacto. Confirmado:
"quantidade igual pra todos" antes de implementar.

Implementado: `ativo` passa a aceitar também uma LISTA de códigos fixos — um token novo por código,
todos com a MESMA `qtd` (fixa ou dinâmica via `{"soma": ...}` da TASK-045, resolvida uma única vez
antes de qualquer token ser adicionado). Extraída `_validar_qtd_adicionar()` para eliminar a
duplicação de validação entre o caso de ativo único e o de lista. Diferente do ativo dinâmico da
TASK-046 (que não combina com `qtd` dinâmico), lista SIM combina — os dois recursos são
independentes.

Exemplo exato do pedido: `{"acao": "adicionar_ativo", "tabela": "outros", "ativo": ["90525",
"90542", "92540"], "qtd": -1}`.

7 testes novos (`tests/test_ajustes_planilhas.py`). Suíte completa (755 testes) sem regressão.
Smoke test real-server + Playwright confirma o fluxo completo pela UI (alternar "ativo" para "lista
de códigos", preencher códigos separados por vírgula, modelo em memória e frase ao vivo corretos).

---

## 2026-10-06 — TASK-046: `adicionar_ativo` com ativo dinâmico (nome = código encontrado na própria linha)

**Tipo:** feature (motor de Ajustes) · `services/ajustes_planilhas.py`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-046-06-10-2026.md`

Usuário pediu, logo após a TASK-045: "caso tenha um trafo mono TR1* adicione com a mesma quantidade
negativa o texto TR1* acrescido com VP no fim desse texto" — o nome do ativo adicionado precisa ser
o CÓDIGO especificamente encontrado na linha (ex.: `TR110`, `TR115`...), não um literal fixo
`"TR1*VP"` (curinga de busca não é um código válido).

Implementado: `ativo` passa a aceitar também `{"igual_a": SELETOR, "prefixo"?: texto, "sufixo"?:
texto}` — para CADA item da própria linha que casa com `SELETOR`, gera um token novo
`"<prefixo><código do item><sufixo>"`, com quantidade = quantidade DAQUELE item × `qtd` (aqui `qtd`
é o FATOR por item, não soma — não combina com `{"soma": ...}` da TASK-045). Reaproveitados os
mesmos `_contexto()`/`_casa()` da TASK-045, nenhuma mudança em `regras_dominio.py`.
`_aplicar_um_ativo()` extraída de `_t_adicionar_ativo` (merge/append de um token) para ser chamada
em loop, uma vez por item casado. **Guard de idempotência:** como `igual_a` tipicamente é um
curinga amplo (`"TR1*"`) que também casaria com o próprio token recém-criado (`"TR110VP"` começa
com `"TR1"`), um item cujo nome já carrega o `prefixo`/`sufixo` configurado é excluído dos
candidatos — sem isso, reaplicar a ação geraria `"TR110VPVPVP..."` indefinidamente.

Exemplo exato do pedido: `{"acao": "adicionar_ativo", "tabela": "outros", "ativo": {"igual_a":
"TR1*", "sufixo": "VP"}, "qtd": -1, "quando": {"tem": "TR1*"}}`.

8 testes novos (`tests/test_ajustes_planilhas.py`). Suíte completa (748 testes) sem regressão.
Smoke test real-server (uvicorn em processo, `database.DB_PATH` apontado para um SQLite temporário
antes do `init_db()`, nunca `banco_resumo.db`) + Playwright confirma o fluxo completo pela UI
(alternar "ativo" para dinâmico, preencher seletor/prefixo/sufixo/fator, modelo em memória e frase
ao vivo corretos). Mesmo achado colateral da TASK-045 reencontrado (guard `C1-OUT-NUM` com código
hifenado + token negativo) — teste ajustado para evitar, fora do escopo.

---

## 2026-10-06 — TASK-045: `adicionar_ativo` com quantidade dinâmica (mesma quantidade de outro ativo)

**Tipo:** feature (motor de Ajustes) · `services/ajustes_planilhas.py`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-045-06-10-2026.md`

Usuário, cadastrando uma regra dentro de um contexto (TASK-044), pediu: "se tem U3 adicione em mesma
quantidade 90277 negativo" — a quantidade de "90277" precisa igualar a quantidade de U3 encontrada
NAQUELA linha, que varia (`1-U3`, `2-U3`, `3-U3`...), não um número fixo. Confirmado com o usuário
que U3 realmente varia (não é sempre 1) antes de implementar, já que `adicionar_ativo.qtd` só
aceitava um número literal.

Implementado: `qtd` passa a aceitar também `{"soma": SELETOR, "fator"?: número}` — soma das
quantidades dos itens da PRÓPRIA linha que casam com `SELETOR` (mesma linguagem de seletor de
sempre: código, curinga, `@GRUPO`, lista, regex), multiplicada por `fator` (padrão 1; `-1` gera o
prefixo `"*"` automaticamente, mesma regra de sinal da TASK-038). Reaproveitado:
`services/ajustes_planilhas.py::_contexto()` (já construía um `_Linha` a partir do texto de
qualquer linha, usado por `_quando()`) e o mesmo `_casa()` que `_cond()`'s `soma` já usa para
validação — nenhuma mudança em `regras_dominio.py`. Soma zero (seletor não casa na linha) é no-op,
não erro. `descrever_acao()` e a UI da gaveta (`painel_ajustes.js`) ganharam suporte ao novo modo
("quantidade": fixa/dinâmica, com campos "soma de"/"fator").

Exemplo exato do pedido: `{"acao": "adicionar_ativo", "ativo": "90277", "qtd": {"soma": "U3",
"fator": -1}, "quando": {"tem": "U3"}}`.

8 testes novos (`tests/test_ajustes_planilhas.py`). Suíte completa (740 testes) sem regressão.
Smoke test real-server + Playwright confirma o fluxo completo pela UI (trocar para quantidade
dinâmica, preencher seletor+fator, modelo em memória correto). Achado colateral, não corrigido
(fora do escopo): um guard de contrato pré-existente (`C1-OUT-NUM`) trata mal códigos hifenados
(ex.: `SUP-L`) combinados a um token negativo `"*"` — reproduz igual com `qtd` fixo, não é algo
introduzido por esta tarefa.

---

## 2026-10-06 — TASK-044: contextos dentro de um projeto (terceira camada de Ajustes)

**Tipo:** feature (arquitetura de dados + UI) · `database.py`, `scripts/schema_supabase.sql`,
`services/contextos.py`, `services/repo_json.py`, `routers/validacao_ajustes.py`,
`static/resumo.html`, `static/resumo.js`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-044-06-10-2026.md`

Usuário pediu: dentro de um mesmo projeto, cadastrar "contextos" nomeados (ex.: "obras de 34,5kV"),
cada um com seu próprio conjunto de ajustes ligados/desligados/sobrescritos, selecionáveis por um
seletor na linha de totais (postes/cabos/equipamentos).

Desenho (decidido com o usuário após 3 perguntas de estratégia): **terceira camada de overlay**
sobre o que já existia — `padrão (DEFAULT) → overlay do projeto → overlay do contexto`. A peça-chave
é que `services/ajustes_camadas.py::efetivo(padrao, overlay)` não sabe que só existe uma camada —
ela só recebe uma lista e um overlay, então pode ser **chamada em cadeia** sem alterar uma linha
dela: `efetivo(efetivo(padrao, overlay_projeto)[0], overlay_contexto)`. Zero mudança no motor de
camadas existente.

Implementado:
- **Tabelas novas**: `contextos` (metadados de quais contextos existem por projeto — pensada para
  um dia ser compartilhada por Regras de Domínio também) e `ajustes_contextos`/
  `ajustes_contextos_historico` (overlay de Ajustes por contexto, mesmo shape do overlay de
  projeto). Nenhuma tabela existente foi alterada.
- **`services/repo_json.py`**: nova classe `RepoJsonContexto`, chaveada por
  `(projeto_codigo, contexto)` — irmã de `RepoJson` (que continua intacta, chaveada só por
  `projeto_codigo`).
- **`routers/validacao_ajustes.py`**: `resolver(projeto_codigo, contexto=None)` resolve a cadeia;
  toda rota (`GET/POST /`, `/preview`, `/preview-lote`) ganhou o campo opcional `contexto` —
  **sem ele, comportamento idêntico a antes** (garantido por teste de regressão). Novas rotas
  `GET/POST /contextos` e `POST /contextos/excluir`.
- **Frontend**: novo seletor "Contexto" na linha de totais do Resumo (`#resumo-rede-painel`),
  mesmo padrão visual do seletor de projeto, com "+ Novo"/"✕ Excluir" admin-only, persistido por
  projeto em `localStorage`. O botão "Ajustar" e o editor de Ajustes da gaveta (`painel_ajustes.js`)
  passam a considerar o contexto selecionado — a gaveta mostra "Editando o contexto..." e desabilita
  Histórico/Restaurar nesse modo (decisão de escopo: histórico por contexto fica para outra tarefa).
- Trocar de contexto **não recalcula nada na tela na hora** — só define o que roda na próxima vez
  que o usuário clicar em "Ajustar" (decisão do usuário).

Escopo desta rodada: só **Ajustes**. Modo autônomo (pasta monitorada) e Regras de Domínio não usam
contexto ainda — ficam para tarefas futuras, reaproveitando o mesmo desenho de dados.

11 testes novos (`tests/test_ajustes_contextos.py`): CRUD, resolução em cadeia, independência entre
contextos, exclusão isolada, regressão byte a byte sem `contexto`, `preview-lote` respeitando o
contexto. Suíte completa (732 testes) sem regressão. Smoke test real-server + Playwright confirma o
fluxo completo pela UI: criar contexto, editar dentro dele na gaveta, salvar, e confirmar que o
projeto (sem contexto) permanece intacto.

---

## 2026-10-06 — TASK-043: suporte a .dwg via conversão online (CloudConvert)

**Tipo:** feature · `services/cloudconvert_service.py`, `routers/upload.py`, `routers/health.py`,
`services/autonomo/pipeline.py`, `static/index.html`, `static/script.js`, `requirements.txt` ·
`.ai/tasks/TASK-043-06-10-2026.md`

Usuário perguntou se o programa conseguia ler `.dwg`. Resposta: não diretamente — é formato binário
proprietário da Autodesk, e `ezdxf` (usado em `services/dxf_service.py`) só lê `.dxf`. Usuário pediu
para o próprio programa converter internamente usando uma ferramenta online. Duas decisões
confirmadas antes de implementar: (1) serviço CloudConvert (API v2, chave simples, free tier) em
vez de Autodesk Platform Services (OAuth2 mais complexo); (2) credencial no mesmo padrão já usado
para a chave do Gemini (variável de ambiente como padrão + o usuário pode digitar a própria, salva
no servidor só depois de confirmada por um uso com sucesso).

Implementado:
- **`services/cloudconvert_service.py`** (novo): `converter_dwg_para_dxf()` — cria job no
  CloudConvert (import/upload → convert → export/url), envia o arquivo, espera (polling), baixa o
  `.dxf` resultante. `ErroConversaoDwg` com mensagem clara em qualquer falha.
- **`routers/upload.py`**: resolve a credencial (header `X-CloudConvert-Key` > `configuracoes` >
  `CLOUDCONVERT_API_KEY` do ambiente) e converte `.dwg` → `.dxf` antes de extrair, em `POST /upload`
  e `GET /extract-local`. `.dxf`/`.pdf` continuam sem tocar no CloudConvert.
- **`services/autonomo/pipeline.py`**: `.dwg` aceito na pasta monitorada (`EXTENSOES`), convertido
  antes de ler (só credencial salva/padrão — sem header possível em segundo plano).
- **Frontend** (`static/index.html`/`script.js`): `.dwg` aceito no seletor de arquivo da tela
  Leitor; novo botão "Chave DWG" + modal para configurar a própria chave do CloudConvert
  (localStorage, enviada como header).
- **`routers/health.py`**: novo campo `cloudconvert_key_source`, mesmo padrão de `ai_key_source`.

Pela primeira vez no projeto, uma funcionalidade de **leitura de arquivo** passa a depender de
internet (só para `.dwg`) — aceito como exceção pontual, mesmo padrão já aplicado à IA e à
sincronização com Supabase.

Testado com `httpx` sempre mockado (nenhuma chamada de rede real em teste): 9 testes novos de
`cloudconvert_service`, 10 de dispatch em `upload.py`, 3 do pipeline autônomo. `tests/conftest.py`
passou a zerar `CLOUDCONVERT_API_KEY` do ambiente também. Suíte completa (721 testes) sem
regressão. Smoke test real-server + Playwright confirma a UI nova sem erros de JS.

---

## 2026-10-06 — TASK-042: editor de Regras de Vinculação e Regras de Conectores dentro do modal

**Tipo:** feature (UI) · `static/resumo.html`, `static/resumo.js` · `.ai/tasks/TASK-042-06-10-2026.md`

Usuário pediu para tornar as `regras_vinculacao` (quantidade de cabos e compatibilidade por tipo
de estrutura) editáveis numa UI dentro do modal de vinculação, e para poder acessar o ativo
composto (cabo+estrutura vinculados, ex. `CAA2_U4`) a partir de uma regra com formato amigável,
também dentro do modal.

Investigação confirmou dois mecanismos já existentes, mas sem UI própria no lugar certo:

- `regras_vinculacao` já tinha endpoint (`GET/POST /api/regras-vinculacao`), mas nenhuma tela
  editava — só o algoritmo de vínculo automático lia.
- O ativo composto só existe em memória durante o cálculo da Totalizadora
  (`itens_vinculo_para_totalizadora`) e já podia ser convertido em algo real (ex. um conector) via
  **Regra de Conversão com `origem: VINCULO`** (TASK-011/032/034) — só que misturado com as regras
  de CABOS/OUTROS na view Totalizadora, exigindo digitar `ativo_de` cru (`"CAA2_U4"`).

**Decisão: não estender o motor de Ajustes** (`services/ajustes_planilhas.py`) para ler ativos
compostos — ele roda antes do vínculo/Totalizadora existir, e replicar a lógica de vínculo ali
duplicaria o que a Regra de Conversão com `origem: VINCULO` já resolve. Em vez disso, duas seções
accordion novas foram adicionadas dentro de `#modal-vinculacao`:

1. **Regras de Vinculação**: tabela editável (tipo/qtd/compatibilidade/descrição) sobre o endpoint
   já existente `/api/regras-vinculacao`.
2. **Regras de Conectores**: formulário amigável (tipo de cabo + tipo de estrutura + ativo
   resultante + quantidade) sobre a MESMA tabela de Regras de Conversão da Totalizadora
   (`tableStates.regras.data`, `/api/regras/conversao`), filtrada para `origem === 'VINCULO'` e
   traduzindo `ativo_de ↔ "<CABO>_<ESTRUTURA>"` automaticamente. Editar aqui ou na Totalizadora é a
   mesma regra.

Nenhum endpoint ou tabela nova. Verificado via smoke test real-server + Playwright (DB temporário,
nunca o `banco_resumo.db`): adicionar/editar/salvar nas duas seções persiste corretamente
(confirmado por `GET` de volta), e uma linha inválida é rejeitada pelo backend com erro visível.

---

## 2026-10-05 — TASK-039/040/041: disponibilidade de dados e reconexão com o Supabase

**Tipo:** correção de robustez + feature · `services/supabase_client.py`,
`services/connectivity_monitor.py`, `app.py`, `routers/obras.py`, `recs.py`, `projetos.py`,
`admin.py`, `static/resumo.html`, `static/resumo.js`, `tests/test_supabase_client.py` ·
`.ai/tasks/TASK-039-05-10-2026.md`, `TASK-040-05-10-2026.md`, `TASK-041-05-10-2026.md`

Usuário relatou "a conexão com o Supabase está se perdendo após um tempo" e pediu garantia de que
os dados fiquem sempre disponíveis, mantendo a última versão carregada em caso de queda. Investigação
confirmou a causa raiz:

- `services/supabase_client.py` criava o cliente Supabase como singleton **uma única vez por
  processo** — existia `reset_supabase_client()`, mas nada no projeto a chamava.
- Toda falha de chamada à nuvem era engolida (`except Exception: pass`, sem log nenhum) em
  `routers/obras.py`, `recs.py`, `projetos.py`, `admin.py` (padrão já catalogado em `.ai/STATE.md`
  item 21).
- Resultado: quando o Supabase fecha uma conexão ociosa (comportamento normal do
  PostgREST/pgbouncer), a primeira falha era silenciosa e **nunca mais havia tentativa de
  reconexão** até o processo reiniciar.
- `services/connectivity_monitor.py` já existia pronto (ping periódico + evento WebSocket), mas
  nunca era chamado, e era restrito a `APP_MODE != "server"`.
- A gravação local (SQLite) **já era garantida** independente da nuvem (`save_obra` sempre grava
  local) — o problema real na escrita era a falta de aviso, não perda de dado.

Implementadas as três tarefas registradas:

- **TASK-039** (reconexão reativa): `get_supabase()` ganhou um TTL proativo de 180s (recria o
  cliente sozinho mesmo sem falha observada) combinado com reset reativo via nova função
  `registrar_falha(contexto, exc)`, que reseta o singleton só quando a exceção é classificada como
  erro de rede (`httpx.TransportError`/`httpx.TimeoutException`/`ConnectionError`/`OSError`/
  `TimeoutError`), nunca em erro de autenticação/dados.
- **TASK-040** (monitor de conectividade): `services/connectivity_monitor.py` corrigido (loop
  `asyncio` capturado uma vez no startup, broadcast via `run_coroutine_threadsafe` em vez de
  criar/fechar loops a cada ciclo) e ligado em `app.py` no startup, **sem** a restrição de
  `APP_MODE` — roda em desktop e servidor. A transição offline→online agora força
  `reset_supabase_client()`. Novo indicador visual no frontend (`static/resumo.html`/`resumo.js`,
  `#indicador-conectividade`) escuta o evento WebSocket `{"type": "connectivity", ...}` e exibe um
  selo "Sem conexão com a nuvem — trabalhando localmente" quando offline.
- **TASK-041** (visibilidade de falha): os 16 pontos de `except Exception: pass` silenciosos em
  `obras.py`/`recs.py`/`projetos.py`/`admin.py` agora chamam `registrar_falha(contexto, e)` —
  sempre loga em `warning`, nunca mais 100% silencioso. Decisão tomada: manter o comportamento
  atual de escrita (sempre grava local, sempre responde sucesso), só com log — mudança de maior
  impacto (erro visível ao usuário) descartada por não ter sido pedida explicitamente e por exigir
  tratamento em todas as telas de salvamento. Confirmado por leitura de código que nenhuma leitura
  troca dado local bom por resultado de nuvem vazio por falha disfarçada (o `except` sempre
  intercepta antes do `if res.data is not None`).

Novo `tests/test_supabase_client.py` (14 testes) cobre TTL, reset reativo por tipo de exceção e log
sempre emitido. Suíte completa (699 testes) sem regressões.

---

## 2026-10-05 — TASK-038: `adicionar_ativo` aceita quantidade negativa (`*`, linha viva)

**Tipo:** melhoria (backend) · `.ai/tasks/TASK-038-05-10-2026.md`

Pedido do usuário: um ajuste que, ao achar TR110 (trafo mono) na linha, adiciona um conjunto de
ativos — quatro positivos e um negativo (`*1-PR`). A ação `adicionar_ativo` (TASK-022) exigia
`qtd > 0` e, mesmo que aceitasse, formatava o token com `-` literal, que o tokenizador de Outros
não lê como sinal (só `*` é lido como negativo, por `services/orcamento_calc.py:tokenizar_outros`).

- `services/ajustes_planilhas.py`: `validar_acoes` passa a aceitar `qtd` negativo (só `0` é
  rejeitado); `_num` agora reconhece um token existente com `*` e devolve negativo; nova
  `_fmt_qtd()` (inverso de `_num`) gera o prefixo `*` corretamente; usada nas 3 montagens de token
  de `_t_adicionar_ativo` (novo, somar, substituir) e na frase legível do ajuste.
- `data/manual_regras_e_ajustes.md`: documentado o novo comportamento com exemplo.
- Confirmado isolado: `_num`/`_juntar`/`_RE_NUM`/`_fmt_qtd` são privados deste módulo (não
  importados de `regras_dominio.py`), sem efeito em validação de domínio, Totalizadora ou modo
  autônomo.

Testado: `pytest tests/test_ajustes_planilhas.py` (76 testes, 5 novos) + `pytest tests/` completo
(684 testes) + verificação manual do exemplo completo do usuário (5 ações, mesma condição `{"tem":
"TR110"}`) produzindo exatamente `1-TR110 1-PR15 2-P50 1-PR127 *1-PR 3-EST35`.

---

## 2026-10-03 — TASK-037: coordenada do Leitor chega ao Resumo (fecha o 4º bug do vínculo)

**Tipo:** correção de bug · `.ai/tasks/TASK-037-03-10-2026.md`

Último elo que faltava para o vínculo automático por coordenada (TASK-032) funcionar de ponta a
ponta na **tela manual**. Os três bugs da TASK-036 (dígito do tipo, ativo composto, descarte de
`_x`/`_y` no DXF) já tinham corrigido o lado do dado; faltava a ponte entre as duas telas: o botão
"Processar" da tela Leitor (`static/script.js`) montava o objeto exportado para
`localStorage['processar_dados']` só com `{pagina, texto, cor, layer, entidade, operacao, ativo}` —
sem `_x`/`_y`, mesmo quando `extractedDataCache[idx]` já os tinha.

- `static/script.js`: inclui `_x`/`_y` condicionalmente no objeto exportado, mesmo padrão já usado
  em `services/autonomo/montagem.py:_item_exportado`.

**Testado end-to-end com um DXF real, sem nenhum atalho sintético** (primeira vez que isso acontece
nesta funcionalidade): servidor real + Playwright — upload de um DXF com 1 cabo e 2 estruturas →
extração real → "Processar" → aba do Resumo já recebe `_x`/`_y` → vínculo automático cadastrado via
API → `tentarVincularAutomaticamente()` vincula corretamente, respeitando a ordem de proximidade.
`pytest tests/` completo continua em 679/679 (mudança restrita a `static/script.js`, sem teste
Python tocando esse arquivo).

Com isso, a cadeia completa TASK-032/033/034/036/037 está funcional nos dois modos (manual e
autônomo). Única pendência remanescente conhecida: nenhuma, para esta funcionalidade.

---

## 2026-10-03 — TASK-036: vínculo cabo↔estrutura no modo autônomo (+ 3 bugs achados e corrigidos)

**Tipo:** nova funcionalidade (backend) + correção de bugs pré-existentes · `.ai/tasks/TASK-036-03-10-2026.md`

Implementa o que a TASK-036 tinha registrado: o pipeline do modo autônomo (TASK-031) agora gera
vínculo cabo↔estrutura/poste por coordenada, roda a validação de vínculo (TASK-033) e aplica a
origem `VINCULO` das Regras de Conversão (TASK-034) ao montar o orçamento — as três coisas que,
até aqui, só existiam dentro de `static/resumo.js`.

- `services/autonomo/montagem.py`: `_x`/`_y` passam a chegar até as linhas de Cabos/Outros; nova
  `prefixo_cabo()`.
- Novo módulo `services/autonomo/vinculacao.py`: porte fiel de `tentarVincularAutomaticamente`,
  `avaliarVinculacaoLocal` e do bloco de ativo composto de `syncTotalizadora`.
- `services/autonomo/pipeline.py`: nova etapa `vincular`; `Contexto.regras_vinculacao`.
- `services/autonomo/totalizadora.py`: `montar_totalizadora()` ganhou `itens_vinculo`; composto
  sem regra correspondente nunca é mantido (decisão já tomada na TASK-034, replicada aqui).
- **Decisão do usuário:** auto-link sem combinação válida nunca bloqueia nem vira pendência de
  confirmação — só um achado de aviso no relatório.

**Três bugs pré-existentes achados durante a implementação (confirmados com o leitor real antes de
corrigir, não por suposição) — corrigidos por decisão explícita do usuário:**

1. `_tipoEstruturaVinculo` (TASK-032/033) pegava o dígito da **quantidade**, não do tipo da
   estrutura — o ativo real de uma linha Outros é sempre `"<qtd>-<código>"` (ex.: `"1-U4"`, nunca
   `"U4"` sozinho; confirmado rodando o motor de regras real). Isso quebrava silenciosamente todo
   o casamento de regra por tipo desde que a TASK-032/033 foi mesclada. Corrigido nos dois lados
   (`static/resumo.js` e `services/autonomo/vinculacao.py`).
2. O mesmo problema na geração do ativo composto em `syncTotalizadora` (`CAA2_1-U4` em vez de
   `CAA2_U4`). Mesma causa raiz, mesma correção.
3. `services/dxf_service.py:extract_dxf_content()` calculava `_x`/`_y` só para ordenar a lista de
   saída e depois os **descartava** antes de devolver — nenhum arquivo DXF jamais teve coordenada
   disponível para o vínculo automático (só PDF, desde a captura adicionada no início desta
   sessão). Corrigido: os campos agora ficam no retorno.

**Pendência registrada, não corrigida aqui** (`.ai/tasks/TASK-037-03-10-2026.md`): `static/script.js`
(botão "Processar" da tela Leitor) ainda não repassa `_x`/`_y` para `localStorage['processar_dados']`
— o vínculo automático por coordenada continua não funcional na **tela manual** até isso ser
corrigido (o modo autônomo não depende de `script.js`, por isso já funciona com as correções acima).

Testado: `pytest tests/` completo (679 testes, todos passando — 23 novos em
`test_autonomo_vinculacao.py`, +2 em `test_autonomo_montagem.py`, +2 de pipeline completo em
`test_autonomo_pipeline.py`) + verificação manual em servidor real/Playwright na tela `/resumo`
confirmando o auto-link escolhendo a estrutura mais próxima com os bugs corrigidos.

---

## 2026-10-03 — TASK-035: filtro por projeto no histórico do modo autônomo (+ registro da TASK-036)

**Tipo:** melhoria (backend + frontend) · `.ai/tasks/TASK-035-03-10-2026.md`

Pedido do usuário depois de discutirmos uma estratégia para rodar vários projetos pelo modo
autônomo (TASK-031) de uma vez e só depois ir conferindo. A arquitetura já suporta múltiplos
projetos numa única instância (`entrada/<código-do-projeto>/`), mas o histórico de execuções
(`GET /api/autonomo/execucoes`) só filtrava por status — inviável de triar visualmente num lote
grande com vários projetos misturados.

- `services/autonomo/execucoes.py:listar()` ganhou o parâmetro `projeto` (filtro `AND` com
  `status`); `routers/autonomo.py` repassa.
- `static/autonomo.html`/`autonomo.js`: novo `<select id="filtro-projeto">`, populado a partir de
  `GET /api/projetos` (que devolve `{"projetos": [...]}` — achado durante o teste: a primeira
  versão do JS lia como lista direta).
- Testado: `pytest` (33 testes, 2 novos) + servidor real/Playwright confirmando o seletor populado
  e os dois filtros combinando sem erro de console.

**Também registrada (sem implementar ainda):** `.ai/tasks/TASK-036-03-10-2026.md` — o modo
autônomo foi construído (TASK-031, 2026-10-02) **antes** do vínculo cabo↔estrutura/poste
(TASK-032/033/034, 2026-10-03) e não tem nenhuma referência a ele (confirmado por busca no
código): não gera vínculo automático por coordenada, não roda a validação de vínculo, e não
aplica a origem `VINCULO` das Regras de Conversão ao montar o orçamento. Um arquivo processado
pelo autônomo num projeto que depende de vínculo sai sem essa parte do orçamento. Fica registrada
para decisão/implementação futura.

---

## 2026-10-03 — TASK-032/033/034: vínculo cabo↔estrutura, validação e ativo composto

**Tipo:** nova funcionalidade (backend + frontend) · `.ai/tasks/TASK-032-02-10-2026.md`,
`TASK-033-02-10-2026.md`, `TASK-034-02-10-2026.md`

**TASK-032** — modal "VINCULAR CABOS" na aba Resumo: vincula cada cabo a 1 estrutura/poste
(operação M/\*M) ou 2 (operação I), manualmente ou por um algoritmo opcional de vínculo
automático por coordenada (`_x`/`_y` — já existia no DXF; `services/pdf_service.py` passou a
capturar o `bbox` do PyMuPDF para o PDF também). Nova tabela por projeto `regras_vinculacao`
(quantidade de cabos e compatibilidade — mesmo/qualquer tipo-fase-operação — exigida por tipo de
estrutura, identificado pelo primeiro dígito do ativo). O vínculo vive como um campo
(`vinculoEstruturas`) direto na linha do cabo, o que faz persistência (undo/redo, salvar/carregar
obra) funcionar sem nenhum caminho novo de serialização — só foi preciso estender a `deepClone()`
compartilhada de `static/resumo.js`, que até então descartava qualquer campo fora de
`entidade`/`operacao`/`ativo`/`qtdAtivos` (incluindo a própria coordenada).

**TASK-033** (depende da TASK-032) — ao validar (botão já existente `#btn-validar`), cada
estrutura/poste tem seus vínculos conferidos contra a regra do seu tipo; qualquer violação
(zero vínculos, quantidade errada, incompatibilidade) gera um achado de aviso. Zero vínculos é
perdoado se existir, em qualquer lugar do projeto, um cabo M/\*M ou um ativo `RETCA`.
**Decisão de integração**: em vez de um painel visual novo (como o desenho original previa, nos
moldes do "Resumo da rede" da TASK-008), os achados foram integrados ao sistema de Validação já
existente (`routers/validacao.py`/`services/regras_dominio.py`, construído em sessões anteriores)
— como o vínculo só existe no frontend, a verificação roda localmente
(`avaliarVinculacaoLocal()`) e seus achados são concatenados aos de `/api/validacao/planilhas`
antes de exibir, sem nenhuma mudança no motor/backend de domínio.

**TASK-034** (depende da TASK-032) — cada vínculo gera um ativo composto
`<prefixo_do_cabo>_<ativo_da_estrutura>` (ex.: `CAA2_N4`), que existe só dentro dos dados do
vínculo mas fica disponível como entrada para as Regras de Conversão (TASK-003) via uma nova
origem `VINCULO` — ex.: `origem=VINCULO, ativo_de=CAA2_N4, ação=ADIÇÃO, ativo_para=ALCA2,
fator=3` gera `3-ALCA2` na Totalizadora ao montar o orçamento, pelo mesmo motor ADIÇÃO/SUBST que
já existe. Diferente de `CABOS`/`OUTROS`, um composto sem regra correspondente não fica na
Totalizadora (não é um código de orçamento real).

**Renumeração**: as três tarefas foram inicialmente registradas como TASK-009/010/011, mas esses
números já tinham sido usados por outra sessão (fundidos em `main` enquanto esta sessão
trabalhava) — renumeradas para TASK-032/033/034 antes de implementar, mesma situação já ocorrida
no início desta sessão com TASK-001-004.

---

## 2026-10-02 — TASK-031: achados do primeiro teste real (projeto 027/229)

**Tipo:** Backend (modo autônomo)

(1) Ajustes habilitados do projeto 027 somavam **53 ações** e o limite de 50 (válido nas rotas/tela) barrava o arquivo inteiro no autônomo → `ajustar(..., limite_acoes)` e `validar_acoes(..., limite)` ganharam parâmetro;
o pipeline usa `LIMITE_ACOES_AUTONOMO = 300`; rotas e tela continuam com 50. (2) A mensagem de "tipo de arquivo não suportado" agora mostra a extensão encontrada e o nome do arquivo.

---

## 2026-10-02 — INCIDENTE: testes rodaram contra o Supabase real (corrigido) — `docs/INCIDENTE_TESTES_NA_NUVEM_2026-10-02.md`

**Tipo:** Testes / segurança de dados

Ao rodar `pytest` no PC do usuário (com `.env` do Supabase), os testes usaram a nuvem REAL e gravaram regras/ajustes/prompts/obras de teste (DEFAULT, 229, 027; usuários `@x.com`, `u1`, `u2`).
Causa: os testes assumiam "sem credenciais" (Linux de desenvolvimento) e nada garantia isso. Correção: `tests/conftest.py` zera `SUPABASE_*`/chaves de IA, usa banco temporário por teste e recusa rodar
(`pytest.exit`) se houver cliente Supabase; `tests/test_isolamento.py` confere. Roteiro de verificação/recuperação (via histórico) no documento do incidente. **Regra nova: teste nunca toca a nuvem; precisa de nuvem → cliente falso.**

---

## 2026-10-02 — TASK-031: preparação para o Windows (trabalhador por socket, arquivo em uso, roteiro de testes)

**Tipo:** Backend + scripts + docs · `.ai/tasks/TASK-031-02-10-2026.md`

O processo auxiliar do leitor passou a falar por **socket local com token** (no `.exe` sem console não há stdin/stdout) e abre sem janela no Windows; `app.py` repassa `<porta> <token>`.
A vigia **adia** o movimento de arquivo em uso (PermissionError) em vez de reprocessar/errar. `build.bat` inclui `quickjs`. Novos: `docs/ROTEIRO_TESTE_WINDOWS_MODO_AUTONOMO.md`,
`scripts/autonomo_verificar_ambiente.py`, `scripts/autonomo_arquivos_teste.py`.

---

## 2026-10-02 — TASK-031 fase D: tela de controle do modo autônomo (`/autonomo`)

**Tipo:** Frontend + correções no backend do modo autônomo · `.ai/tasks/TASK-031-02-10-2026.md`

Página `/autonomo` (só admin; link no Painel Admin): estado, ligar/desligar, varrer agora, **confirmação de exclusões (Sim / Não / Sim para todos / Não para todas)**, histórico com
detalhes, reverter e reprocessar, e configuração (dono das obras, pastas, intervalo, espera). Sem biblioteca externa; só `textContent`. Correções achadas no teste real: varreduras
simultâneas (agora serializadas) e falhas pré-pipeline passam a aparecer no histórico. O modo continua **desligado por padrão**.

---

## 2026-10-02 — TASK-031 fase C: pasta monitorada, fila e API de controle do modo autônomo

**Tipo:** Backend + `app.py` · `.ai/tasks/TASK-031-02-10-2026.md`

`services/autonomo/pasta.py` (vigia em thread: `entrada/<projeto>/` → pipeline, um por vez, só arquivos estáveis; originais para `processados/` ou `erros/` + `.erro.txt`),
`config_autonomo.py` (chave `autonomo_config` em `configuracoes`) e `routers/autonomo.py` (`/api/autonomo/*`, só admin: config, ligar/desligar, status, varrer, histórico,
confirmar/rejeitar/reverter/reprocessar). `app.py` ganhou: argumento `--trabalhador-leitor-js` (subprocesso do QuickJS no `.exe`), include do router e hooks `startup`/`shutdown`
que religam/param a vigia conforme a configuração. Desligado por padrão: nada muda até um admin ligar.

---

## 2026-10-02 — TASK-031 fase B: pipeline autônomo (aplicador de diff, Totalizadora, orçamento, obra e pasta)

**Tipo:** Backend (código novo, ainda sem rota/tela) · `.ai/tasks/TASK-031-02-10-2026.md`

`services/autonomo/` ganhou `aplicar.py` (porte de `aplicarOperacoesAjuste`), `totalizadora.py` (porte de `syncTotalizadora` + payload do orçamento),
`pipeline.py` (ler → montar → validar → ajustar → confirmação de exclusões → Totalizadora → orçamento → obra + pasta), `execucoes.py` (tabela local
`execucoes_autonomas`, criada sob demanda; sem Supabase) e `saida.py` (JSON + CSV). Exclusões (`excluir_linhas`, `remover_ativo`) deixam o arquivo
"aguardando_confirmacao"; `confirmar` (Sim / Sim para todos), `rejeitar`, `reverter`. O leitor roda em subprocesso (`trabalhador_js.py`).
`ramais.py` gera os ativos dos RAMAIS (decisão: entram no orçamento autônomo; etapa `ramais` do pipeline). **Alterar `aplicarOperacoesAjuste`, o modal RAMAIS, `syncTotalizadora`, `obterPayloadCalculo` ou a montagem em `resumo.js` exige rodar `tests/test_autonomo_*.py`** (o oráculo fatia o código real).

---

## 2026-10-02 — TASK-031 fase A: leitor via QuickJS e montagem em Python (modo autônomo)

**Tipo:** Backend (código novo, não carregado pelo app por padrão) · `.ai/tasks/TASK-031-02-10-2026.md`

`services/autonomo/leitor_js.py` executa o **mesmo** `static/regras_leitor_engine.js` num QuickJS em **processo separado** (a regex catastrófica
não é interrompida pelo limite de tempo do QuickJS; o processo principal mata e recria o trabalhador). `services/autonomo/montagem.py` porta a
montagem de Cabos/Outros/Ramais e o cálculo de quantidades de `resumo.js`. Paridade provada contra o JS real (Node) em `tests/test_autonomo_*.py`.
**Mudar `regras_leitor_engine.js` ou as funções de montagem em `resumo.js` exige rodar esses testes** (marcadores fatiados em `tests/oraculo_tela.py`).
Dependência nova: `quickjs`. Nenhuma mudança no modo manual.

---

## 2026-10-02 — TASK-030: manual de uso de regras e ajustes

**Tipo:** Documentação do usuário (backend leve + frontend) · `.ai/tasks/TASK-030-02-10-2026.md`

Manual em Markdown (`data/manual_regras_e_ajustes.md`) servido por `GET /api/manual` (login, qualquer perfil) e renderizado em `/manual`;
botão **Manual** na gaveta Regras de validação. Os exemplos JSON do manual são validados em `tests/test_manual_exemplos.py`
(ao mudar a linguagem de regras/ajustes, o teste acusa exemplo desatualizado — atualize o manual junto).

---

## 2026-10-02 — TASK-029: botão Ajustar executa todos os ajustes habilitados

**Tipo:** Ajustes (backend + frontend) · `.ai/tasks/TASK-029-02-10-2026.md`

`Ajustar` deixou de ser menu: abre um modal com um **cartão por ajuste habilitado que muda algo** (caixa por mudança, **Executar correção**)
e, acima, **Executar todas as correções** (em cadeia, ordem do cadastro, um passo de histórico). Nova rota `POST /api/validacao/ajustes/preview-lote`
(só simula; uma chamada devolve cartões, ajustes sem mudança, configs inválidas e a cadeia; limites 30 ajustes / 50 ações na cadeia).
Exclusões pedem confirmação extra; guarda da camada 1 e proteção de linha alterada mantidas. Após executar, os cartões são recalculados.

---

## 2026-10-01 — TASK-028: IA do chat consulta e comanda as obras salvas

**Tipo:** IA do chat (backend + frontend) · `.ai/tasks/TASK-028-01-10-2026.md`

Quando o pedido menciona **obra/obras**, o backend acrescenta ao prompt um **índice enxuto** (id | nome | data, até 50, só do usuário e
do projeto selecionado, sem `dados_json`) e o bloco `data/prompt_obras.txt` ao system prompt; o **conteúdo** de uma obra só vai
se o usuário pedir ("o que tem na obra X?") e só dela (até 2, teto de linhas). **Pedidos sem a palavra "obra" seguem idênticos
a antes** (nenhuma consulta ao banco, mesmo prompt). Comandos: a IA responde `{"acao_ui":"obra","acao":"adicionar|subtrair|carregar",
"obra_id":…}`; o frontend valida o id contra as obras reais (`GET /api/obras/{id}`), exige que a obra seja do projeto selecionado e
**sempre abre um popup de confirmação** (carregar avisa que substitui tudo); aplicar é um passo de histórico. A IA não salva nem
exclui obras. O Fast-Path não intercepta frases com "obra". Refatoração: `adicionarObraAoProjeto`/`subtrairObraDoProjeto`
compartilhadas pelo modal e pela IA; o modal Carregar Obra deixou de usar `innerHTML` com o nome da obra (XSS armazenado).

---

## 2026-10-01 — TASK-027: microfone no chat de IA do Resumo

**Tipo:** frontend (Resumo) · `.ai/tasks/TASK-027-01-10-2026.md`

Novo `static/voz.js`: botão 🎤 ao lado de enviar dita em pt-BR pelo reconhecimento de voz do **navegador** (Web Speech API); o
texto aparece ao vivo no campo (anexado ao que já estava) e **nunca é enviado sozinho**; para por clique, Esc ou 4 s de silêncio;
erros (permissão, sem microfone, sem rede, sem fala) com mensagem clara. **Leitura da resposta em voz alta** (`speechSynthesis`),
**desligada por padrão** (🔇/🔊, `localStorage` `chat_voz_resposta`), com "🔊 Ouvir" por resposta e ⏹ para parar (não lê blocos de
código nem tabelas). Sem suporte (ex.: janela de desktop): botão desabilitado com explicação; o chat digitado segue igual. Sem
backend, sem custo, sem guardar áudio.

---

## 2026-10-01 — TASK-026: controles de validação embutidos na gaveta lateral

**Tipo:** frontend (Resumo) · `.ai/tasks/TASK-026-01-10-2026.md`

A barra do Resumo ficou só com `Regras de validação`, `Ajustar ▾` e `Validar` (além de Montar Orçamento). As caixas de modo
(Determinística, Regras, IA, Pular IA se houver erro de contrato) e o **Validar ao montar** foram para a nova aba **Execução**,
a primeira da gaveta e visível também para o operador (são preferências locais, mesmas chaves de `localStorage`, valem sem a
gaveta aberta). Ao lado do `Validar` há um resumo do que vai rodar (ex.: "Det + Regras + IA · ao montar", com dica) e um
atalho ⚙ que abre a gaveta na aba Execução. `resumo.js` expõe `window.validacaoPrefs`; sem mudança de backend.

---

## 2026-10-01 — TASK-025: ajustes por IA com ações estruturadas

**Tipo:** camada 3 (IA) + frontend · `.ai/tasks/TASK-025-01-10-2026.md`

Novo prompt semente **`ajustar-planilhas`** (editável; entra sozinho na próxima inicialização, sem sobrescrever prompts já
editados) e rota `POST /api/validacao/ajustes-ia`: a IA propõe **ações** da linguagem de ajustes (inclusive **excluir**),
o backend valida cada ação (schema + grupos do projeto), **recusa `operacao_nova`**, deduplica, limita a 10 e roda o motor da
TASK-022 nas tabelas completas, devolvendo `acoes` (com frase em português) + `diff` + `destrutivo`. O botão **Pedir ajuste à IA**
do painel de validação agora abre a pré-visualização com aviso em vermelho quando há exclusão, confirmação extra ao aplicar e
**Cadastrar como ajuste** (admin): abre a aba Ajustes com as ações em rascunho. A rota antiga `/corrigir` (substituição de linha
inteira) continua existindo, mas a tela não a usa mais. Sem mudança de banco.

---

## 2026-10-01 — TASK-024: fluxo unificado Validar + Ajustar

**Tipo:** frontend (Resumo) · `.ai/tasks/TASK-024-01-10-2026.md`

O painel de validação agora oferece **Ajustar** por achado (ajustes cadastrados e ligados que corrigem a regra do achado),
**Ajustar tudo (determinístico)** e **Pedir ajuste à IA** (só para achados sem ajuste cadastrado), com filtro
Todos · Determinístico · IA. Todo ajuste passa por uma **pré-visualização** (editar / inserir / EXCLUIR / reordenar, com caixa
por mudança; nada é aplicado sem aceite). Aplicar é **um passo de histórico** (um Ctrl+Z desfaz o lote), ignora mudança cuja
linha foi alterada desde a pré-visualização e **revalida automaticamente**, mostrando "antes do ajuste". No Validar ao
montar, o ciclo painel → ajuste → revalidação continua até o usuário decidir; se sobrar só info, segue para o orçamento.
Novo botão **Ajustar ▾** na barra roda um ajuste cadastrado avulso. "Não, vou corrigir" deixou de chamar a IA sozinho.

---

## 2026-10-01 — TASK-023: cadastro de ajustes recorrentes por projeto (receitas)

**Tipo:** persistência + frontend admin · `.ai/tasks/TASK-023-01-10-2026.md` · **exige SQL no Supabase** (`scripts/schema_supabase.sql`)

Receitas de ajuste (ações da TASK-022 com nome, ligada/desligada e vínculo a regras de validação) agora são **cadastradas**:
`DEFAULT` guarda o container completo e cada projeto só o overlay (adicionadas/sobrescritas/ocultas), como as regras
(`services/ajustes_camadas.py`). Tabelas novas `ajustes_planilhas` e `ajustes_planilhas_historico` no SQLite e no Supabase
(nuvem primeiro, fallback local; falha de sincronização avisada) — gravação **só no botão Salvar**, com histórico e
reversão. Semente (todas desligadas): normalizar postes, `SUP-L→SUPL`, adicionar `1-SUPL` à linha com CFU (vinculada a
`C2-CFU-SUPL`), excluir linhas vazias, ordenar Outros. Rotas `GET/POST /api/validacao/ajustes` (+histórico, reverter,
restaurar-semente); `preview` aceita `receitas: [ids]`. Nova aba **Ajustes** na gaveta (`painel_ajustes.js`): selos, −/+,
voltar ao padrão, editor visual das ações, modo JSON e **pré-visualização do diff nas tabelas atuais** (nada é aplicado:
isso é a TASK-024). Operador vê a lista e pré-visualiza; só admin edita.

---

## 2026-10-01 — TASK-022: motor de ajustes determinísticos (ADR-006)

**Tipo:** novo motor de regra de negócio (backend) · `.ai/tasks/TASK-022-01-10-2026.md`

Novo `services/ajustes_planilhas.py`: 7 ações declarativas sobre Cabos/Outros (substituir texto/item, normalizar, ordenar,
excluir linhas, adicionar linha, adicionar/remover ativo na linha, mesclar duplicadas), reaproveitando condições, seletores e
grupos da linguagem de regras v2. Devolve um **diff** (editar/inserir/excluir/mover) sem alterar a entrada; a camada 1 descarta
ajuste que gere erro de contrato novo; a operação da linha só muda com `operacao_nova`; idempotente. Rotas
`POST /api/validacao/ajustes/preview` e `/descrever`. Sem mudança de banco. Telas e cadastro: TASK-023/024.

---

## 2026-10-01 — TASK-019: painel de regras de validação vai do Admin para o Resumo

**Tipo:** frontend (reorganização) · `.ai/tasks/TASK-019-01-10-2026.md`

O editor das regras de domínio (TASK-017/018) e os prompts de validação saíram de `/admin` e passaram a uma **gaveta
lateral** no Resumo (botão **Regras de validação**), com abas Regras · Grupos · Prompts da IA · Testar, que segue o projeto
de `#select-projeto` (o padrão de todos é uma opção da gaveta). **Só admin edita**; operador vê Regras e Grupos em leitura
(frases, selos). Reestilizado sem Tailwind (`painel_regras.css`); JS carregado sob demanda. Mesmos dados e rotas, nada perdido;
o Admin mostra só um aviso de mudança. Prompts e Testar são abas só de admin.

---

## 2026-10-01 — TASK-020: validação em modos independentes (determinística e IA)

**Tipo:** comportamento do frontend · `.ai/tasks/TASK-020-01-10-2026.md`

Barra do Resumo ganhou quatro caixas (preferência **local**, `localStorage` `validacao_modos`): **Determinística**
(contrato + regras), **Regras** (liga/desliga a camada 2; o contrato fica sempre com a determinística), **IA** e **Pular IA se
houver erro de contrato** (padrão ligado). Valem para o botão Validar e para "Validar ao montar". A IA roda **depois** da
determinística; com IA desligada nenhuma chamada a `/api/validacao/ia` é feita; com tudo desligado o Validar avisa "nada a
validar" e o orçamento segue. O painel diz o que rodou. "Corrigir com IA" só aparece com a IA ligada. Sem mudança de backend
(usa `incluir_dominio`).

---

## 2026-10-01 — TASK-021: desfazer/refazer sem desvio

**Tipo:** correção de comportamento (frontend) · `.ai/tasks/TASK-021-01-10-2026.md`

`static/resumo.js`: `undo()` passa a gravar o estado ao vivo quando ele está à frente do ponteiro do histórico e
`pushHistory()` ignora estado repetido (`historyDirty` controla o botão Desfazer). Funciona para os chamadores que
empilham antes de mudar e para os que empilham depois. Duas edições seguidas agora exigem dois Ctrl+Z; a correção
assistida deixou de precisar do "push duplo". Alterações em lote (ajustes das próximas tasks) serão um passo cada.

---

## 2026-09-30 — TASK-018 + TASK-017: regras em camadas por projeto e editor visual

**Tipo:** armazenamento/regra de negócio + frontend admin · `.ai/tasks/TASK-018-30-09-2026.md`, `TASK-017-30-09-2026.md`

**Camadas (TASK-018):** o `DEFAULT` guarda regras e grupos completos; cada projeto guarda só o **overlay**
(`adicionadas`, `sobrescritas`, `ocultas`, `grupos`) no mesmo `regras_json` — **sem SQL novo**. Efetivo = padrão + overlay,
vale a sobrescrita do projeto; o que muda no padrão chega aos projetos. Cópias integrais antigas (lista v1 ou container v2)
são lidas como overlay equivalente (o que faltava no padrão vira "oculta"), sem mudar resultados; virarão overlay no
próximo salvamento. Novo `services/regras_camadas.py` (funções puras) e `GET /api/validacao/regras` devolve `origem`,
`oculta`, `frase`, `grupos_origem`, `grupos_usados`, `avisos`. `restaurar-semente` num projeto = descartar os ajustes;
`adicionar-novas` passou a valer só no `DEFAULT`. Nova `GET /api/validacao/regras/ativos` (autocomplete).

**Editor visual (TASK-017):** novo `static/regras_editor.js` substitui o editor JSON da TASK-013: cartão em frase (vinda
do backend), selos *padrão / ajustada neste projeto / só deste projeto*, botão **−/+** (oculta/reexibe regra herdada),
"voltar ao padrão", construtor SE/ENTÃO/EXCETO SE com E/OU/NÃO e comparador de quantidade, autocomplete de ativos e
`@GRUPO`, editor de grupos com "usado por", assistente com 4 modelos, modo avançado (JSON) e painel de teste que explica
também por que uma linha **não** disparou.

**Correção:** os seletores de projeto do painel admin (prompts e regras) nunca listavam projetos, porque `/api/projetos`
devolve `{projetos: [...]}` e o código esperava uma lista; agora aceitam os dois formatos.

---

## 2026-09-30 — TASK-016: motor de regras de domínio v2 (ADR-005)

**Tipo:** mudança de motor de regras (backend) · `.ai/tasks/TASK-016-30-09-2026.md`

Regras de domínio passam a uma linguagem declarativa v2: condições sobre **quantidade** (`=, ≠, ≥, ≤, >, <, entre`), grupos
de ativos nomeados (`@ESTRUTURA_MT`), combinadores E/OU/NÃO, valores dinâmicos, regras sobre a **planilha inteira** (P50 e
contagens) e `explicacao` em português em cada achado; `descrever(regra)` gera a frase. Regras v1 salvas continuam
funcionando (conversão determinística); o resultado das 11 regras da semente é **idêntico** ao do motor anterior, provado
contra uma cópia congelada dele em milhares de cenários. Sem SQL novo (mesmo `regras_json`, agora um container com grupos).
Painel de achados mostra "Por quê". Camadas por projeto e editor visual: TASK-018 e TASK-017.

---

## 2026-09-30 — Regras de domínio pendentes (`C2-TR-EF`, `C2-BT-EXT`) e "adicionar regras novas da semente"

**Tipo:** regras de negócio (decisão do usuário) + ação do admin · `.ai/tasks/TASK-013-30-09-2026.md` (acréscimo)

`C2-TR-EF`: trafo com elo fusível só é aceito com chave `CFU` ou `CFUR` no poste (`CFA`, `CL` e as de reinstalação
não liberam). `C2-BT-EXT`: estrutura `SI` exige `RA2`, salvo estrutura `S#`. Ambas na semente, desligadas, aviso,
operações I/*I. O tipo `proibe` ganhou `exceto_se_regex`. Nova rota admin `POST /api/validacao/regras/adicionar-novas`
(botão no painel): acrescenta à versão do projeto só as regras da semente que faltam, sempre desligadas, sem tocar nas
existentes — necessária porque a semente só carrega quando não há versão DEFAULT. **Fecha o catálogo de validações.**

---

## 2026-09-30 — TASK-015: correção assistida por IA com aceite explícito

**Tipo:** nova funcionalidade (backend + frontend) · `.ai/tasks/TASK-015-30-09-2026.md`

`POST /api/validacao/corrigir`: a IA propõe correções só para as linhas citadas nos achados; o servidor reconfere
cada proposta (linha existente, muda algo, **passa de novo na camada 1**) e devolve também as descartadas com o
motivo. Na tela, "Não, vou corrigir" e o botão "Corrigir com IA" abrem as propostas (antes → depois) com aceite por
linha ou "Aceitar todas"; aplicar é um único passo no histórico e ignora linhas que mudaram. A operação nunca é
alterada. Refatorações mínimas: laço de lotes compartilhado (`executar_em_lotes`) e preâmbulo da IA (`_preparar_ia`).
O prompt `corrigir-planilhas` perdeu a frase sobre alterar a operação (instalações existentes: "Restaurar semente").
**Fecha o conjunto ADR-004 (TASK-009 a TASK-015).**

---

## 2026-09-30 — TASK-014: validação na tela (botão, opção automática) e revisão por IA (camada 3)

**Tipo:** nova funcionalidade (backend + frontend) · `.ai/tasks/TASK-014-30-09-2026.md`

Botão **Validar** e opção **Validar ao montar** (radio Não/Sim, `localStorage`, padrão Não) na aba Resumo, com painel
de achados das 3 camadas e o diálogo "continuar mesmo assim?" (só erro/aviso interrompem). `POST /api/validacao/ia`
usa o prompt salvo (TASK-012), lotes de 40 linhas, resposta JSON validada (descarta linha inventada) e nunca derruba
as camadas 1-2: falha da IA vira aviso. Refatoração mínima: `obterPayloadCalculo()` extraída do handler de "Montar
Orçamento" (comportamento igual). `localStorage['processar_dados']` e `orcamentoPayload` inalterados. O botão
"Não, vou corrigir" ainda só fecha o painel (a correção é a TASK-015).

---

## 2026-09-30 — TASK-013: regras de domínio da validação (camada 2), editáveis pelo admin

**Tipo:** nova funcionalidade (backend, banco, painel admin) · `.ai/tasks/TASK-013-30-09-2026.md`

Motor `services/regras_dominio.py` com 5 tipos de regra declarativa, avaliadas por linha da tabela Outros
(P50 pelo total da planilha). Regras por projeto com fallback `DEFAULT`, histórico e reversão, painel de teste
do rascunho e editor no painel admin. Semente com 9 regras derivadas do `prompt_rede_eletrica.txt §5`
(intocado), **todas desligadas**; valem só para operações I/*I. As 3 regras de P50 do catálogo viraram uma
(a exigência é uma soma). `POST /api/validacao/planilhas` passa a incluir a camada 2.
**Ação necessária:** rodar o trecho novo de `scripts/schema_supabase.sql` no Supabase real.

---

## 2026-09-30 — TASK-012: prompts de validação editáveis pelo admin

**Tipo:** nova funcionalidade (backend, banco, painel admin) · `.ai/tasks/TASK-012-30-09-2026.md`

Prompts de validação (camada 3 do ADR-004) viram dados: semente em `data/validacoes/*.md`, tabelas
`prompts_validacao`/`prompts_validacao_historico` (SQLite + Supabase), rotas `/api/validacao/prompts`
(leitura autenticada, escrita admin), validação de cabeçalho e placeholders antes de salvar, histórico
com reversão e restauração da semente. Versão por projeto com fallback `DEFAULT`. UI no painel admin.
**Ação necessária:** rodar o trecho novo de `scripts/schema_supabase.sql` no Supabase real. Contratos da
Regra 5: nenhum alterado; `prompt_rede_eletrica.txt` intacto.

---

## 2026-09-30 — TASK-011: camada 1 da validação das planilhas (primeiros testes do projeto)

**Tipo:** nova funcionalidade (backend) + refatoração mínima · `.ai/tasks/TASK-011-30-09-2026.md`

`POST /api/validacao/planilhas` valida o contrato de Cabos e Outros (formato, operação, quantidade,
duplicidade e, opcionalmente, existência do ativo na base sobre o payload de cálculo). Para reaproveitar
o parser sem duplicá-lo, `orcamento_calc.py` teve o parsing por linha extraído em
`extrair_ativo_cabo`/`extrair_ativos_outros`/`tokenizar_outros` — saída do cálculo comprovadamente
idêntica. **Achado:** a tabela Cabos da tela (`CAA 2 ABC 35 m`) e o payload de cálculo (`CAA2 1 35`) têm
formatos diferentes; a validação de Cabos espelha a normalização do frontend (nova duplicação
conhecida). `ADR-004` passa a ACEITA. Suíte `pytest` criada (`tests/`, dependências em `requirements-dev.txt`).

---

## 2026-09-30 — TASK-009: chave de IA padrão do sistema

**Tipo:** nova funcionalidade + correção (backend/frontend/config) · `.ai/tasks/TASK-009-30-09-2026.md`

O chat de IA passa a funcionar sem o usuário digitar chave. `resolver_credencial()` define a
precedência **usuário > salva > padrão** (`GEMINI_API_KEY`/`GOOGLE_API_KEY`). Corrige um bug: a UI
enviava o texto `SAVED_IN_BACKEND` quando o campo estava vazio, e por ser "verdadeiro" em Python ele
impedia que a variável de ambiente fosse consultada — resultado 401 mesmo com chave no servidor.
No desktop, `config.py` carrega o `.env` ao lado do `.exe`. Rate-limit por usuário só com a chave
padrão. `/api/health` ganha `ai_key_source`. Contratos da Regra 5: nenhum alterado.

## 2026-09-30 — TASK-010 / ADR-004: base do sistema de validação das planilhas

**Tipo:** documentação e sementes (sem lógica) · `.ai/tasks/TASK-010-30-09-2026.md`

Decisão de validar Cabos e Outros em três camadas (contrato em código, regras de domínio editáveis
pelo admin, IA com prompts salvos editáveis) — `ADR-004` (PROPOSTA). Criados o catálogo
`.ai/VALIDACOES.md`, os prompts iniciais em `data/validacoes/` e a semente das regras de domínio,
todas inativas até confirmação. `prompt_rede_eletrica.txt` não foi alterado. Achado documentado: o
parser de `orcamento_calc.py` assume valores e descarta tokens em silêncio.

---

## 2026-09-29 — Correção: regra de CERCA ("FIOS") sempre saía com operação R quando a cor era cinza

**Tipo:** correção de dado (seed) · fora do escopo de uma TASK, reportado pelo usuário via edição
manual no admin

A regra de Classificação que atribui `entidade: CERCA` a partir do texto "FIOS"
(`data/regras_leitor_classificacao_seed.json`, ordem 40) estava sem o campo
`operacao_ajustada: "I"` — o mesmo campo que todas as regras irmãs (RAMAIS, APOIO, IP) usam para
forçar a operação. Sem ele, a operação de uma linha de cerca ficava com o que a Tabela de
Processamento decidisse (a regra de "FIOS" lá também não define operação), caindo no fallback por
cor (`operacaoPelaCor`: vermelho→I, **cinza→R**, resto→M) — cercas desenhadas em cinza saíam como
"R" em vez de sempre "I". A própria `.ai/tasks/TASK-006-25-09-2026.md` já documentava que CERCA
deveria forçar operação I junto com RAMAIS/APOIO/IP; ficou de fora só dessa regra no seed, uma
omissão de transcrição.

Corrigido o seed (`operacao_ajustada: "I"` adicionado). O usuário já havia corrigido o dado no
banco em produção diretamente pelo editor de regras (TASK-007) antes desta correção do seed —
esta mudança só evita que o mesmo bug volte a aparecer em uma instalação nova/reset do banco.

---

## 2026-09-29 — TASK-007: regras do leitor ficam editáveis pela UI (admin)

**Tipo:** nova funcionalidade (backend + frontend) · **Tarefa:** `.ai/tasks/TASK-007-29-09-2026.md`

O modal "Regras do Leitor" (TASK-006, antes só visualização) ganha edição completa das duas
tabelas de regras, restrita a usuários com role `admin`: formulário estruturado por regra (campos
condicionais por modo — `DEFINIR`/`SUBSTITUIR`/`SUBSTITUIR_TOTAL` — e fase), reordenação por
drag-and-drop, painel de teste que roda `RegrasLeitorEngine` no navegador contra o conjunto de
regras em edição (ainda não salvas, sem round-trip ao backend), validação de schema no backend
antes de salvar (toda regex precisa compilar, campos obrigatórios por modo — nunca aceita um
payload que quebraria o motor), e histórico de versões com reversão (cada salvamento preserva a
versão anterior; reverter também vira uma nova entrada de histórico, a cadeia nunca perde uma
versão). Rotas `POST` protegidas por `Depends(require_role("admin"))`; `GET` continua aberto a
qualquer usuário autenticado.

**Correção de planejamento descoberta na implementação:** o refinamento inicial da tarefa assumia
que a ordem de execução das regras era a posição delas no array/JSON — na verdade o motor
(`regras_leitor_engine.js`) ordena por um campo `ordem` explícito dentro de cada fase. O
drag-and-drop reescreve esse campo (múltiplos de 10, mesmo espaçamento do seed) em vez de só
reordenar o array.

Novas tabelas `regras_leitor_processamento_historico`/`_classificacao_historico` (SQLite +
Supabase). Duplicação de lógica registrada como risco conhecido: a validação de schema do backend
precisa se manter em sincronia com o que `regras_leitor_engine.js` de fato interpreta.

---

## 2026-09-29 — TASK-008: painel "Resumo da rede" na aba Resumo

**Tipo:** nova funcionalidade (frontend) · **Tarefa:** `.ai/tasks/TASK-008-29-09-2026.md`

Adicionado um painel com 4 contagens em tempo real na aba Resumo, logo abaixo da barra de botões
dos modais geradores (Postes/Cabos/Conexões/Ramais/...): **postes instalando** (soma de quantidade
dos tokens `qtd-ativo` da tabela Outros que começam com `DT`/`CV`), **rede de média instalando**
(soma do comprimento bruto dos Cabos que começam com `CAA`/`P`/`CAL`), **rede de baixa instalando**
(idem para `M2X`/`M3X`) e **equipamentos instalando** (soma de quantidade dos tokens Outros que
começam com `TR`/`CFU`/`CFA`). Todas as 4 consideram só linhas com operação `I`.

**Duplicação de parsing (risco registrado):** `static/resumo.js` ganhou uma porta em JavaScript do
parsing `qtd-ativo` (Outros) e `ATIVO FASE COMPRIMENTO` (Cabos) que hoje só existia em Python
(`services/orcamento_calc.py`), necessária para recalcular em tempo real no navegador sem round-trip
ao backend. **O comprimento usado nunca é multiplicado pela quantidade de ativos/fases** —
diferente do cálculo de orçamento, que multiplica — decisão explícita do usuário.

Não persiste em banco — é só um cálculo em cima de `tableStates.cabos.data`/`outros.data` já
carregados na tela, atualizado a cada render e a cada edição direta (mudança de operação/entidade,
edição do texto do ativo).

Junto: registrada (não implementada) `.ai/tasks/TASK-007-29-09-2026.md` — tornar as regras do
leitor (TASK-006) editáveis via UI para usuários `admin`, em stand-by até refinamento posterior.

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
