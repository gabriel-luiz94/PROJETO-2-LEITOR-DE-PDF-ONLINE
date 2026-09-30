# ADR-005 — Linguagem de regras de domínio v2 e regras em camadas por projeto

## Contexto

A camada 2 do ADR-004 (TASK-013) guardava cada regra como `tipo` + `parametros` com **regex sobre o nome do ativo**.
Limites reais encontrados: (a) não dava para condicionar pela **quantidade** (pedido do usuário: "só quando for `1-CFU`");
(b) a lista de estruturas de média tensão estava **copiada** em duas regras, com regex de 80+ caracteres; (c) o resultado
não explicava **por que** disparou; (d) a versão de um projeto era uma **cópia integral** da lista, que parava de receber
as atualizações do padrão. Pedido (2026-09-30): regras "mais legíveis e mais poderosas" e "regras únicas por projeto".

## Problema

Como escrever regras que o admin consiga ler e editar sem regex, com quantidade e combinações lógicas, sem quebrar as
regras já salvas e sem duplicar o que é comum a vários projetos.

## Alternativas consideradas

### 1 — Só acrescentar `se_qtd` ao tipo `requer`
- Prós: mudança mínima. Contras: resolve um caso; regex e duplicação continuam; cada necessidade nova vira um campo novo.

### 2 — Linguagem de expressões livre (mini-DSL em texto ou JS)
- Prós: máxima liberdade. Contras: difícil de validar e de editar em formulário; risco de execução de código.

### 3 — Linguagem **declarativa em JSON** com condições/seletores/grupos/expressões (escolhida)
- Prós: validável no salvamento, editável em formulário, traduzível para frase em português e para explicação.
  Contras: linguagem própria para manter.

## Decisão

**JSON declarativo v2** (detalhe em `services/regras_dominio.py`, cabeçalho do módulo):
- **Condições** `tem`/`soma` (com comparador de quantidade `=, ≠, ≥, ≤, >, <, entre`), `texto`, e combinadores `todos`,
  `algum`, `nenhum`, `nao`. Regra de **linha**: `quando` → `entao`, com `excecoes`; sem `entao` = "é proibido".
  Regra de **planilha**: `deve_ser {esq, cmp, dir}` sobre somas de quantidades (Outros) e de metros (Cabos).
- **Seletores** de ativo: código, curinga (`TR[0-9]*`), `@GRUPO`, `{regex}` (avançado), `{e: [..]}`, lista (OU).
- **Grupos** nomeados reutilizáveis; **valores dinâmicos** (`base`, `vezes`, `mais`, `se`) para "dobra com duas descidas".
- Cada achado leva `explicacao` em português; `descrever(regra)` gera a frase da regra.
- **Compatibilidade:** regras v1 continuam aceitas; `converter_v1` as traduz de forma determinística. A equivalência é
  provada por teste contra uma **cópia congelada do motor v1** (`tests/referencia_regras_v1.py`), em milhares de cenários.
- **Armazenamento:** o mesmo `regras_json` (TEXT), agora um container `{"versao": 2, "grupos": {...}, "regras": [...]}`;
  o formato antigo (lista) continua sendo lido. **Sem SQL novo no Supabase.**
- **Camadas por projeto** (TASK-018): o padrão (`DEFAULT`) guarda regras e grupos completos; um projeto guarda só o
  **overlay** (`adicionadas`, `sobrescritas`, `ocultas`, `grupos`). Vale a do projeto quando ele sobrescreve; o padrão que
  muda chega aos projetos; "−/+" oculta/reexibe uma regra herdada naquele projeto. Cópias integrais antigas viram
  overlay equivalente (o que faltava no padrão vira "oculta"), sem mudar nenhum resultado.

## Consequências

### Positivas
- Regras legíveis (frase + explicação), sem regex para o caso comum; ajuste de um grupo vale para todas as regras.
- Regras por projeto sem copiar a lista e sem perder atualizações do padrão.
- Nenhuma regra já salva quebra; nenhum SQL novo.

### Negativas / custos aceitos
- Linguagem própria para documentar e manter; regex continua existindo como recurso avançado.
- Overlay adiciona um nível de indireção: o "efetivo" precisa ser calculado a cada leitura.

### Implicações para implementações futuras
- Nova capacidade = novo nó/comparador em `services/regras_dominio.py` **com** teste de equivalência quando tocar regras existentes.
- Nunca mover regra de **contrato** (camada 1) para dados editáveis (ADR-004).
- Ao mudar a semente, atualizar também `tests/fixtures/regras_dominio_seed_v1.json` só se o oráculo v1 precisar mudar (não deve).

## Evidência no código

`services/regras_dominio.py`, `services/regras_camadas.py`, `routers/validacao_regras.py`,
`data/validacoes/regras_dominio_seed.json`, `tests/test_regras_v2.py`, `tests/referencia_regras_v1.py`.

## Data

2026-09-30

## Status

`PROPOSTA` · `ACEITA` · `SUPERADA POR ADR-YYY` · `DEPRECIADA`

**Atual:** ACEITA (2026-09-30; a parte de camadas por projeto é entregue na TASK-018)
