# ADR-006 — Motor de ajustes determinísticos das planilhas (ações declarativas com diff)

## Contexto

A validação (ADR-004/005) só **aponta** problemas; a correção existente é só a IA (`correcao_ia.py`), que troca o texto
`ativo` de uma linha. O usuário pediu "validação + ajuste": a correção feita **direto nas tabelas de entrada**, com
funções recorrentes cadastráveis (substituição de texto, ordenação, exclusão/adição de linhas). Decisões (2026-10-01):
a operação da linha pode mudar **só se determinado explicitamente**; a lista de 7 ações está completa; **nada é aplicado
sem pré-visualização**.

## Problema

Como transformar as planilhas de forma previsível, testável e segura (sem quebrar o contrato que o cálculo lê), de modo
que as mesmas ações sirvam ao ajuste manual, às receitas cadastradas (TASK-023) e às propostas da IA (TASK-025).

## Alternativas consideradas

### 1 — Código de UI por ajuste (como o chat da IA faz hoje)
- Prós: nada novo no backend. Contras: sem schema, sem teste, sem diff, regra de negócio duplicada no frontend.

### 2 — Linguagem livre (script/expressão)
- Prós: poder total. Contras: execução arbitrária, difícil de validar e de pré-visualizar.

### 3 — Ações declarativas em JSON, motor puro no backend que devolve **diff** (escolhida)
- Prós: validável, descritível em português, reutiliza condições/seletores/grupos da linguagem de regras v2, testável
  (idempotência, mutação) e serve a regras, receitas e IA com o mesmo formato.
- Contras: mais uma linguagem para manter (mitigado compartilhando condições e seletores).

## Decisão

`services/ajustes_planilhas.py` (função pura). **7 ações:** `substituir`, `normalizar`, `ordenar`, `excluir_linhas`,
`adicionar_linha`, `adicionar_ativo`/`remover_ativo`, `mesclar_duplicadas` (detalhe no cabeçalho do módulo).
- Entrada: lista de ações + tabelas Cabos/Outros (+ grupos efetivos do projeto). **Não altera a entrada.**
- Saída: `operacoes` (`editar`/`inserir`/`excluir`/`mover`, por id de linha, com `antes/depois` e as ações que as causaram),
  as tabelas finais com `id`, `descartadas`, `avisos` e `resumo`. O **frontend aplica** (um passo de histórico) só depois do aceite.
- **Guarda da camada 1:** edição/inserção que deixe a linha com um **erro de contrato novo** é descartada e listada.
- **Operação da linha:** mantida; só muda com `operacao_nova` (e só nas linhas que a ação de fato alterou).
- **Idempotência:** reaplicar o resultado não gera operações (testado).
- `ordenar` em Cabos devolve **aviso** (linha de cabo sem fase herda a da seguinte, RN-04).
- Rota `POST /api/validacao/ajustes/preview` (qualquer autenticado; só simula) e `/descrever`. Cadastro/persistência: TASK-023.
- Sem mudança de banco/Supabase nesta etapa.

## Consequências

### Positivas
- Um único formato de ação para ajuste manual, receitas e IA; diff para pré-visualização; reaproveita o motor de regras.

### Negativas / custos aceitos
- Linguagem própria adicional; ações de texto operam sobre o texto do ativo (tokens separados por espaço), então
  `remover_ativo`/`mesclar_duplicadas` normalizam espaços da linha.

### Implicações para implementações futuras
- Nova ação = nova entrada em `ESPECIFICOS` + validação + teste de idempotência/mutação.
- Nunca aplicar diff sem pré-visualização e sem passar pela camada 1.

## Evidência no código

`services/ajustes_planilhas.py`, `routers/validacao_ajustes.py`, `tests/test_ajustes_planilhas.py`.

## Data

2026-10-01

## Status

`PROPOSTA` · `ACEITA` · `SUPERADA POR ADR-YYY` · `DEPRECIADA`

**Atual:** ACEITA
