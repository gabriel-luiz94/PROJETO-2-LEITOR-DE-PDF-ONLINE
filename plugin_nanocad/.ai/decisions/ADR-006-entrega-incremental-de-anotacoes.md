# ADR-006 — Entrega incremental de anotações

## Decisão

A montagem manual entrega um lote das tabelas ao Resumo autenticado da instalação local. O plugin localiza o processo e lê uma credencial do registro por usuário. O Resumo verifica a obra aberta e o projeto antes de acrescentar o lote ao histórico das tabelas. O aplicativo confirma o estado e a confirmação pode ser consultada pelo plugin.

## Consequências

O destino continua vinculado ao arquivo de trabalho atual. Os índices de vínculo do lote são deslocados pela quantidade de linhas já existentes em Outros. Lotes confirmados são idempotentes durante a execução do programa. Os pedidos de abertura de arquivos .obra seguem o contrato do ADR-005. A entrega incremental precisa de uma versão do aplicativo com suporte à TASK-218.

## Estado

Aceita em 22/11/2026. [TASK-036](../tasks/TASK-036-22-11-2026.md).
