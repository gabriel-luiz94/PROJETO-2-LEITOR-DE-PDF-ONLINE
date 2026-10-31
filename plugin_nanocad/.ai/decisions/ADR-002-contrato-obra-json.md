# ADR-002 — Contrato do `.obra.json`: envelope leitor-obra v1 + vínculo nas linhas + seção `cad`

## Contexto

O briefing (§4) exige que o arquivo seja aceito por `/api/obras/importar` e leve o vínculo. Conferido no código:
`services/obras_arquivo.normalizar_dados` guardava só `entidade, operacao, ativo, qtdAtivos, texto` por linha — o
vínculo (`vinculoEstruturas`) e a geometria se perdiam na importação (e na exportação).

## Alternativas consideradas

### Formato novo só do plugin
- Contras: segunda porta de entrada na aplicação; Carregar Obra não entenderia.

### Parâmetros explícitos no texto do ativo (`N1(P)`)
- Contras: a ordem dos parâmetros vem da declaração na Base de Orçamento (não disponível no CAD); o vínculo se perderia para a tela de vínculo.

### Mesmo envelope + campos de vínculo que a tela já grava (escolhida)
- Prós: Carregar Obra/Totalizadora/vínculo da aplicação funcionam como numa obra salva pela tela; a exportação passa a levar o vínculo também.

## Decisão

- O plugin grava o envelope `leitor-obra` v1 com as linhas de Cabos/Outros no formato da tela (`resumo.js › deepClone`):
  `entidade, operacao, ativo, textoOriginal, _x, _y, _geo, _geoTipo, _poste, vinculoEstruturas`.
- A aplicação (TASK-103 do principal) aceita esses campos com validação campo a campo (`CAMPOS_VINCULO`).
- Seção `cad` (ignorada pela aplicação) guarda a obra do plugin inteira para Continuar.

## Consequências

- Campo novo por linha precisa existir dos dois lados (`RULES.md` P4).
- O teste de contrato importa o exemplo e compara o vínculo com o da aplicação.

## Evidência no código

`src/Nucleo/ObraArquivo.cs`; `services/obras_arquivo.py › CAMPOS_VINCULO`; `exemplos/Projeto_Rua_das_Flores.obra.json`;
`tests/test_plugin_nanocad_contrato.py`.

## Data

2026-10-31

## Status

**Atual:** ACEITA
