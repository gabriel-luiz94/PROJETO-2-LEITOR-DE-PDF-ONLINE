# ADR-003 — Motor de regras de conversão no frontend, entre a classificação e o cálculo

> ⚠️ **ADR retroativo.** Decisão já implementada quando a documentação `.ai` foi criada
> (2026-09-18), reconstruída a partir do código. Sem registro histórico da deliberação original.

## Contexto

Entre o que está desenhado no projeto e o que entra no cálculo do orçamento existe um conjunto de
ajustes que **variam por concessionária e mudam com frequência**. Exemplos visíveis no código:

- margem de perdas em cabos — a regra default criada pelo sistema é
  `origem=CABOS, acao=ADICAO, fator=1.05` (`resumo.js:2547-2551`);
- substituição de um ativo por outro equivalente conforme o projeto;
- arredondamento e limites mínimo/máximo de quantidade.

Um comentário em `services/orcamento_calc.py:93` registra a intenção explicitamente:
*"A margem de perdas (ex: +5%) agora é gerida pela Tabela de Regras no frontend."* — o "agora"
indica que essa lógica já esteve em outro lugar e foi deliberadamente movida.

## Problema

Onde aplicar ajustes que (a) variam por concessionária, (b) mudam com frequência, (c) precisam ser
inspecionados e corrigidos pelo próprio usuário antes de virarem orçamento.

## Alternativas consideradas

`[NÃO DETERMINADO PELO CÓDIGO]` — não há registro das alternativas avaliadas. O código evidencia
apenas que a margem de perdas **esteve embutida no cálculo** e foi movida para a tabela de regras.

## Decisão

**O motor de regras roda no frontend, sobre a camada unificada, antes de enviar o payload ao
backend.** As regras são dados editáveis pelo usuário, não código.

Posição no pipeline:

```
CABOS + OUTROS → camada unificada (rawItems) → MOTOR DE REGRAS → TOTALIZADORA (editável)
                                                                         ↓
                                                        POST /api/orcamento/calcular
```

Estrutura de uma regra (10 campos):
`ORIGEM · OP_DE · ATIVO_DE · ACAO · OP_PARA · ATIVO_PARA · FATOR · ARREDONDAMENTO · VAL_MIN · VAL_MAX`

Semântica central (`resumo.js:2918-3006`):
- campo de match vazio = "qualquer"; `%` = padrão LIKE; senão regex ancorada com `^…$`
- `ADIÇÃO` → gera o item transformado **e mantém o original**
- `SUBST` → gera o transformado e **remove o original**
- todas as regras que casam são aplicadas, cada uma gerando sua própria saída

Persistência por projeto, com nuvem tendo precedência na leitura e `localStorage` como fallback:
`configuracoes['regras_conversao_<projeto_codigo>']` ↔ `localStorage['regras_orcamento_<projCode>']`.

## Consequências

### Positivas
- O usuário ajusta as regras sem deploy e sem alterar código
- Regras diferentes por concessionária sem ramificação no backend
- **O resultado da aplicação das regras é visível e editável** na Totalizadora antes do cálculo —
  o usuário confere e corrige o que a regra fez
- `services/orcamento_calc.py` permanece uma função pura de cálculo, sem políticas variáveis
- Import/export em CSV permite versionar e compartilhar conjuntos de regras

### Negativas / custos aceitos
- **Regra de negócio fora do backend:** qualquer cliente que chame `POST /api/orcamento/calcular`
  diretamente pula o motor de regras inteiro. O endpoint, inclusive, é público
- **Não auditável:** não há registro de quais regras foram aplicadas a um REC salvo. Reproduzir um
  orçamento antigo exige saber quais regras vigiam naquele momento
- **Regex fornecida pelo usuário executada no navegador:** `new RegExp` sobre campo livre
  (`resumo.js:2946`); há `try/catch`, mas nada impede uma expressão patológica
- Regras ficam em `localStorage` quando a nuvem falha — limpar o navegador perde ajustes locais
- A ordem de aplicação é a ordem da tabela, sem precedência declarada; regras que se sobrepõem
  produzem múltiplas saídas

### Implicações para implementações futuras
1. **Não replique a margem de perdas (ou qualquer ajuste equivalente) no backend.** Ela foi
   deliberadamente retirada de lá. Reintroduzi-la aplicaria o fator duas vezes.
2. **Ajuste que varia por concessionária é regra de dados, não código.** Antes de escrever um `if`
   por projeto no cálculo, verifique se o caso não é expressável na tabela de regras.
3. **Mudar a semântica de `ADIÇÃO`/`SUBST` quebra as regras já salvas** dos usuários, em nuvem e em
   `localStorage`. É mudança destrutiva.
4. Mover o motor para o backend é uma decisão arquitetural — exige novo ADR superando este, e um
   plano de migração das regras já persistidas.
5. Ao alterar os 10 campos da regra, atualize junto: o render da tabela, `updateRegraRow`, o
   import e o export CSV, e o default criado em `carregarRegras`.

## Evidência no código

- `resumo.js:2918-3006` — motor de regras (match, ação, fator, arredondamento, limites)
- `resumo.js:2816-2916` — construção da camada unificada (`rawItems`, `baseId`, `origem`)
- `resumo.js:2515-2559` — carregamento com precedência nuvem → `localStorage` → default
- `resumo.js:2614-2686` — export e import CSV das regras
- `routers/regras.py:34-69` — persistência por projeto em `configuracoes`
- `services/orcamento_calc.py:92-94` — comentário registrando a saída da margem de perdas do backend
- `resumo.js:1833-1886` — payload montado **a partir da Totalizadora**, já com as regras aplicadas

## Data

2026-09-18 (documentação retroativa; data da decisão original desconhecida)

## Status

**ACEITA**
