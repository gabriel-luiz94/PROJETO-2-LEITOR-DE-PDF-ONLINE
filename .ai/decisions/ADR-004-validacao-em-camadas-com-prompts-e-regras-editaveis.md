# ADR-004 — Validação das planilhas de entrada em camadas, com regras e prompts editáveis pelo admin

## Contexto

As tabelas **Cabos** e **Outros** (planilhas de entrada do orçamento) são preenchidas por leitura de
PDF/DXF, pelos modais geradores, pelo chat de IA e por edição manual. Hoje nada as valida antes de
`POST /api/orcamento/calcular`: o único retorno é `nao_encontrados` (ativo ausente na base técnica).
As regras técnicas de domínio existem só como instrução ao modelo em `prompt_rede_eletrica.txt` e
não são verificadas por código (`STATE.md › Limitações`).

Também: o chat de IA só funciona com chave digitada pelo usuário (TASK-009 corrige).

## Problema

Como validar, checar e corrigir as planilhas de entrada de forma barata, previsível e mantida pelo
próprio usuário, sem deploy a cada ajuste de regra.

## Alternativas consideradas

### Alternativa 1 — Tudo por IA (um prompt grande)
- Prós: mais simples de construir.
- Contras: caro, não determinístico, não reproduzível, repete erros que o código detecta exato.

### Alternativa 2 — Tudo em código
- Prós: exato e testável.
- Contras: regra de domínio muda por concessionária/projeto; exigiria deploy; não cobre coerência semântica.

### Alternativa 3 — Camadas: contrato em código + regras de domínio como dados + IA para o restante
- Prós: cada camada faz o que faz melhor; regras editáveis sem deploy (mesmo padrão do ADR-003 e da TASK-007).
- Contras: três mecanismos para manter.

## Decisão

Alternativa 3. **Confirmado pelo usuário em 2026-09-30.**

1. **Camada 1 — Contrato (código, NÃO editável).** Formatos de ativo (CABOS `[ATIVO] [FASE] [COMPRIMENTO] m`,
   OUTROS `<qtd>-<ativo>`), tabela errada para o formato, quantidade zero/vazia, operação inválida,
   ativo inexistente na base técnica. Estes formatos são contratos lidos por `orcamento_calc.py`
   (RULES Regra 5); permitir editá-los na UI quebraria o cálculo.
2. **Camada 2 — Regras de domínio (dados, EDITÁVEIS pelo admin).** Regras declarativas por projeto
   (ex.: "CFU exige SUPL"), com severidade, ativa/inativa, histórico de versões, painel de teste e
   validação de schema — mesmo padrão de `routers/regras_leitor.py`. Semente inicial montada a partir de
   um **prompt/catálogo próprio**, sem alterar `prompt_rede_eletrica.txt`.
3. **Camada 3 — IA (prompts salvos, EDITÁVEIS pelo admin).** Prompts em tabela com histórico, semente em
   arquivo do repositório. Recebem os achados das camadas 1 e 2 como entrada e devolvem JSON de schema
   fixo. Correções são **propostas** (diff por linha), aceitas pelo usuário, passando pelo undo/redo.
4. **Gatilho:** botão sob demanda + opção (radio) "validação automática ao montar orçamento",
   preferência local (`localStorage`), desligada por padrão. Com achados, pergunta "continuar mesmo
   assim?"; se o usuário recusar, entra o fluxo de correção.
5. **Chave de IA padrão:** variável de ambiente no servidor; no desktop, lida de `.env` ao lado do `.exe`.

## Consequências

### Positivas
- Erros determinísticos nunca gastam token nem variam entre execuções.
- Admin ajusta regras e prompts sem deploy, com reversão.
- `prompt_rede_eletrica.txt` permanece intacto.

### Negativas / custos aceitos
- Editar regra de domínio errada pode gerar falsos alertas ou silenciar alertas reais: mitigado por
  histórico/reversão, painel de teste e severidade.
- Conteúdo das planilhas sai para API externa quando a camada 3 roda (privacidade); não vai para log.

### Implicações para implementações futuras
- Nunca mover regras de **contrato** (camada 1) para dados editáveis.
- Não criar uma terceira cópia de lógica de classificação (RULES Regra 4): o validador consome o
  parsing existente ou o centraliza em um módulo único.
- Correção nunca é aplicada sem aceite explícito do usuário.

## Evidência no código

Ainda não implementado. Pontos de encaixe: `routers/ai_chat.py`, `routers/regras_leitor.py` (padrão de
seed + histórico + validação de schema), `services/orcamento_calc.py:46-122` (formatos de ativo).

## Data

2026-09-30

## Status

`PROPOSTA` · `ACEITA` · `SUPERADA POR ADR-YYY` · `DEPRECIADA`

**Atual:** PROPOSTA (vira ACEITA quando a TASK-011 for concluída)
