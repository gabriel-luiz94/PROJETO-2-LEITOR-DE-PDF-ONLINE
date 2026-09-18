# TASK-XXX

> Copie este arquivo para `.ai/tasks/TASK-XXX-nome-curto.md` e preencha.
> Numeração sequencial. Não apague tarefas concluídas — mude o Status para `CONCLUÍDA`.
>
> Este arquivo deve conter **apenas o específico da tarefa**. Tudo que é contexto geral do projeto
> já está em `.ai/CONTEXT.md`, `.ai/ARCHITECTURE.md` e `.ai/GLOSSARY.md` — não repita aqui.

## Objetivo

Descrever claramente o que deve ser alterado. Uma ou duas frases.

## Contexto

Explicar somente o contexto específico necessário para esta tarefa.

Não reexplique o projeto. Se precisar apontar algo já documentado, referencie:
`CONTEXT.md §7 RN-06`, `ARCHITECTURE.md §8.7`, `GLOSSARY.md › COMPONENTE`.

## Situação atual

O que acontece hoje? Cite arquivo e linha quando souber (`services/orcamento_calc.py:187`).

## Comportamento desejado

O que deve acontecer depois da alteração?

## Restrições

O que **não** deve ser alterado.

Considere explicitamente os contratos listados em `RULES.md` (Regra 5) que esta tarefa toca:
- [ ] `localStorage['processar_dados']` (script.js → resumo.js)
- [ ] `localStorage['orcamentoPayload']` (resumo.js → resultado_orcamento.html)
- [ ] Payload de `POST /api/orcamento/calcular`
- [ ] Formato de ativo `[ATIVO] [FASE] [COMPRIMENTO]` (CABOS)
- [ ] Formato de ativo `<qtd>-<ativo>` (OUTROS)
- [ ] Colunas do CSV da base técnica
- [ ] Colunas do CSV de regras de conversão
- [ ] Schema do Supabase
- [ ] Nenhum dos acima

## Critérios de aceite

Condições objetivas e verificáveis que determinam se a tarefa está concluída.

- [ ] …
- [ ] …
- [ ] Documentação `.ai` atualizada conforme `RULES.md` Regra 9
- [ ] `STATE.md` atualizado conforme `RULES.md` Regra 10

## Exemplos

### Entrada

```
```

### Resultado esperado

```
```

> Para regras de negócio, use o formato entrada → regra → saída de `ARCHITECTURE.md §8`.

## Arquivos potencialmente afetados

- `caminho/arquivo.py:linha` — o que muda

⚠️ Se a tarefa toca regra de classificação de entidade/operação/ativo, verifique **as duas** cópias:
`static/script.js` e `static/resumo.js` (RULES.md, Regra 4).

## Riscos

O que pode quebrar. Consulte `ARCHITECTURE.md §10` (dependências entre módulos).

## Plano de teste

Como a alteração será verificada. O projeto não tem testes automatizados (`STATE.md`), então
descreva a verificação manual ou os testes que serão criados.

## Status

`PLANEJAMENTO` · `EM ANDAMENTO` · `BLOQUEADA` · `CONCLUÍDA` · `CANCELADA`

**Atual:** PLANEJAMENTO
