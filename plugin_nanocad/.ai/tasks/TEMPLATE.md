# TASK-XXX

> Copie para `.ai/tasks/TASK-XXX-dd-mm-aaaa.md` (numeração própria do plugin) e preencha.
> Não apague tarefas concluídas — mude o Status para `CONCLUÍDA`.
> Só o específico da tarefa: o contexto está em `CONTEXT.md`, `ARCHITECTURE.md`, `REQUISITOS.md` e `GLOSSARY.md`.

## Objetivo

Uma ou duas frases.

## Requisitos tocados

`RQ-xx` de `REQUISITOS.md`.

## Situação atual

O que acontece hoje (arquivo e linha quando souber).

## Comportamento desejado

## Restrições

- [ ] Contrato do `.obra.json` (`RULES.md` P4) — se mudar, atualizar exemplo + teste de contrato + aplicação
- [ ] Paridade com a aplicação (`RULES.md` P3) — se tocar o motor/anotações/vínculo
- [ ] C# 5 / .NET 4 (`RULES.md` P2)
- [ ] Alteração no desenho do usuário (`RULES.md` P5)
- [ ] Nenhum dos acima

## Critérios de aceite

- [ ] …
- [ ] `dotnet test tests/Nucleo.Testes` e `pytest tests/test_plugin_nanocad_contrato.py` passando
- [ ] DLL do nanoCAD 26 recompilada (`./compilar.sh`)
- [ ] `STATE.md`, `CHANGELOG.md` (e `REQUISITOS.md`, se for o caso) atualizados

## Verificação no nanoCAD (pelo usuário)

Passos curtos que o usuário executa no CAD e o que deve ver.

## Status

`PLANEJAMENTO` · `EM ANDAMENTO` · `AGUARDANDO TESTE NO CAD` · `BLOQUEADA` · `CONCLUÍDA` · `CANCELADA`

**Atual:** PLANEJAMENTO
