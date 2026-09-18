# PROMPT_INICIAL

> Prompt para colar no início de uma conversa com outra IA, para que ela continue o
> desenvolvimento deste projeto seguindo o mesmo protocolo já estabelecido em `.ai/`.
> Não é uma tarefa (`TASK-XXX`) — é um texto de bootstrap, reutilizável a cada nova sessão
> com um agente que ainda não tem contexto deste projeto.

---

```
Você vai continuar o desenvolvimento do projeto "Leitor de Projetos Online Pro"
(sistema de leitura de PDF/DXF + orçamento de redes elétricas, FastAPI + SQLite/Supabase).

Este projeto usa uma estrutura de documentação para IA na pasta `.ai/`. Antes de fazer
qualquer alteração, leia OBRIGATORIAMENTE, nesta ordem:

1. .ai/CONTEXT.md     — o que o sistema é, entradas, saídas, regras de negócio
2. .ai/RULES.md       — regras de trabalho que você deve seguir à risca (protocolo de
                         6 fases, regra de não inventar comportamento, alteração mínima,
                         não duplicar lógica, preservar compatibilidade, etc.)
3. .ai/STATE.md       — o que está pronto, o que está quebrado, problemas conhecidos
4. .ai/ARCHITECTURE.md — módulos, fluxo, banco, API, dependências
5. .ai/GLOSSARY.md    — termos de domínio

Depois leia .ai/CHANGELOG.md para entender o histórico recente de alterações e o
padrão de investigação/correção já usado (veja TASK-001, TASK-002, TASK-003 em
.ai/tasks/ como exemplos reais de diagnóstico → plano → implementação → validação).

Regras que você deve seguir sem exceção (detalhadas em .ai/RULES.md):
- Nunca presuma comportamento ambíguo — pergunte antes de implementar.
- Altere só o necessário para a tarefa pedida; não refatore nem "limpe" código vizinho
  sem pedido explícito.
- Antes de codar, apresente um plano (OBJETIVO / ARQUIVOS AFETADOS / ALTERAÇÕES / RISCOS
  / TESTES) e espere aprovação.
- Depois de implementar, valide de verdade (suba o servidor, teste os endpoints/fluxos
  afetados) — o projeto não tem testes automatizados, então a validação manual é
  obrigatória e deve ser relatada.
- Ao final, atualize .ai/STATE.md e .ai/CHANGELOG.md com o que mudou, e feche o arquivo
  da tarefa em .ai/tasks/.
- Nunca grave credenciais em código/commit; nunca versione dados de usuário (.db, .pdf,
  .dxf já estão no .gitignore).

Minha próxima tarefa é a seguinte (preencha usando .ai/tasks/TEMPLATE.md como modelo —
objetivo, contexto, situação atual, comportamento desejado, restrições, critérios de
aceite, exemplos):

[DESCREVA AQUI A TAREFA]
```
