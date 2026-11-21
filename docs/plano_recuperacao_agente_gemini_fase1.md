# Fase 1 — Recuperação do agente Gemini no OrçaCore

## Objetivo
Recuperar execução do assistente na tela Resumo, com Gemini como único provedor exigido nesta fase. Não alterar a branch main até validação.

## Código existente confirmado
- routers/assistente.py: POST /api/assistente/chat, contexto do projeto, skills, NDJSON.
- services/assistente/agente.py: adaptadores, tool calling, timeout de chamada (45s) e prazo total (120s).
- services/assistente/ferramentas.py: buscar_ativo, detalhar_ativo, ler_tabela, validar_linhas, propor_alteracoes, ajustar_tela.
- services/assistente/skills.py: skills gerais e por projeto, incluindo carregar_skill.
- static/assistente.js: exibe/aplica plano após confirmação, com desfazer.

## Escopo e critérios de aceite
1. **Gemini**: comando gera eventos de ferramenta, resposta ou erro compreensível; jamais fica indefinidamente em "Gerando". Testar sucesso, 429, 503, timeout, chave inválida, stream interrompido e erro em ferramenta. Garantir evento terminal e liberação do botão.
2. **Consulta por projeto**: "Como cadastrar uma N3 em Pernambuco?" usa projeto selecionado, busca_ativo/detalhar_ativo e base real. Explica passos e exemplo com dados reais; nunca inventa códigos, componentes ou parâmetros. Distingue não encontrado de falha técnica.
3. **Adicionar/excluir/editar**: plano validado, exibição antes da aplicação, confirmação e desfazer. Respeitar linhas congeladas/bloqueadas, vínculos e duplicidades; não editar o estado diretamente no backend.
4. **Ordenar**: "ordene Outros por ativo" ordena visualmente a tabela correta, sem romper IDs e vínculos; testar asc/desc.
5. **Agrupar**: reutilizar a semântica de grupos do Resumo (não criar formato novo), com títulos e limites bem definidos, preservar vínculos e permitir desfazer.
6. **Comentar cada linha**: inserir UMA linha de comentário separada e próxima do ativo, usando o formato de comentário já existente no Resumo. Evitar duplicação em execução repetida; não mudar quantitativos.
7. **Montar rede simples**: consultar skill montar_projeto e regras específicas, buscar ativos existentes no projeto, montar proposta separada em Cabos/Outros com postes/estruturas; explicitar hipóteses e dados faltantes; validar candidatos pela Totalizadora; não tratar proposta como projeto aprovado.
8. **Proteção**: nenhum plano destrutivo é aplicado sem confirmação. Desfazer recupera estado anterior. Mudanças devem preservar integridade dos vínculos.

## Sequência de implementação
A. Reproduzir erro com logs/testes de adaptador Gemini e streaming NDJSON; corrigir causa observada, não apenas aumentar timeout.
B. Testes de integração buscar_ativo/detalhar_ativo por projeto.
C. Testes de UI para ordenação e operações do plano.
D. Implementar grupos/comentários por meio das rotinas existentes do Resumo.
E. Montagem assistida via skills + consulta de base + validação + plano.
F. Rodar suíte Python/Node e testar manualmente no Resumo com dados descartáveis.

## Observações técnicas
- O executor atual rejeita texto iniciado por @ ou # em propor_alteracoes: comentário/grupo precisa de caminho próprio, não remover essa proteção indiscriminadamente.
- A rota de chat convencional /api/gemini/chat não é a rota do agente; investigar /api/assistente/chat.
- Confirmar no frontend onde os eventos acao_ui são efetivamente consumidos antes de afirmar que ordenação funciona.
- Preservar a branch main. Não fazer merge sem revisão e testes.

## Comandos de teste
- "Pesquise N3 na base de Pernambuco e explique como cadastrar, com exemplo."
- "Adicione 2 N3 na tabela Outros com operação I."
- "Remova o ativo da linha selecionada."
- "Ordene a tabela Outros por ativo em ordem crescente."
- "Agrupe os ativos por entidade e crie títulos de grupo."
- "Comente cada linha da tabela Outros em linhas separadas."
- "Monte uma proposta de rede simples com postes, estruturas e cabos usando as skills de Pernambuco."
