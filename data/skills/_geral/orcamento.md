descricao: Como calcular o orçamento do conjunto de dados da tela (Totalizadora → itens ODI/ODD), conferir ativos não encontrados e comparar cenários.
# Calcular o orçamento

O orçamento é calculado em duas etapas, exatamente como na tela: as tabelas Cabos/Outros passam pela **Totalizadora** (regras de conversão do projeto, cabos expandidos por fase) e o resultado é cruzado com a **Base de Orçamento** para gerar as linhas (MDO, código, descrição, total).

## Fluxo
1. `ler_tabela` para saber o que há na tela (se estiver vazia, avise em vez de calcular).
2. `totalizar` — mostra os ativos finais. Confira `nao_encontrados`: ativo que a base não reconhece NÃO entra no orçamento. Se houver, use `buscar_ativo` para achar o código certo e proponha a correção.
3. `calcular_orcamento` — devolve as linhas do orçamento. Resuma: quantas linhas, os itens principais (por MDO), avisos e ativos não encontrados.
4. Para "e se...": proponha com `propor_alteracoes` e rode `calcular_orcamento` com `com_plano=true`; compare com o cálculo sem o plano e diga a diferença.

## Como responder
- Números sempre das ferramentas. Não some nem arredonde de cabeça.
- Cite o que ficou de fora (não encontrados, avisos de parâmetros) antes de qualquer conclusão: um orçamento com ativos não encontrados está incompleto.
- Se o resultado vier cortado (`cortado: true`), diga que mostrou só os primeiros itens e ofereça detalhar por MDO ou por código.
- Exportar para planilha (ODI/ODD) ou PDF SAP é feito pelos botões da tela de Orçamento; você não exporta.

## Obras salvas
- `listar_obras` / `consultar_obra` para ver o que existe; `propor_obra` com modo adicionar, subtrair (retirada) ou carregar (substitui a tela). Obra MODELO pede parâmetros digitados pelo usuário.
