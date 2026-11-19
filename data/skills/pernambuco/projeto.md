descricao: Convenções do projeto Pernambuco (código 081): nomes de poste, notação, onde achar a base e regionais.
# Projeto Pernambuco (081)

- **Postes** (tipos): DT duplo T (sufixo D), RC redondo (R), CV concreto (F), FV fibra de vidro. Altura de 7 a 16 m; esforço com 3 ou 4 dígitos. Formas válidas do mesmo poste: `DT11/300`, `D11/300`, `11300D`. O nome oficial é o primeiro (`DT11/300`).
- **Ativos**: só valem os códigos da Base de Orçamento do projeto. A skill `ativos` (catálogo gerado da base) lista as famílias e códigos; gere-a de novo no botão «Gerar catálogo» quando a base mudar. Na dúvida, `buscar_ativo`.
- **Parâmetros**: ativos como N1 podem ter variantes por tipo de cabo/poste (`N1(XLPE,DT)`); a ordem e os padrões estão na declaração da base. Não invente valores.
- **Cabos**: aceitam a notação do label (`3xC16(16)+16 35 M`) além do formato clássico; a Totalizadora expande cada fase.
- **Regional da obra** (município/UTD → regional) vem da tabela regional do projeto e preenche os parâmetros OBRA.* usados em algumas regras do orçamento.

> Edite este arquivo (botão Skills do assistente) para registrar o que for específico do seu dia a dia: padrões que você sempre usa, exceções, nomes de obras-modelo.
