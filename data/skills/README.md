# Skills do assistente

Cada arquivo `.md` aqui ensina o assistente de IA algo sobre o seu trabalho. Pastas:

- `_geral/` — vale para todos os projetos.
- `<projeto>/` — o nome do projeto em minúsculas e sem acento (`pernambuco/`). Skill do projeto vence a `_geral` de mesmo nome.

Formato: a 1ª linha pode ser `descricao: quando usar esta skill` (é o que o assistente lê antes de decidir abrir o arquivo); o resto é markdown livre, com listas curtas e regras objetivas. Nome do arquivo: letras minúsculas, números, `_` e `-`.

No programa, o botão «Skills» do assistente lista, edita e cria skills (administrador) e gera o catálogo `ativos` do projeto a partir da Base de Orçamento. O que for editado lá fica na pasta `skills/` ao lado do programa (no .exe) e sobrepõe a skill que veio no pacote.
