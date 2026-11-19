descricao: Como escrever as linhas das tabelas Outros e Cabos (quantidade, postes, parâmetros, notação de cabos, operação I/R). Leia antes de propor qualquer linha.
# Formato das linhas

## Operação
- `I` = instalar; `R` = retirar. `*I` / `*R` (com asterisco) = linha viva — SÓ quando o usuário disser "linha viva", "energizado" ou "sem desligamento".

## Tabela Outros
Cada linha é UM poste com o que vai nele, ou um grupo de ativos soltos (sem poste). Há dois formatos, **nunca misturados na mesma linha**:
- **Compacto** (o do dia a dia; itens separados por vírgula + espaço): `11300D, N1, 2 N3, 1 CFU, 3 PR220`. Cada item é `[quantidade ]ATIVO` — a quantidade vem ANTES, separada por espaço (`2 N3`), sem hífen. Sem quantidade = 1.
- **Antigo** (tokens `quantidade-ATIVO` separados por espaço, sem vírgulas): `DT11/300 1-N1 1-CFU 3-PR220`. Linha com vírgula + espaço usando `1-CFU` é ERRO ("quantidade do formato antigo numa linha em vírgulas").
- Escolha o formato que já predomina nas linhas da tela (`ler_tabela`); sem referência, use o compacto. O atalho "adicionar 2 CON" gera `2-CON` (um item só, formato antigo) e está certo.
- Vírgula SEM espaço depois é decimal (`0,5 N3`). `11600D,N1` (vírgula entre não-dígitos sem espaço) é erro.
- O poste vem PRIMEIRO e sem quantidade. Poste compacto = altura + esforço + letra (D = DT, R = RC, F = CV, FV): `11300D` = `DT11/300`. Prefira DT; CV só se o usuário disser (bandolado, sem acesso, local remoto).
- Nunca `1-DT11/300`, nem `POSTE11`, nem poste sem esforço. Nunca repita um ativo igual na mesma linha: some a quantidade.
- Sem unidades nem descrições livres ("kVA", "metros"): só códigos da base.

## Parâmetros de um ativo
- Alguns ativos mudam de itens conforme o contexto (tipo de cabo, tipo de poste). Escreve-se `N1(XLPE,DT)` (valores na ordem cadastrada) ou `N1(cabo=XLPE, poste=DT)`. `N1` sozinho assume o padrão da base.
- Quem decide os valores é o usuário/vínculo: NÃO invente parâmetros. Se o pedido não os cita, escreva o ativo sem parênteses.

## Tabela Cabos
- Formato clássico: `ATIVO FASES COMPRIMENTO m` — ex.: `CAA2 3 35 m` (3 fases, 35 metros). Sem fases: `CAA2 35 m`.
- Notação do label do projeto (também aceita): `[Nx]CÓDIGO[(neutro)][-TENSÃOkV] COMPRIMENTO` — ex.: `3xC16(16) 35 M` = 3 fases C16 + neutro C16, 35 m cada; `3xC16(16)+16 35 M` = 3 fases + neutro + um condutor extra C16 (5 ativos). `-13,8kV` indica média tensão e não altera a conta.
- A tabela Cabos NUNCA recebe itens de quantidade (`2 N3` / `2-N3`), e Outros NUNCA recebe linha terminada em comprimento `m`.
- Comprimento sempre em metros, vindo do usuário ou do desenho; nunca estime.

## Antes de propor
1. Confirme cada código com `buscar_ativo` (ou a skill `ativos` do projeto).
2. Rode `validar_linhas` nas linhas novas.
