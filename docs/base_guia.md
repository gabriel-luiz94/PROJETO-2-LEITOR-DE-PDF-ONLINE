# Guia da base de orçamento — como escrever a sua

Este guia é para quem vai montar a **própria base** (a planilha que diz quais códigos cada ativo gera). Leia do simples ao avançado: cada recurso tem um exemplo curto e o resultado. Para ver tudo funcionando, importe o **pacote de exemplo** (`data/pacotes/exemplo/`, aba **Pacote** do modal Vinculação) num projeto novo e peça os mesmos exemplos abaixo.

## Como a base funciona

Cada **linha** da base diz: *quando o pedido contém o ATIVO, some esta quantidade deste CÓDIGO*. Colunas: `ATIVO`, `DESC ATIVO`, `COMPONENTE`, `PROJETO`, `MDO` (MATERIAL ou MAO-DE-OBRA), `CODIGO`, `DESC CODIGO`, `FATOR I` (instalar), `FATOR R` (retirar), `FILTRO`, `ORIGEM`, `LV`, `ARRED`.

- Quantidade do código = quantidade pedida × fator (`FATOR I` ao instalar, `FATOR R` ao retirar).
- O mesmo código pedido por várias linhas **soma**.
- Só a **primeira** linha de um ativo traz o nome dele em `ATIVO`; as outras ficam em branco, mas repetem o `COMPONENTE`.

## 1. Ativo simples

```
ATIVO | COMPONENTE | CODIGO | MDO           | FATOR I | FATOR R
N1    | N1         | 7101   | MATERIAL      | 1       | 1
      | N1         | 7102   | MATERIAL      | 1       | 1
      | N1         | 7103   | MAO-DE-OBRA   | 1       | 1
```
Pedido `2-N1` (instalar duas N1) → 7101 × 2, 7102 × 2, 7103 × 2. Pedido de retirada → os mesmos códigos com o `FATOR R`.

## 2. Parâmetro (o ativo muda conforme o pedido)

Declare o parâmetro na primeira linha: `N3=N3(cabo=C)` — `cabo` é o parâmetro e `C` é o **padrão**. As linhas comuns ficam no componente `N3`; cada valor ganha um **grupo**:

```
N3=N3(cabo=C)  | N3   | 7201 CRUZETA            (sempre)
               | N3   | 7202 INSTALAR N3        (sempre)
N3(cabo=C)     | N3.C | 7203 ISOLADOR COBRE × 3
N3(cabo=A)     | N3.A | 7204 ISOLADOR ALUMÍNIO × 3
```
Pedido `1-N3(A)` → 7201, 7202 e 7204 × 3. Pedido `1-N3` → usa o padrão `C` (7203). Cada grupo precisa do **seu** COMPONENTE (`N3.C`, `N3.A`).

## 3. Grupos com condição

A condição aceita `%` (qualquer trecho), `|` (ou) e `!=` (diferente): `N3(cabo=%35%)`, `N3(cabo=C|A)`, `N3(cabo!=C)`. **Todos** os grupos cuja condição casa com o pedido entram — por isso dois grupos que podem valer juntos não devem repetir o mesmo código (veja "Armadilhas").

## 4. Linha viva e linha morta (coluna LV)

`LV` = `LV` (a linha só vale em linha **viva**, operação com `*`), `LM` (só linha **morta**) ou `A` (os dois; padrão).
```
N4 | N4 | 7301 CRUZETA N4              | A
   | N4 | 7302 CARTUCHO N4             | LV
   | N4 | 7303 INSTALAR N4 (VIVA)      | LV
   | N4 | 7304 INSTALAR N4 (MORTA)     | LM
```
`1-N4` → 7301 e 7304. `*1-N4` (com `*`) → 7301, 7302 e 7303.

## 5. Retirada

O `FATOR R` vale ao **retirar** (operação `R`). `FATOR R` = 0 significa que retirar não devolve aquele código (típico de material que não volta). Ao pedir `M` (mudar), o programa usa só o que o seu projeto define para M na vinculação.

## 6. Cabos e arredondamento

Cabo é um ativo como outro: `C16` → código 7401 com `FATOR I` = `1,02` (2% de sobra por metro). A coluna `ARRED` arredonda **cada uso**: `CIMA` (para cima), `BAIXO`, `INTEIRO` ou `NORMAL` (não arredonda). Dois vãos de 30 m com `0,125` e `CIMA` dão `4 + 4 = 8`, não `ceil(7,5)`.

## 7. Composto (vários ativos com um nome só)

Uma linha só com o ATIVO (sem CÓDIGO nem COMPONENTE):
```
KIT_N3:=N3(C) + 3*ALCA(C) + 2*CON
```
Pedido `1-KIT_N3` → uma N3 de cobre, 3 alças de cobre e 2 conectores. Os itens precisam existir na base; ciclos (`A:=B`, `B:=A`) são recusados. Composto também pode ter parâmetros: `CF(cabo, cabo2=S04):=CHAVE_FACA(cabo,cabo2) + TERM(cabo)`.

## 8. Dados da obra (`OBRA.*`)

Parâmetros que o **programa** preenche a partir do cabeçalho da obra: `OBRA.MUNICIPIO`, `OBRA.BAIRRO`, `OBRA.UTD`, `OBRA.UTEP`, `OBRA.REGIAO`, `OBRA.CONTEXTO` (e os campos marcados "OBRA.*" no perfil de exportação do projeto). Declare por último: `FLBIM=FLBIM(poste=D, obra.regiao=CENTRO)` e use nos grupos: `FLBIM(obra.regiao=CENTRO|OESTE)`.

## 9. Armadilhas (o que dá erro ou resultado errado)

- **B-005 / B-006 — declaração.** Só uma declaração (`N3=N3(...)`) por ativo, e na **primeira** linha dele.
- **B-007 — parâmetro fora da declaração.** Todo parâmetro usado num grupo (`N3(cabo=C)`) precisa estar na declaração do ativo.
- **B-008 — padrão sem grupo.** Se o padrão é `cabo=C`, precisa existir um grupo que aceite `C`; senão pedir `N3` sem parâmetro falha. (Parâmetro usado só em fórmula não precisa de grupo.)
- **B-009 — componente repartido.** Dois grupos diferentes não podem usar o mesmo `COMPONENTE`: cada variante tem o seu (`N3.C`, `N3.A`).
- **B-010 — código em dois grupos que valem juntos.** O mesmo código em grupos que o pedido pode ativar ao mesmo tempo **soma duas vezes** (aviso). Ex.: `cabo=%35%` e `cabo=%A%` aceitam `A35P`: coloque o código num só.
- **B-013 — equipamento no grupo geral.** Item de catálogo de equipamento (transformador, chave, religador) nas linhas **gerais** de um ativo que tem parâmetro desse tipo: vale para **todo** pedido, mesmo quando o parâmetro escolhe outro equipamento (o caso do `CE_2_TR` que saía com dois transformadores). Coloque o equipamento no grupo do parâmetro (`CE_2_TR(trafo=%30%)`).
- **Valor não cadastrado.** Pedir `N3(X)` quando nenhum grupo aceita `X` recusa o ativo inteiro (aparece em "não encontrados").
- **Mesmo código duas vezes no mesmo grupo** fica um só.
- **Nomes:** maiúsculas/minúsculas e acentos contam no CÓDIGO da base; o ATIVO do pedido é convertido para maiúsculas.

## 10. Conferindo

1. **Validador:** ao importar, cada erro ou aviso traz um código (`B-0xx`) que aponta para a armadilha acima.
2. **Ficha da base:** abre a linha e mostra, em português, o que ela faz ("vale quando cabo contém A35P; soma 3 de 7204 na instalação").
3. **Casos de referência:** guarde 3–5 pedidos com o resultado certo (no Resultado: *Guardar como caso de referência*). Depois de mexer na base, use *Conferir base*: o programa mostra o que mudou, código a código.

## Códigos do validador

| Código | O que significa |
|---|---|
| B-001 | ativo com sintaxe inválida |
| B-002 | na base os parâmetros levam nome (`cabo=XLPE`) |
| B-003 | parâmetro `OBRA.*` que o projeto não conhece |
| B-004 | apelido repetido |
| B-005 | mais de uma declaração |
| B-006 | declaração fora da primeira linha do ativo |
| B-007 | parâmetro não declarado |
| B-008 | padrão sem grupo cadastrado |
| B-009 | COMPONENTE usado por grupos diferentes |
| B-010 | código que soma duas vezes (grupos que valem juntos) |
| B-011 | fórmula de FATOR com problema |
| B-012 | composto com problema (nome, item, parâmetro, ciclo) |
| B-013 | equipamento no grupo geral de ativo com parâmetro do mesmo tipo |
