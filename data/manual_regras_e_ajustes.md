# Manual: regras de validação e ajustes

Este manual explica, passo a passo, como cadastrar **regras de validação** e **ajustes** no Leitor de Projetos. Não é preciso saber programar: quase tudo é feito pelo editor visual. Os exemplos em formato de código (JSON) são para quem preferir editar direto, e cada um deles é testado automaticamente.

## 1. Visão geral

### Validação e ajuste

- **Validação** olha as tabelas **Cabos** e **Outros** e *avisa* o que parece errado. Ela nunca muda nada.
- **Ajuste** *corrige* as tabelas. Todo ajuste mostra uma **pré-visualização** (antes → depois) e só altera algo depois que você aceita. Um único **Ctrl+Z** desfaz a aplicação inteira.

### As três camadas da validação

1. **Contrato (fixa):** confere o formato das linhas (operação válida, `<qtd>-<ativo>` em Outros etc.). Está no código e não é editável.
2. **Regras técnicas (cadastráveis):** regras de domínio como "chave CFU precisa de SUPL na linha". É o que você cadastra neste manual.
3. **IA (opcional):** a IA revisa o que as regras não cobrem. Pode ser ligada ou desligada.

Cada camada pode ser ligada ou desligada na aba **Execução**.

## 2. Abrindo a gaveta

No **Resumo do projeto**, clique em **Regras de validação**. A gaveta lateral tem as abas:

| Aba | Para que serve |
|---|---|
| **Execução** | O que os botões *Validar* e *Ajustar* vão rodar (contrato, regras, IA). |
| **Regras** | Cadastro das regras técnicas. |
| **Grupos** | Listas de ativos com um nome (ex.: `@CHAVE_MT`), para reaproveitar nas regras e ajustes. |
| **Ajustes** | Cadastro dos ajustes (receitas) recorrentes. |
| **Prompts da IA** e **Testar** | Só administrador. |

No alto da gaveta, o campo **Editando** escolhe se você mexe no **Padrão do sistema** (vale para todos os projetos) ou em um **projeto** específico.

### Padrão × projeto

- O **padrão** vale para todos os projetos.
- Um **projeto** herda o padrão e pode **adicionar** regras próprias, **sobrescrever** uma regra do padrão ou **ocultar** (botão **−**) uma regra que não quer usar. Nada do padrão é perdido.
- Só **administradores** salvam alterações. Operadores veem tudo e usam normalmente.
- As alterações ficam em rascunho até você clicar em **Salvar**. Cada salvamento entra no **histórico** e pode ser revertido.

## 3. Regras de validação

### Anatomia de uma regra de linha

Cada linha da tabela **Outros** é um poste (ou um conjunto de ativos). Uma regra de linha diz: "**quando** a linha tem isto, **então** ela deve ter aquilo, **exceto** se...".

| Campo | Significado |
|---|---|
| `id` | Nome curto e único (ex.: `C2-CFU-SUPL`). |
| `mensagem` | O aviso mostrado ao usuário. |
| `severidade` | `erro`, `aviso` ou `info` (no editor, **gravidade**). |
| `ativa` | `true` liga, `false` desliga. |
| `operacoes` | Em quais operações a regra vale (padrão: `I` e `*I`). |
| `quando` | O gatilho (no editor, **SE**). Sem ele, a regra olha toda linha. |
| `entao` | O que deve valer quando o gatilho vale (**ENTÃO deve valer**). Sem ele, "é proibido". |
| `excecoes` | Se valer, a regra não se aplica (**EXCETO SE**). |

A regra gera alerta quando `quando` vale, `excecoes` não vale e `entao` **não** vale.

Exemplo: "linha com chave CFU precisa de pelo menos 1 SUPL".

```json regra
{
  "id": "EX-CFU-SUPL",
  "versao": 2,
  "escopo": "linha",
  "operacoes": ["I", "*I"],
  "ativa": true,
  "severidade": "aviso",
  "mensagem": "Chave CFU sem SUPL na linha.",
  "quando": {"tem": "CFU"},
  "entao": {"soma": "SUPL", "qtd": {">=": 1}}
}
```

### Cadastrando no editor visual

1. Aba **Regras** → **Adicionar regra** (em um projeto, **Adicionar regra só neste projeto**). Para mudar uma regra existente, clique em **Editar**.
2. Preencha o **id**, a **mensagem** e a **gravidade** (`erro`, `aviso` ou `info`).
3. Em **SE (gatilho)**, escolha o tipo da condição e preencha (veja a seção seguinte).
4. Em **ENTÃO deve valer** e, se quiser, **EXCETO SE**, faça o mesmo.
   Deixe **SE** vazio para valer em todas as linhas; deixe **ENTÃO** vazio para a linha que cair no gatilho ser considerada proibida.
5. Confira a **frase em português** que o editor mostra: ela descreve a regra por extenso. Se a frase não é o que você quer, a regra não está certa.
6. Marque **ativa** e clique em **Salvar**.

### Condições

| Tipo no editor | JSON | O que significa |
|---|---|---|
| **tem o ativo** | `{"tem": "CFU"}` | A linha tem o ativo (qualquer quantidade). |
| **tem o ativo** com quantidade | `{"tem": "CFU", "qtd": {"=": 1}}` | Tem o ativo **na quantidade** indicada. |
| **soma dos ativos** | `{"soma": "SUPL", "qtd": {">=": 1}}` | A soma das quantidades dos ativos que casam. |
| **texto casa com** | `{"texto": "^DT"}` | O texto da linha casa com a expressão regular. |
| **TODAS (E)** | `{"todos": [A, B]}` | Vale quando **todas** as condições valem. |
| **ALGUMA (OU)** | `{"algum": [A, B]}` | Vale quando **pelo menos uma** vale. |
| **NENHUMA** | `{"nenhum": [A, B]}` | Vale quando **nenhuma** vale. |
| **NÃO** | `{"nao": A}` | Inverte a condição. |

Comparadores de quantidade: `=`, `!=`, `>=`, `<=`, `>`, `<` e `entre` (ex.: `{"entre": [1, 3]}`).

### Como usar o "E" (ex.: CFU **e** U4)

Escolha o tipo **TODAS (E)** e, dentro dele, adicione duas condições **tem o ativo**: uma com `CFU` e outra com `U4`. Exemplo: "linha com CFU e U4 precisa ter SUPL".

```json regra
{
  "id": "EX-CFU-U4-SUPL",
  "versao": 2,
  "escopo": "linha",
  "operacoes": ["I", "*I"],
  "ativa": true,
  "severidade": "aviso",
  "mensagem": "Linha com CFU e U4 sem SUPL.",
  "quando": {"todos": [{"tem": "CFU"}, {"tem": "U4"}]},
  "entao": {"soma": "SUPL", "qtd": {">=": 1}}
}
```

Para o **OU** ("CFU **ou** U4"), troque `todos` por `algum`:

```json condicao
{"algum": [{"tem": "CFU"}, {"tem": "U4"}]}
```

Para "tem CFU **e não** tem SUPL":

```json condicao
{"todos": [{"tem": "CFU"}, {"nao": {"tem": "SUPL"}}]}
```

### Escolhendo os ativos (seletores)

Onde a regra pede um ativo, você pode usar:

| Forma | Exemplo | Casa com |
|---|---|---|
| Código exato | `CFU` | Só o ativo `CFU`. |
| Curinga | `TR[0-9]*`, `EF*` | `*` = qualquer texto, `?` = um caractere, `[..]` = um dos caracteres. |
| Grupo | `@CHAVE_MT` | Qualquer ativo do grupo (veja a aba **Grupos**). |
| Lista | `["CFU", "CFA"]` | Qualquer um da lista. |
| Expressão regular | `{"regex": "^SI\\d+$"}` | Quem casa com a expressão. |

Maiúsculas e minúsculas não importam.

Exemplo: "trafo sem PR15, exceto se houver RPR", usando um grupo e uma exceção:

```json regra
{
  "id": "EX-TR-PR15",
  "versao": 2,
  "escopo": "linha",
  "ativa": true,
  "severidade": "aviso",
  "mensagem": "Trafo sem 1-PR15 (nem RPR).",
  "quando": {"tem": "@TRAFO"},
  "entao": {"soma": "PR15", "qtd": {">=": 1}},
  "excecoes": {"tem": "RPR"}
}
```

### Grupos

Na aba **Grupos** você dá um nome a uma lista de ativos (nome em maiúsculas, ex.: `CHAVE_MT`) e depois usa `@CHAVE_MT` em qualquer regra ou ajuste. Se o grupo mudar, todas as regras que o usam acompanham. Um grupo pode usar outro (`@OUTRO`), mas não pode apontar para si mesmo, direta ou indiretamente.

### Regras de planilha (totais)

Além das regras por linha, há regras sobre os **totais** das tabelas, por exemplo "a quantidade de P50 deve ser igual ao total de postes":

```json regra
{
  "id": "EX-P50-POSTES",
  "versao": 2,
  "escopo": "planilha",
  "ativa": true,
  "severidade": "aviso",
  "mensagem": "A quantidade de P50 não bate com os postes.",
  "deve_ser": {
    "esq": {"soma_qtd": "P50"},
    "cmp": ">=",
    "dir": {"soma_qtd": "@ESTRUTURA_MT"},
    "rotulos": {"esq": "P50", "dir": "estruturas de MT"}
  }
}
```

### Testando uma regra

Rode **Validar** no Resumo: os achados aparecem com a mensagem da regra. Se algo vier errado, use o botão de explicação do achado para ver o que a regra observou na linha.

## 4. Ajustes (correções)

Um **ajuste** (ou *receita*) é uma lista de **ações** que corrigem as tabelas. Ele pode ser cadastrado uma vez e usado sempre.

### Cadastrando um ajuste

1. Aba **Ajustes** → **Novo ajuste**.
2. Dê um **nome** e uma **descrição**.
3. Adicione uma ou mais **ações** (veja abaixo) e escolha a **tabela** (`cabos`, `outros` ou `ambos`).
4. Marque **ativo** (só ajustes ativos entram no botão **Ajustar**).
5. Clique em **Salvar**.

Cada ação aceita, opcionalmente, `operacoes` (só linhas com essas operações) e `quando` (uma condição como as das regras; em **Cabos** só vale `texto`).

### As ações

**substituir:** troca um texto por outro.

```json acoes
[{"acao": "substituir", "tabela": "outros", "de": "SUP-L", "para": "SUPL"}]
```

Com `"modo": "item"`, troca o **ativo** de cada item `<qtd>-<ativo>` (mantendo a quantidade). Com `"palavra_inteira": true`, só troca a palavra completa.

**normalizar:** arruma o formato. As regras são `espacos` (espaços repetidos), `maiusculas` e `poste` (tira o `1-` antes de `DT11/300` no início da linha).

```json acoes
[{"acao": "normalizar", "tabela": "outros", "regras": ["espacos", "poste"]}]
```

**ordenar:** reordena a tabela por colunas (`operacao`, `ativo`, `entidade`), com `ordem` `asc` ou `desc` e, se quiser, uma ordem personalizada em `valores`.

```json acoes
[{"acao": "ordenar", "tabela": "outros", "por": [
  {"coluna": "operacao", "valores": ["I", "*I", "R", "*R", "M", "*M"]},
  {"coluna": "ativo"}
]}]
```

**excluir_linhas:** apaga linhas vazias, duplicadas, que casam com uma condição ou com um texto.

```json acoes
[{"acao": "excluir_linhas", "tabela": "ambos", "onde": {"vazias": true}}]
```

**adicionar_linha:** insere uma linha em uma tabela (`cabos` ou `outros`, nunca `ambos`), no `fim`, no `inicio` ou antes/depois de uma condição.

```json acoes
[{"acao": "adicionar_linha", "tabela": "outros",
  "valores": {"operacao": "I", "ativo": "1-RA2", "entidade": "0"},
  "posicao": "fim", "apenas_se_nao_existir": true}]
```

**adicionar_ativo** (só Outros): acrescenta um ativo nas linhas que casam. `se_ja_existe` define o que fazer se ele já estiver lá: `ignorar`, `somar` ou `substituir`.

Exemplo: "em toda linha com CFU **e** sem SUPL, adicionar 1-SUPL".

```json acoes
[{"acao": "adicionar_ativo", "tabela": "outros", "ativo": "SUPL", "qtd": 1,
  "operacoes": ["I", "*I"],
  "quando": {"todos": [{"tem": "CFU"}, {"nao": {"tem": "SUPL"}}]}}]
```

Use `"qtd"` **negativo** para adicionar como "linha viva"/retirada — o ativo entra com o prefixo `*` (ex.: `qtd: -1` gera `*1-PR`, nunca `-1-PR`, que o cálculo não leria como negativo).

```json acoes
[{"acao": "adicionar_ativo", "tabela": "outros", "ativo": "PR", "qtd": -1,
  "quando": {"tem": "TR110"}}]
```

**remover_ativo** (só Outros): tira um ativo (aceita os mesmos seletores das regras) das linhas que casam.

```json acoes
[{"acao": "remover_ativo", "tabela": "outros", "ativo": "SUP-L"}]
```

**mesclar_duplicadas** (só Outros): soma as quantidades do mesmo ativo repetido na linha (`1-CFU 1-CFU` vira `2-CFU`).

```json acoes
[{"acao": "mesclar_duplicadas", "tabela": "outros"}]
```

**criar_linha_derivada:** pra cada linha que casar com `de` (um regex, com grupo(s) entre parênteses), cria uma linha
**nova** logo depois (ou antes) dela — nunca edita a linha original. `ativo_novo` é um texto que pode reusar o que foi
capturado por `de` com `\1`, `\2`... (o que o grupo 1, grupo 2... capturou NAQUELA linha). Diferente de
`adicionar_linha` (texto sempre igual, não depende de regex), esta ação roda uma vez por linha que casar, com
conteúdo derivado dela.

Exemplo real: em Cabos, `"(2)"` depois da bitola significa "tem neutro, bitola 2" — se a bitola da linha já é 2, o
neutro só precisa de uma letra na fase (outro ajuste, de `substituir`, cuida disso — ver abaixo); se é diferente,
o neutro precisa de uma linha própria, com o mesmo comprimento da linha original:

```json acoes
[{"acao": "criar_linha_derivada", "tabela": "cabos",
  "quando": {"nao": {"texto": "(?<![A-Za-z0-9/])2\\(2\\)"}},
  "de": {"regex": "^[A-Za-z]+\\s*[\\d/]+\\(2\\)\\s+[A-Za-z]+\\s+([\\d.,]+)\\s*m\\s*$"},
  "ativo_novo": "CA 2 N \\1 m"}]
```

Com a linha `"CA 1/0(2) ABC 35m"` (bitola `1/0`, diferente de `2`), essa ação cria a linha `"CA 2 N 35 m"` logo
abaixo — o `35` veio do grupo capturado NAQUELA linha, não de um valor fixo na regra.

Campos opcionais: `operacao_nova` (operação da linha nova; sem ele, usa a operação da linha que casou),
`entidade_nova` (padrão `"0"`), `posicao` (`"depois"`, padrão, ou `"antes"`), `apenas_se_nao_existir` (padrão
`true` — não duplica se a linha derivada já existir). A linha recém-criada nunca é reavaliada contra `de` na
mesma execução do ajuste (sem risco de crescer em cadeia).

**Atenção à ordem das ações**: se outra ação do mesmo ajuste remove o trecho que `de` procura (no exemplo, o
`"(2)"`), ela precisa rodar **depois** de `criar_linha_derivada` — senão, quando a ordem chegar nela, o regex já
não bate mais e a linha nova deixa de ser criada, sem nenhum erro visível.

### Mudando a operação de uma linha

Por segurança, uma ação **não muda a operação** (`I`, `R`, `M`...) da linha. Para mudar, a ação precisa dizer explicitamente `"operacao_nova": "*I"`; e só as linhas que a ação de fato alterou mudam.

### Ajustes ligados a regras

Um ajuste pode apontar as regras que ele corrige (campo `regras`). Assim, no painel de validação, o achado dessa regra ganha o botão **Ajustar**.

### Usando os ajustes: o botão Ajustar

1. No Resumo, clique em **Ajustar**.
2. O programa roda **todos os ajustes ativos** nas tabelas atuais e abre um painel com **um cartão por ajuste que muda alguma coisa**.
3. Em cada cartão, abra **Ver e escolher as mudanças**, desmarque o que não quiser e clique em **Executar correção**.
4. Para aplicar tudo de uma vez, use **Executar todas as correções** (no topo): os ajustes rodam em sequência, na ordem do cadastro, e um **Ctrl+Z** desfaz o conjunto.
5. Quando um ajuste **exclui** linhas ou itens, ele aparece destacado e o programa pede uma confirmação extra.
6. Ajustes que não mudariam nada não aparecem como cartão.

Proteções: se uma correção deixar a linha com um erro de contrato novo, ela é descartada e listada no cartão; se a linha mudou desde a pré-visualização, aquela mudança é ignorada.

### Exportar e importar em arquivo (JSON)

Os **ajustes** (aba Ajustes) e as **regras do leitor** (botão **Regras do Leitor** da tela do Leitor, tabelas de Processamento e de Classificação) podem ser levados para um arquivo `.json` e trazidos de volta:

- **Exportar JSON** baixa a lista como está na tela, inclusive o que ainda é rascunho e as regras desligadas. O arquivo traz `formato`, `versao`, o `projeto` de origem e a lista (`ajustes` ou `regras`). Não leva usuário, e-mail nem dados de obra.
- **Importar JSON** lê o arquivo, confere tudo (formato, tamanho, regras e regex) e, antes de aplicar, mostra o resumo: quantas são **novas**, quantas **já existem** e quantas seriam **removidas**. Você escolhe **Acrescentar** (mantém o que existe e junta o novo; as regras do leitor novas entram no fim, com a ordem renumerada; ajustes com o mesmo código são atualizados) ou **Substituir** (a lista fica igual à do arquivo).
- **Nada é gravado na importação**: o resultado vai para o **rascunho**. Só vale ao clicar em **Salvar** / **Salvar alterações**, que valida de novo no servidor e guarda a versão anterior no **Histórico** (dá para reverter).
- Arquivo de **outro projeto** é aceito (serve para copiar regras entre projetos), com um aviso na janela. Arquivo da **outra tabela** (ex.: Classificação importado em Processamento) ou de outro cadastro é recusado.
- Limites: até 1 MB, 1.000 itens e 500 caracteres por regex. Também é aceita uma lista pura (como os arquivos de semente).
- Cada **regra do leitor** ainda tem o **Modo avançado (JSON)** no editor: mostra a regra em JSON, que você edita e aplica; JSON inválido fica em vermelho e não deixa continuar (os ajustes já tinham esse modo).

### Ler o projeto direto no Resumo

O botão **Ler projeto** (no alto do Resumo) faz tudo de uma vez: lê o PDF/DXF/DWG, extrai os textos e aplica as regras do leitor do projeto selecionado, preenchendo Cabos e Outros **sem sair da aba**. Clique nele ou arraste o arquivo sobre ele. Se as tabelas já tiverem linhas, escolha **Substituir** (padrão, Enter), **Adicionar** ou **Cancelar**. Ctrl+Z desfaz. Se o projeto não tiver regras do leitor, nada é alterado.

### Tabelas como planilha (Cabos e Outros)

As tabelas **Cabos** e **Demais Ativos** funcionam como uma planilha: clique e arraste (ou Shift+clique) para selecionar várias células, clique no **número da linha** para selecionar a linha, **Ctrl+A** para tudo. **Ctrl+C** copia em formato de planilha (cola direto no Excel); **Ctrl+V** cola várias linhas do Excel a partir da célula atual — sobrescreve as linhas existentes e cria novas se faltar. Uma coluna cola no Ativo; duas colunas, Operação + Ativo (em Cabos, uma terceira com `-` marca a linha como negativa). Operação inválida é ignorada (mantém a que estava) e aparece um aviso. **Delete** esvazia o Ativo do intervalo; **Excluir** remove as linhas selecionadas. **Ctrl+Z** desfaz a colagem inteira de uma vez.

Em **Demais Ativos** a **Entidade** é a primeira coluna: ela é recalculada sozinha quando o ativo muda; se a regra errar, clique nela e escolha outra (a linha fica travada 🔒) ou use o botão direito → **Definir entidade…** ("Automática" destrava). Em Cabos não há Entidade.

Para **classificar**, clique no **título da coluna**: o 1º clique ordena do menor para o maior (▲), o seguinte inverte (▼); **Shift** + clique adiciona um segundo critério. Com o **botão direito** no título você escolhe **Agrupar por esta coluna** (separadores com a contagem de cada grupo) ou **Limpar agrupamento**. Classificar reordena as linhas de verdade (Ctrl+Z desfaz) e mantém os vínculos cabo↔estrutura; em Cabos a **Qtd Ativos** não é recalculada. Para filtrar por valor use a caixa “Todos” abaixo do título. O ícone de **olho** ao lado do contador de Demais Ativos oculta ou mostra a coluna Entidade.

### Geradores de Cabos e de Postes/Estruturas

Os dois geradores têm painéis de atalhos. **Operação** (I, R, M…) cria uma linha nova; **ativo, fase e comprimento** alteram a linha em destaque (amarela): a linha selecionada ou, sem seleção, a última. Em **＋ Novo** você cria atalhos seus e escolhe onde guardar: **só neste computador** (⌂) ou **no projeto** (☁, todos do projeto veem; só o administrador exclui os do projeto). **✎ Editar atalhos** permite editar, excluir ou ocultar (os de fábrica só podem ser ocultados). **★** guarda a linha (operação, ativo, fase e comprimento) como **favorito**; clicar no favorito insere a linha pronta. As linhas são como as de uma planilha: selecione, **Ctrl+C**, **Ctrl+V** (inclusive uma coluna só com “CAA 2 ABC 80 m”), **Delete** e **Ctrl+Z** dentro do modal.

### Resultado do Orçamento: modelo REC ou SAP

Na tela de resultado há um alternador **REC | SAP**. O **REC** é a tabela única de sempre. Na **SAP** os mesmos itens aparecem em duas tabelas: **ODI** (tudo que está instalando) e **ODD** (tudo que está removendo), com TIPO MDO abreviado (**M** = material, **N** = mão de obra) e uma coluna **CENTRO** (**1600** para material, **1000** para mão de obra — acompanha o TIPO MDO). A ordem das colunas é TIPO MDO, CÓDIGO, QUANTIDADE, CENTRO e DESCRIÇÃO (o Copiar e o Excel seguem a mesma ordem). As tabelas têm os mesmos campos editáveis, filtros, busca e ordenação; **⇄** move uma linha entre ODI e ODD; **Copiar** (acima de cada tabela) copia só aquela tabela. **Exportar Excel** (só na SAP) gera um arquivo com as planilhas ODI e ODD. Os botões **Copiar** de ODI e ODD levam TIPO MDO, CÓDIGO, QUANTIDADE e CENTRO (sem a descrição). Também dá para **selecionar células** (arrastando ou com Shift+clique) e copiá-las de uma vez com **Ctrl+C** ou **Copiar seleção**; o que já foi copiado fica **marcado em verde-azulado com ✓** (a marca some se o valor mudar; **Limpar marcações** recomeça). A SAP é apenas outra forma de ver: salvar no banco e exportar o `.rec` continuam exatamente como no REC.

### Resultado do Orçamento: comparar com um REC

Além de **somar** e **subtrair**, dá para **comparar** o REC atual com outro: **⇄ Comparar** (arquivo `.rec`) ou **⇄ Comparar** ao lado da obra, em **RECs do Banco**. Aparece uma coluna **REFERÊNCIA** antes de QUANTIDADE (nas visões REC e SAP) e as linhas ganham cor: **verde** = igual, **amarelo** = quantidade diferente, **vermelho** = existe só no atual (a referência vale 0), **azul** = existe só na referência (o atual vale 0; a linha é só visual, com um **＋** para adicioná-la ao atual). A barra do modo liga/desliga a comparação (a referência fica guardada até **✕ Remover**) e tem um filtro para ver só um tipo de diferença. Salvar, exportar `.rec`, Copiar e Exportar Excel não levam a coluna nem as linhas azuis.

## 5. Usando a IA

- **Validar com IA:** ligue a camada de IA na aba **Execução**. A IA aponta o que as regras não cobrem. Só administradores editam os prompts da IA.
- **Ajustar com IA:** descreva o que quer corrigir; a IA propõe ações de ajuste. Elas **sempre** passam pela pré-visualização, inclusive quando excluem algo.

## 6. Receitas prontas

Corrigir grafia de um ativo e padronizar os postes:

```json acoes
[
  {"acao": "substituir", "tabela": "outros", "de": "SUP-L", "para": "SUPL"},
  {"acao": "normalizar", "tabela": "outros", "regras": ["espacos", "poste"]}
]
```

Limpar e ordenar:

```json acoes
[
  {"acao": "excluir_linhas", "tabela": "ambos", "onde": {"vazias": true}},
  {"acao": "mesclar_duplicadas", "tabela": "outros"},
  {"acao": "ordenar", "tabela": "outros", "por": [{"coluna": "ativo"}]}
]
```

## 7. Perguntas frequentes

**Como faço "A e B"?** Use a condição **TODAS (E)**, com uma condição para cada. Exemplo na seção 3.

**Como faço "A ou B"?** Use **ALGUMA (OU)**.

**Como aceito variações como `SUP-L` e `SUPL`?** Use uma expressão regular: `{"regex": "^SUP-?L$"}`.

**Quero que uma regra do padrão não valha no meu projeto.** Escolha o projeto no campo **Editando** e clique no botão **−** da regra (*Ocultar*). Ela continua listada, mas desligada só naquele projeto; o botão **+** a religa. O padrão não é alterado.

**Errei e salvei.** Use o **histórico** da aba e reverta para a versão anterior.

**O botão Ajustar não mostra nada.** Ou não há ajuste **ativo** (ligue na aba Ajustes), ou nenhum deles mudaria as tabelas atuais.

**Por que o Salvar não aparece?** Só administradores salvam. Operadores usam as regras e ajustes cadastrados.

## 8. Erros comuns

- **Regra que alerta em toda linha:** faltou o `quando`. Defina o gatilho.
- **Grupo não encontrado:** o nome do grupo está errado ou foi ocultado no projeto.
- **Ciclo de grupos:** o grupo A usa o B e o B usa o A.
- **Expressão regular inválida ou grande demais:** simplifique.
- **Ajuste que "não faz nada":** confira a tabela (`cabos`/`outros`), as `operacoes` e o `quando`.
