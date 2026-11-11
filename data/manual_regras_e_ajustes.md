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

**somar_linhas** (só Outros): junta, **entre linhas diferentes**, as quantidades dos ativos que você escolher. A soma fica na primeira linha em que o ativo aparece e as outras ocorrências saem das suas linhas (a linha que ficar vazia é excluída). Só junta itens da mesma operação (`I` com `I`, `*I` com `*I`) e, por padrão, da mesma entidade (`"por_entidade": false` ignora a entidade). Linhas em notação compacta (com vírgulas, como `9400D, 1-CFU`) não são editadas. No editor visual: **Somar ativos de várias linhas em uma**.

```json acoes
[{"acao": "somar_linhas", "tabela": "outros", "ativo": ["CFU", "PR"], "por_entidade": false}]
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
- **Só administradores** veem Exportar/Importar JSON e editam as regras do leitor. Quem não é admin vê apenas a lista (ou o aviso de lista vazia). Os projetos 027 e 229 já nascem com as regras de fábrica.
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

## Modais montáveis (barra de modais do Resumo)

O botão **＋** (só administrador) cria modais próprios do **projeto selecionado**; todos os usuários do projeto usam os botões que aparecem na barra, depois de CABOS, CONEXÕES, RAMAIS e VINCULAR CABOS.

- **Criar/editar:** ＋ → "+ Novo em branco" ou um dos exemplos (**Ramais colados**, **Títulos de contrato**) → monte na aba **Visual** (blocos com formulário; ▲▼ reordena, ✕ remove, "Adicionar bloco" no fim) ou na aba **JSON** (as duas ficam sincronizadas) → **Salvar**. O servidor confere tudo (blocos, regex, expressões) e mostra os erros.
- **Blocos:** `texto`, `campo` (texto, número, lista, opções), `grade` (cole várias linhas do Excel na primeira célula), `tabela_interna` (planilha de consulta embutida), `interpretar` (regras: se o texto casa a regex → gera ativo + quantidade; a quantidade aceita `+ - * /` e capturas como `{1}*20`; a primeira regra que casa vale), `destino` (Outros ou Cabos), `acao_adicionar` (botão "Adicionar ativos") e `saida_texto` (modelo com `{{campo}}` e `{{buscar(tabela, colunaChave, valor, colunaRetorno)}}`, com botão **Copiar**).
- **Condicional:** qualquer bloco aceita `visivel_se` e cada regra de interpretação aceita `se`. Uma condição é `{"campo": "regiao", "op": "igual", "valor": "SUL"}` (ops: igual, diferente, vazio, preenchido, em, contem, maior, menor, maior_igual, menor_igual) ou um grupo `{"todos": [...]}` (E), `{"algum": [...]}` (OU), `{"nao": {...}}`. No editor visual use "+ condição", "+ grupo" e "NÃO". Na quantidade e no ativo você pode usar o valor de um campo com `[id]` (ex.: `{1}*[fator]`, `P{1}-[regiao]`).
- **Colagem da grade** (`colagem`): `colunas` (padrão: colunas do Excel), `linha` (linha inteira na 1ª coluna) ou `dividir` (divide a linha pelo separador, ex.: `;`); `pular_cabecalho` ignora a 1ª linha colada. `carregar_de: "outros"` ou `"cabos"` cria o botão **Carregar da tabela** (colunas de id `operacao`, `ativo`, `entidade`), útil para modais de processamento.
- **Interpretação:** `quando.global: true` gera um item por ocorrência na linha; `normalizar` (maiusculo/minusculo) padroniza o ativo; as linhas que nenhuma regra reconheceu aparecem em laranja abaixo do resultado.
- **Cálculo e expressões:** o bloco `calculo` (`{id, rotulo, expressao}`) mostra um resultado e o torna disponível como `[id]`. As expressões (também na quantidade das regras e em `{{= … }}` nos modelos) aceitam números, `"texto"`, `[campo]`, `{1}` (captura da regex), `+ - * /`, comparações (`= <> < > <= >=`) e as funções `SE(cond; a; b)`, `E`, `OU`, `NAO`, `MAX`, `MIN`, `ABS`, `ARRED(x; casas)`, `SOMA`, `CONCAT`, `MAIUSCULO`, `MINUSCULO` (argumentos separados por `;`). Exemplo de quantidade: `SE([zona]="A"; {1}*20; {1}*10)`. Um cálculo só pode citar campos e cálculos definidos **antes** dele. Expressão inválida = item ignorado; o servidor recusa ao salvar.
- **Adicionar ativos:** em **Outros** entra uma linha `qtd-ativo` por item; em **Cabos** o ativo se repete `qtd` vezes. Ctrl+Z desfaz.
- **Exportar / Importar / Histórico (admin):** no ＋ há **Exportar JSON** (todos os modais do projeto), **Importar JSON** (mostra quantos são novos, quais já existem e quais seriam removidos; só grava ao escolher **Acrescentar/atualizar** ou **Substituir tudo**) e **Histórico** (as últimas 20 versões, com data e autor; **Restaurar** volta àquela versão e guarda a atual). Use o arquivo para levar modais de um projeto para outro.
- Os modais fixos (Cabos, Postes, Conexões, Ramais, Vincular) não mudam. Nada precisa ser rodado no Supabase.

## Linha viva x linha morta (coluna LV da base de orçamento)

Na tela da base de orçamento há a coluna **LV** com três valores: **LV** (a linha só vale para serviços em **linha viva**), **LM** (só para **linha morta**) e **A** (vale para os dois; é o padrão e o valor de todas as linhas antigas).

- **Linha viva** é a operação com `*` antes: `*I` (instalar em linha viva), `*R`, `*M`. Sem `*` é linha morta.
- No cálculo, uma linha **LV** só entra para operações com `*`, uma **LM** só para operações sem `*`, e **A** entra nas duas. Isso vale na busca do ativo e também ao trazer todos os códigos do componente.
- Se um ativo só existe com o outro tipo (ex.: `*I` num ativo que só tem linhas LM), ele aparece em **não encontrados** com a dica "só existe para linha morta/viva".
- No CSV, a coluna se chama `LV`; arquivo sem ela = tudo `A`; valor diferente de LV/LM/A é recusado (a mensagem diz a linha) e nada é alterado.
- Depois de atualizar, o administrador roda o `schema_supabase.sql` no Supabase (cria a coluna) — nada muda no orçamento até alguém marcar LV ou LM.

## Arredondamento por linha (coluna ARRED da base de orçamento, TASK-107)
Na tela da base há a coluna **ARRED** (depois de LV) com quatro valores: **NORMAL** (padrão: não arredonda cada uso; o total do código sai com 2 casas, como sempre), **INTEIRO** (arredonda a 0 casas, meio para cima), **CIMA** (próximo inteiro acima) e **BAIXO** (inteiro abaixo).

- O arredondamento vale para **cada uso da linha**, antes da soma por código: `ceil(quantidade × fator)` por pedido/vão. Dois vãos de 30 m com FATOR I `0,125` e CIMA geram `4 + 4 = 8`, não `ceil(7,5)`.
- Instalação usa o **FATOR I** e remoção o **FATOR R**; os dois passam pelo mesmo ARRED da linha.
- Exemplo (espaçador, 1 a cada 8 m): linha com COMPONENTE `CAZ7,94`, CÓDIGO `3426163`, FATOR I `0,125`, ARRED `CIMA` → vão de 30 m = 4, de 64 m = 8, de 5 m = 1.
- Código direto (pedir o próprio código) continua com fator 1 e sem arredondamento.
- No CSV a coluna se chama `ARRED`; arquivo sem ela = tudo NORMAL; `PARA CIMA`/`PARA BAIXO` também valem; valor diferente é recusado (a mensagem diz a linha) e nada é alterado.
- Depois de atualizar, o administrador roda o `schema_supabase.sql` no Supabase (cria a coluna e marca tudo como NORMAL).

## Ativos com variantes (parâmetros) na base de orçamento

Quando um ativo muda de itens conforme uma categoria (tipo de cabo, tipo de poste…), cadastre **uma variante por grupo de linhas**, na própria tabela base:

- **Declaração** (1ª linha do ativo): `N1=N1(cabo=XLPE, poste=DT)` — define os parâmetros, a **ordem** e os **valores padrão** (`N1` sozinho vale `N1(XLPE,DT)`). Sem padrão: `N1=N1(cabo, poste)`. Sem declaração nenhuma, a ordem é a da primeira aparição e não há padrão.
- **Grupos:** `N1(cabo=XLPE)` vale para cabo XLPE e qualquer poste; `N1(poste=DT)` para qualquer cabo; `N1(cabo=COBRE, poste=CC)` só para a combinação; `N1` sem parâmetros é o grupo **comum** (vale sempre). Cada grupo tem o seu **COMPONENTE** (ex.: `N1`, `N1.XLPE`, `N1.DT`) — o mesmo componente em grupos diferentes é recusado.
- **Em Outros/Totalizadora** escreva `3-N1(XLPE,DT)` (só os valores, na ordem cadastrada) ou `3-N1(poste=CC, cabo=COBRE)` (com nomes). `3-N1` completa pelo padrão. Valores usam letras, números, `_` e `.` (sem hífen) e não diferenciam maiúsculas.
- **Sem padrão** e pedido incompleto (`N1`, `N1(ALUM)`): calcula o que não depende do que falta e mostra um **aviso** com os valores aceitos. Valor inexistente (`N1(XLIPE)`) aparece em "não encontrados" com a lista de variantes cadastradas.
- Ao importar a base o sistema **recusa** cadastro inconsistente (declaração repetida ou fora da 1ª linha, parâmetro fora da declaração, componente repetido entre grupos, caractere inválido) e **avisa** quando o mesmo código está em grupos que valem juntos (a quantidade soma duas vezes).
- **Condições no grupo (só na base):** `N1(poste=D%)` = poste que **começa com** D (`%T` termina em T, `%CC%` contém CC); `N1(cabo=ALUM|COBRE)` = um **ou** outro; `N1(cabo!=XLPE)` = **qualquer exceto** XLPE (`cabo!=ALUM|COBRE` = nenhum dos dois); vários parâmetros na mesma célula valem **juntos** (`N1(poste=D%, cabo!=XLPE)`; para OU entre parâmetros diferentes, use dois grupos). O pedido continua com valores concretos (`3-N1(ALUM,D12)`): é a base que decide se `D12` casa `D%`. A declaração (`N1=N1(cabo=XLPE, poste=DT)`) só aceita valores concretos como padrão. Um valor do pedido é aceito se for igual a um cadastrado ou casar um padrão; um parâmetro que só tem `!=` aceita qualquer valor.
- **COMPONENTE automático:** deixe o `COMPONENTE` em branco nas linhas dos ativos com parâmetros. Ao importar ou salvar, a linha que começa um grupo (a que tem o ATIVO) recebe um nome como `N1`, `N1.CABO=XLPE`, `N1.CABO=ALUM|COBRE` ou `N1.POSTE=D%,CABO!=XLPE`, e as linhas seguintes sem ATIVO herdam esse nome até o próximo ATIVO. Um componente que você digitou nunca é trocado; importar o mesmo arquivo de novo gera os mesmos nomes; ativos sem parâmetros não são tocados. A tela da base mostra quantos foram preenchidos.
- Combina com a coluna **LV** (linha viva/morta) e com a soma por código (o mesmo código em grupos exclusivos soma normalmente).
- Nada a rodar no Supabase.

## Notação compacta na tabela Outros (11600D, 2 N3, T3)

Além do formato antigo (`1-DT11/600 2-N3 1-T3`), a tabela Outros aceita a lista com **vírgula + espaço**:

- `11600D, 2 N3, T3` — cada item é `[quantidade] ativo` (sem quantidade vale 1; `0,5 N3` usa vírgula decimal; `*2 N3` é negativo).
- Postes: `D` = DT, `R` = RC, `F` = CV, `FV` = FV (`11300D` = `DT11/300`, `7150FV` = `FV7/150`).
- **Apelidos**: na coluna ATIVO da base escreva `DT11/300, 11300D`; `11300D` passa a valer como esse ativo em qualquer formato.
- Uma linha **não mistura** os dois formatos. `11600D,N3` (vírgula sem espaço) é erro: use vírgula + espaço.
- Se a estrutura tem o parâmetro `poste` na base (`N3=N3(poste=11600D)`, grupos `N3(poste=%D)`), ela usa o **poste mais próximo à esquerda** na linha; sem poste à esquerda vale o padrão. O valor do parâmetro `poste` é sempre a forma compacta (`11600D`).
- **Operação M / *M**: o poste da linha é ignorado (só serve para o parâmetro) e todos os demais itens contam como `I` / `*I`. Vale para os dois formatos.
- Os Ajustes de planilhas não editam linhas em notação compacta.

## Projeto Pernambuco (081): regras do Leitor e cabos pelo LABEL

- **Operação:** pela **layer** do DXF (`01_INSTALANDO` = I, `01_REMOVER` = R, `01_REAPLICANDO` = M). Quando não há layer (PDF), vale a **cor exata**: `#000000` = instalar, `#000001` = remover. Textos de ruas, KML e outras cores são ignorados.
- **Poste:** `11300D`, `7150FV`, `10300R`, `9300F` entram como poste, do jeito que estão escritos (a base reconhece `11300D` = `DT11/300`).
- **Cabo:** o LABEL do projeto, com o comprimento na mesma linha ou na de baixo (`45 m`, `45m`, `45M`): vira `3xA35P(7,94)-13,8kV 45 m` (o texto fica como está escrito, com a tensão). Na tabela **Cabos** você também pode digitar assim:
  - `3x` = 3 fases (sem prefixo = 1); a **tensão** (`-13,8kV`, `-13,8KV`, `+20kV`…) continua escrita na tabela, mas é ignorada na Totalizadora.
  - **Neutro** `(…)`: `3xA35P(7,94)` = 3 `A35P` + 1 `CAZ7,94`; `3xC16(16)` = 3 `C16` + 1 `C16`; `3xS50P(35)` = 3 `S50P` + 1 `S35P`. O neutro entra na soma de metros do resumo.
  - `(S04)` e `[C16]` sozinhos são cabos de neutro; `16` sozinho = `C16`; `RAMA25`, `PB35`… = 1 fase; `C 3x1x95(95)` é um único código de ativo.
  - O nome do cabo vale como escrito (ex.: `S1/0`): ele precisa existir na base de orçamento.
- No editor de regras do Leitor há agora **Cor exata** (hex) e, na vizinhança, **Só os itens seguintes**.
- **Poste com itens:** `9200D` e `9200D, N1` são lidos como poste. No DXF, se o texto tem várias linhas (mesmo MTEXT), até **2 linhas de baixo** entram na mesma linha (`9200D, N1` + `N2` → `9200D, N1, N2`) e a de baixo não vira linha própria; uma 3ª linha segue sozinha. No editor há "Só linhas do mesmo MTEXT", "Juntar todas", "Obrigatória" e direção "só os anteriores".

## Vincular cabos a estruturas (modal em abas)

Botão **VINCULAR CABOS** do Resumo:
- **Vínculo:** cabos à esquerda, estruturas à direita (a linha `9200D, N1` aparece como a estrutura **N1** com o poste ao lado, só como contexto). Clique no cabo e depois na estrutura (ou o contrário) para ligar; clique na linha para desfazer. Cores: preto = sem vínculo · roxo = parcial · verde = completo · vermelho = inválido. Ao selecionar um cabo, as estruturas que podem recebê-lo brilham e as outras esmaecem (o motivo aparece ao passar o mouse).
- **Cabo fictício:** botão direito numa estrutura que ficou sem cabo → "Associar cabo fictício" (só aparecem os cabos que a estrutura aceita). Ele completa o vínculo e o parâmetro, mas não tem comprimento e não entra no orçamento.
- **Tabela de Cabos** e **Tabela de Estruturas:** dizem o que cada tipo aceita (a primeira linha que casa vale). O valor de `cabo` (A, P, C…) vem do cabo vinculado e preenche o parâmetro da estrutura (`N1(cabo=…)`); se você digitar o valor na linha, o digitado vale. Estrutura de lados livres com cabos diferentes usa o maior (C > P > A) e avisa.
- **Script:** JSON com passos (`proximidade` ou `sequencia`) para vincular automaticamente; use **Pré-visualizar** antes de **Aplicar**. Vínculo manual nunca é sobrescrito, a menos que o passo diga `"substituir": true`.
- **Mapa:** os mesmos itens pelas coordenadas do desenho (DXF).
- **Supabase:** rode o `scripts/schema_supabase.sql` (tabela `vinculacao_config`) para as tabelas e scripts irem para a nuvem.



## Ajustes de usabilidade (2026-10-27)

- **Resultado (SAP):** barra com "lado a lado/empilhado" (ODI em cima), minimizar ODI/ODD e "↔ Ajustar colunas" (duplo clique na borda do título ajusta uma coluna).
- **Linhas manuais** (REC/SAP) nascem em branco; nada de `NOVO_COD`.
- **Outros:** Delete numa linha vazia apaga essa linha e o foco sobe para a anterior. A coluna Entidade fica oculta; use "Exibir ▾" para mostrá-la, ordenar ou agrupar por entidade. Parâmetros do vínculo (`cabo=C`) aparecem como dica ao passar o mouse no Ativo.
- **Leitor:** a tabela de extração mostra a coluna Layer (depois de Cor).
- **Base de Orçamento:** seleção estilo Excel (arraste/Shift/Ctrl+A, Ctrl+C, Delete), "+ Linha acima/abaixo" da linha selecionada (sem seleção, no fim), Duplicar, Excluir selecionadas, "↔ Ajustar colunas" e menu do botão direito. Tudo só vale após "Salvar Alterações".


## Mapa com a rede e os postes do DXF (Pernambuco)

Ao extrair um DXF, o programa lê sozinho (sem criar linhas na tabela) as linhas da layer REDE_MT (cor vermelha) e REDE_BT (cor verde) e os blocos com "poste" no nome. A **cor decide** MT/BT e a **layer confirma**. A anotação do cabo é ligada à linha mais próxima (até 4× a altura do texto) ou, se houver uma linha-guia saindo do texto, à linha onde a guia termina; a estrutura é ligada ao poste mais próximo. Na aba **Mapa** do vínculo o cabo aparece como a linha da rede, a estrutura como o poste, e as coordenadas surgem ao clicar. Dá para ligar/desligar a rede ao fundo. Arquivos extraídos antes dessa mudança precisam ser extraídos de novo.


## Cabos: Fases, Tensão e Nível (TASK-092) e estruturas com fictício opcional (TASK-093)

- Na tabela de Cabos aparecem **Fases**, **Tensão** e **Nível** lidos do texto original (`3xA35P(7,94)-13,8kV 45 m` → 3 · 13,8kV · MT). Sem `Nx` conta 1 fase; MT = tensão acima de 1 kV; sem kV = BT. Dá para editar à mão e ocultar as colunas ("Ocultar Fases/Nível").
- Na **Tabela de Cabos** (modal de vínculo) uma regra pode filtrar por **Fases**, **Nível** e **Padrão no texto** (regex no texto original). A primeira regra ativa que casa tudo vale.
- Na **Tabela de Estruturas**: `Mín`/`Máx` definem a faixa de cabos e `Reais` quantos podem ser reais; o excedente entra só como **cabo fictício** (botão direito na estrutura). Exemplo SI3: Mín 1, Máx 2, Reais 1 — um cabo real completa a estrutura; o segundo, se houver, é fictício e não vincula o cabo.


## Várias estruturas na mesma linha (TASK-094)

Uma linha de Outros como `9400D, N1, N4` tem um poste e duas estruturas; no modal de vínculo cada uma é um bloco (N1, N4) na mesma linha do poste e recebe seus próprios cabos e cabos fictícios. Poste sem estrutura não aparece (a menos que já tenha cabo ligado); estrutura sem poste aparece sem o card do poste. Vínculos feitos antes continuam valendo (a 1ª estrutura mantém o mesmo identificador). O vínculo automático por script não usa duas estruturas do mesmo poste como as duas pontas de um cabo.


## Operação no vínculo e estrutura fictícia no cabo (TASK-095)

- Só entram no vínculo cabos e estruturas de operação **I, M, \*I e \*M**. Linhas **R** e **\*R** não aparecem no modal de vínculo, não contam como cabo/estrutura e não geram itens `CABO_ESTRUTURA` na Totalizadora.
- No modal, **botão direito no cabo → Estrutura fictícia**: escolhe uma estrutura (ou digita o código) para ocupar uma ponta do cabo sem existir uma linha em Outros. A estrutura precisa ser aceita pelas regras do cabo; ela segue para a Totalizadora como `CABO_ESTRUTURA`.


## Estrutura com dois cabos nos parâmetros (TASK-096)

- Na **Base de Orçamento** o mesmo parâmetro pode aparecer mais de uma vez: declaração `N4=N4(cabo, cabo)` e grupos como `N4(cabo=%A35P%, cabo=%C16%)`. A **ordem não importa**. Valores podem ter `/` (`S1/0`).
- Na **Tabela de Estruturas** (modal de vínculo) a coluna **Parâmetros (vários)** diz de onde vem cada valor: `cabo=ativo, cabo=ativo` (código do cabo, sem 3x e sem tensão), ou `tipocabo=valor_cabo` (tipo da regra do cabo). A 1ª ocorrência recebe o 1º cabo vinculado e a 2ª o 2º; cabos iguais repetem o valor. Com só um cabo, só o primeiro é preenchido e a estrutura fica pendente.
- Em uma linha `9400D, N1, N4` cada estrutura recebe só os parâmetros dela. Pernambuco: N4, B4-MT, CE4 e U4 já vêm configuradas no ponto de partida (configurações já salvas precisam ser ajustadas na tabela).


## Ativo com espaço no nome na tabela Outros (TASK-097)

Na notação compacta o **único delimitador é vírgula + espaço**: `9400D, CE3 PR, N1` tem três itens e `CE3 PR` vai inteiro para a Totalizadora. Quantidade vem antes, separada por espaço (`2 CE3 PR` = 2 unidades de "CE3 PR"). Uma linha sem vírgula e sem nenhum `qtd-ativo` (como `1-N1`) também é compacta: `CE3 PR` sozinho é um item só. Se algum item tiver o formato antigo (`1-N1 2-U4`), a linha segue o formato antigo e os itens se separam por espaço.


## Tensão do cabo com ponto ou vírgula (TASK-098)

O Leitor de Pernambuco e a notação do cabo aceitam a tensão com vírgula ou ponto (`13,8kV` ou `13.8KV`), em maiúsculas ou minúsculas. Se as regras do projeto 081 já tiverem sido editadas à mão, troque `(?:,\d+)?\s*KV` por `(?:[.,]\d+)?\s*KV` nas duas regras de cabo (ou restaure o ponto de partida).


## Vínculo por coordenadas com reserva por sorteio (TASK-099)

Na aba **Script** do vínculo, o botão **Novo (coordenadas + sorteio)** cria um script com os passos por coordenadas (M na amarração mais próxima, I entre as 2 estruturas mais próximas, multiplex em sequência) e, no fim, dois passos de **reserva**: cabos que as coordenadas não resolveram (sem coordenada, fora do raio ou parciais) são ligados por **sorteio** entre as estruturas que as tabelas aceitam. O sorteio é reproduzível (campo `semente`; mudar a semente muda o sorteio) e prefere estruturas ainda vazias (`"puro": true` desliga). `"completar": true` faz o passo só acrescentar o que falta aos cabos parciais, mantendo o que as coordenadas já ligaram. Confira sempre na **Pré-visualização**: o que vem do sorteio não tem base geométrica.


## Cores por TIPO MDO no Resultado (TASK-100)

No REC e na SAP, as linhas de **mão de obra (N / MAO-DE-OBRA)** têm fundo cinza escuro mais escuro e as de **material (M / MATERIAL)** têm cinza escuro mais claro. A cor acompanha o que você edita no TIPO MDO. No modo comparação valem as cores da comparação (verde/amarelo/vermelho/azul).


## `cabo` e `tipocabo` nos parâmetros das estruturas (TASK-101)

- **`cabo`** = a descrição completa do cabo, sem o comprimento: `3xS04-7,97kV 45 m` vira `3XS04-7.97KV` (maiúsculas, sem espaços, vírgula decimal → ponto, parênteses → `_`). Na base você pode usar curingas: `N4(cabo=%S04%, cabo=%C16%)`; a vírgula decimal da tensão pode ser escrita como `7,97KV`.
- **`tipocabo`** = o valor da coluna **Valor de cabo** da Tabela de Cabos do modal de vínculo (C, P, A…).
- Na Tabela de Estruturas: **Parâmetro** (um só) recebe o Valor de cabo — use `tipocabo`; **Parâmetros (vários)** aceita `cabo, cabo, tipocabo, tipocabo` (um valor por cabo vinculado).
- **Estrutura caindo no padrão (ex.: A):** o Valor de cabo vem da PRIMEIRA linha ativa da Tabela de Cabos que casa o código do cabo. Cobre cadastrado depois de uma regra de alumínio que também casa (ex.: `S04` casa `^(S|AA|CAA)\d+`) vira A; sem nenhuma regra o valor fica vazio e vale o padrão da declaração na base. Teste o cabo na caixa "Testar um cabo" e ponha a linha de cobre antes.


## Declaração com condições na Base de Orçamento (TASK-102)

A linha de declaração pode trazer condições e valer como o primeiro grupo: `N1=N1(poste=%D, tipocabo=A)` declara os parâmetros `poste` e `tipocabo`, define `A` como padrão do `tipocabo` e vale só para postes que terminam em D. Curingas (`%`), alternativas (`|`) e exceções (`!=`) na declaração são só condição, sem padrão. Declaração só com valores concretos (`N1=N1(cabo=XLPE, poste=DT)`) continua sendo o grupo comum, como antes. O apelido pode ficar na mesma célula: `B1MT=B1MT(poste=%D,tipocabo=A), B1-MT`. No vínculo: `cabo` = descrição do 1º cabo, `cabo2` = a do 2º, `tipocabo` = o tipo (o maior, se os cabos diferem).

## Formato antigo de Outros: o que é descartado (TASK-105)
No formato antigo (`1-U3 2-N1`) o `-` só separa a **quantidade** do ativo. Um ativo com hífen (`1-U3-TM`) ou um item sem par (`1-U3 TM`) era lido pela metade. Agora **Montar Orçamento** lista o problema em "não encontrados" (o que se entende, `U3`, continua calculado) e a validação já o marca como erro. Para nomes com hífen ou com espaço use a notação compacta: `U3-TM, 2 N1`.

## Nome de ativo com vírgula, espaço ou outro caractere: `[nome]` (TASK-108)
Na Base de Orçamento a vírgula separa **apelidos** (`DT11/600, 11600D`), então um nome como `S336,4` ou `CAZ7,94` seria cortado. Para o nome valer **do jeito que está**, cadastre-o **entre colchetes** na coluna ATIVO: `[S336,4]`, `[CAZ7,94]`, `[CAZ 7,94]`, `[A-B]`.

- Dentro dos colchetes vale qualquer caractere (vírgula, espaço, hífen, `=`, `%`…), **menos** `(`, `)` e `[ ]`. O ativo entre colchetes é simples: **não aceita parâmetros**.
- O nome oficial é o conteúdo, sem os colchetes. Apelidos continuam depois da vírgula: `[S336,4], 11336D`.
- Em **Outros** use o mesmo `[nome]`: `2-[CAZ 7,94]`, `[S336,A], 3 N1`. Os colchetes protegem a vírgula, o espaço e o `-`.
- O neutro do cabo (`3xA35P(7,94)` → `CAZ7,94`) já casa um ativo cadastrado como `[CAZ7,94]`. Nas regras da Totalizadora, `ATIVO_DE`/`ATIVO_PARA` aceitam colchetes e `CAZ7,94` também casa `[CAZ7,94]`.
- Colchete vazio, aberto, com texto fora ou com parênteses dentro é recusado na importação da base (a mensagem diz a linha).
- Sem colchetes nada muda (`S336,4` continua sendo `S336` com apelido `4`).

## FATOR em fórmula (Base de Orçamento, TASK-109)
As colunas **FATOR I** e **FATOR R** aceitam um número (`1,5`) **ou uma fórmula que começa com `=`** e usa os **parâmetros** do ativo. Assim uma única linha cobre variantes como `CERCA(fios=%F)` com `=NUM(fios)`: o pedido `10-CERCA(8F)` gera 80, `10-CERCA(12F)` gera 120.

- **Parâmetro:** pelo nome, sem diferença de caixa (`fios`, `cabo`, `tipocabo`, `poste`); `cabo#2` é o 2º `cabo` de um ativo com o parâmetro repetido.
- **Número do texto:** `NUM(fios)` pega o 1º número de `8F` → 8 (`12,5m` → 12,5). Um parâmetro só entra direto em conta se for número puro (`8`); com `8F` use `NUM`.
- **Operadores:** `+ - * /`, parênteses e comparações `= <> < > <= >=` (texto sem diferença de caixa).
- **Funções** (argumentos separados por `;`): `NUM(x)`, `ABS(x)`, `MAX(a; b; …)`, `MIN(a; b; …)`, `ARRED(x; casas)`, `CIMA(x)`, `BAIXO(x)`, `SE(cond; então; senão)`, `E(…)`, `OU(…)`, `NAO(x)`.
- **Exemplos:** `=NUM(fios)/8` · `=SE(tipo="A"; 2; 1)` · `=MAX(NUM(cabo); 35) * 2` · `=SE(NUM(fios)>=10; 2; 1)`.
- **Com o ARRED:** a coluna ARRED (TASK-107) arredonda o resultado da fórmula (ex.: `=NUM(fios)/8` + `CIMA`).
- **Pedido sem o parâmetro:** a fórmula não pode ser calculada; o cálculo avisa e a linha não soma (as demais linhas do ativo entram normalmente). Padrão na declaração (`CERCA=CERCA(fios=8F)`) vale na fórmula.
- **Na importação:** fórmula com **erro de sintaxe** é recusada com a linha e nada é alterado. Fórmula que usa um **parâmetro que o ativo não declara** salva normalmente e gera só um **aviso** (o parâmetro pode existir apenas na fórmula). Esse **parâmetro livre** é informado **por nome** no pedido: `10-CERCA(fios=8F)` (ou junto dos declarados: `10-CERCA(A, fios=8F)`); por posição não vale (`CERCA(8F)` pede o nome). Sem informá-lo, o cálculo avisa `não informado no pedido` e a linha não soma. Um parâmetro livre aceita qualquer valor (não há lista cadastrada para conferir). No CSV, fórmula com `;` precisa estar entre aspas (a exportação da tela já faz isso).
- Depois de atualizar, o administrador roda o `schema_supabase.sql` no Supabase (cria `formula_i` e `formula_r`). Linhas existentes continuam com o número.

## Ativo composto: agrupar ativos (TASK-110)
Para pedir vários ativos com um nome só (ou dar um apelido a configurações de ativos com parâmetros), cadastre uma linha na Base com ATIVO no formato **`NOME:=item + qtd*item`**:

| ATIVO | DESC ATIVO | COMPONENTE | CÓDIGO |
|---|---|---|---|
| `N1C:=N1(C) + 2*N2(C)` | N1 tipo C com 2 N2 tipo C | (vazio) | (vazio) |

- Pedir `3-N1C` vale como `3-N1(C)` e `6-N2(C)` (a quantidade pedida multiplica a de cada item). O texto continua `N1C` na Totalizadora; o cálculo expande.
- **Itens:** ativo simples (`T3`), com parâmetros (`N1(C)`, `N1(poste=DT)`), nome literal (`[CAZ 7,94]`) ou **outro composto**. Quantidade antes, com `*` e **ponto** decimal (`0.5*T3`; a vírgula separa apelidos). Separe os itens com `+`.
- **Apelido:** `KIT:=N1C + T3, K` (o apelido `K` também pede o kit). Nome com caracteres especiais: `[S 1]:=A+B`.
- **Só definição:** CÓDIGO e COMPONENTE ficam **em branco** (os códigos vêm dos itens). O nome precisa ser diferente de qualquer ativo da base.
- **Itens completos:** os parâmetros do item devem estar no próprio item (ou usar o padrão da declaração); o composto não repassa parâmetros do vínculo cabo⇄estrutura.
- **Por projeto:** o composto do projeto vence o genérico com o mesmo nome.
- **Importação:** é recusada (com a linha) se o composto tiver CÓDIGO/COMPONENTE, nome repetido, quantidade ≤ 0, item que não existe na base, parâmetro inválido (`N1(Z)` com `Z` não cadastrado) ou ciclo (`A:=B`, `B:=A`).
- **Regras da Totalizadora** casam o nome do composto (`N1C`), não os itens.

## Tela Resumo do Orçamento: novo layout (TASK-111)
- **Barra de título:** nº da obra, **Ler projeto** e **Montar orçamento** ao lado; depois **Carregar obra**, **Salvar obra**, **Salvar como modelo**, **Importar JSON** e **Exportar JSON** (o `.obra.json` da obra que está na tela, sem precisar salvar antes). Os botões **Configurações** e **Assistente IA** abrem os painéis laterais.
- **Painéis laterais** (ocultos por padrão): **Configurações** (esquerda: projeto, Tabela Orçamento, Pesquisar ativos, Chave API, restaurar layout) e **Assistente IA** (direita: o chat). O 📌 **fixa** o painel (ele empurra as tabelas); solto, ele flutua sobre as tabelas e fecha ao clicar fora ou com Esc. O estado fica lembrado neste navegador.
- **Barra de modais** (▾ Modais recolhe e mostra): Resumo da rede, Ajustar, Validar, Contexto, Regras de validação, Postes e estruturas, Cabos, Conexões, Ramais, Vincular cabos e os modais montáveis. Um **resumo compacto** (postes, média, baixa, equipamentos) fica sempre na barra de controles.
- **Tabelas:** botões **Cabos** e **Demais ativos** mostram/ocultam cada tabela; **⌄ Ocultar** recolhe só o corpo; **Lado a lado** ou **Empilhadas**. **Padrão/Totalizadora** continua escolhendo a visão.
- **Linhas vazias estilo Excel:** as tabelas Cabos e Demais ativos mostram linhas em branco até o fim da área visível. **Digite direto** no Ativo de uma linha em branco (ela vira linha da tabela, com Ctrl+Z) ou **cole várias linhas** do Excel/e-mail (a colagem começa na primeira linha em branco). Não precisa do botão **Add**.
- **Operação das linhas novas e coladas:** se a tabela está vazia, **I**; senão a **mesma operação da linha preenchida mais próxima acima**. Linhas coladas no formato `operação<TAB>ativo` (por exemplo `*R` e `9400D, B1-MT`) mantêm a operação que veio; as que trazem só o ativo herdam a da linha de cima.

## Parâmetro usado em fórmula aceita qualquer valor (TASK-114 e TASK-115)
Normalmente um parâmetro **só aceita os valores cadastrados** nos grupos (para pegar erro de digitação) e o padrão da declaração exige um grupo que o aceite. Quando o parâmetro é **usado em alguma fórmula de FATOR do ativo** (`=SE(tipo="PR";3.9;…)`), os valores vêm da fórmula e o parâmetro passa a aceitar **qualquer valor**, mesmo que existam grupos que citem só alguns (`ATERINT(tipo=MRT%)`, `ATERINT(tipo=MRT_RDC)`): o pedido `ATERINT(PR)` calcula com as linhas comuns e a fórmula, e só ganha as linhas de um grupo quando o valor casa com ele (`MRT_RPN` → grupo `MRT%`). A importação também não exige grupo para o padrão (`CERCA=CERCA(descr=5F, tipo=PA)`).

- **Consequência:** um valor digitado errado (`RPM` em vez de `RPN`) não é mais recusado; ele cai no `SENÃO` das fórmulas. Se quiser conferir os valores, tire o parâmetro das fórmulas ou use um parâmetro separado só para a conferência.
- Parâmetro **fora** de fórmulas continua fechado nos valores dos grupos, e padrão sem grupo e sem uso em fórmula continua sendo erro de importação.

## Resumo da rede: comprimentos (TASK-116)
O resumo (Postes, Média, Baixa, Equip.) conta só linhas de operação **I**. **Média** = cabos com tensão (`-13,8kV`) ou do formato antigo com prefixo `CAA`, `CAL` ou `P`; **Baixa** = sem tensão ou `M2X`/`M3X`.
- **O neutro não entra na Média:** `3xA35P(7,94)-13,8kV 45 m` soma 45 m (o cabo principal), não 90 m. Na Baixa o neutro continua somando.
- **Cabo sem comprimento na própria linha:** vale o comprimento da linha de baixo, seja ela só o número (`45`, `45 m`, `7,5m`) ou outro cabo com comprimento. Exemplos: `3xA35P(7,94)-13,8kV` + `45 m`; `CAA2 ABC` + `30`; `P 50` + `CAZ9,5 abc 50 m` (os 50 m entram na Média para o `P50`; a linha do `CAZ9,5` continua sendo contada por si mesma, e `CAZ` não entra em Média nem Baixa). Se a linha de baixo não tiver comprimento, o cabo não soma. A operação da linha de baixo não importa. (TASK-118)

## Limpar Tudo e Modais do projeto (TASK-117)
**Limpar Tudo** (barra de controles) agora esvazia, além de Cabos e Outros (com os cabos fictícios do Vincular cabos), os modais que guardam dados: **Ramais** (itens e badge vermelho), **Conexões** (quantidades digitadas) e os geradores de **Postes** e **Cabos**. Os modais montáveis começam vazios a cada abertura, então não há o que limpar neles. A confirmação lista o que será esvaziado e **Ctrl+Z** devolve tudo (tabelas, Ramais e Conexões; Refazer limpa de novo). Se as tabelas já estavam vazias e só os modais tinham dados, a limpeza vale, mas não dá para desfazer (a confirmação avisa).

**Modais do projeto** (painel ⚙ Configurações): lista os botões da barra de modais deste projeto.
- 👁 oculta/mostra o botão **só neste projeto** (nada é apagado); ▲▼ muda a ordem. Os modais do sistema (Resumo da rede, Ajustar, Validar e opções, Contexto, Regras de validação, POSTES E ESTRUTURAS, CABOS, CONEXÕES, RAMAIS, VINCULAR CABOS) só ocultam/reordenam.
- Os **montáveis** também podem ser editados (✏️, abre o gerenciador), **excluídos** (🗑, com confirmação; a versão anterior fica no histórico do gerenciador) e criados (**＋ Novo modal montável**).
- **⧉ Copiar de outro projeto…** copia montáveis (id repetido ganha sufixo `_2`; quem estava oculto na origem fica oculto) — útil para levar da Paraíba para Pernambuco.
- **↺ Restaurar barra padrão** mostra tudo na ordem original.
- Salvo por projeto e vale para todos os usuários do projeto; **só o administrador altera** (os demais veem a lista). Não precisa de SQL no Supabase (usa a tabela `atalhos_modais`).

## Importar JSON direto na tela (TASK-118)
**⤒ Importar JSON** (barra de título) abre o arquivo `.obra.json` e mostra um diálogo com a obra (nome, nº de linhas) e três ações: **Carregar** (substitui as tabelas da tela), **Adicionar** (acrescenta as linhas no fim) e **Subtrair** (acrescenta como retirada: cabos com quantidade negativa e `*` na frente dos ativos de Outros, o mesmo "Subtrair" do Carregar Obra). Tudo pode ser desfeito com **Ctrl+Z**. Nada é gravado no servidor; o arquivo precisa ser do projeto selecionado (senão a tela explica e não importa). **Modelo** com variáveis V continua sendo importado como modelo salvo e aparece em Carregar Obra. O botão "Importar obra" dentro de Carregar Obra segue salvando uma cópia particular da obra.

## Rede, proxy e modo offline (TASK-121) — só no programa desktop
O programa **se adapta sozinho à rede**, sem script e sem mexer em variáveis (a versão online/servidor não usa nada disto):
- **Ao iniciar**, antes de qualquer conexão: lê o proxy configurado (variáveis `HTTPS_PROXY`/`HTTP_PROXY` e a configuração do Windows) e **testa o Supabase de verdade** pelo caminho com proxy e pelo direto (em paralelo, ~3 s). Usa o que responder (prefere o proxy configurado quando ele funciona) e aplica a escolha ao programa todo, inclusive ao chat de IA. Proxy "esquecido" (casa) → vai direto; empresa → usa o proxy. O log mostra: `Rede ao iniciar: ONLINE — acesso direto | proxy configurado: …`.
- **Com o programa aberto**, uma vigia leve reavalia sozinha (a cada ~30 s online; a cada 5–30 s offline) e **na hora** quando uma consulta à nuvem falha por rede. Trocou de rede (casa ↔ empresa)? Ela troca de caminho sozinha. Nenhum caminho responde? Entra em **modo offline**.
- **Modo offline:** o programa vai **direto ao banco local**, sem esperar timeout (as telas respondem em milissegundos). A tela mostra o aviso vermelho *"Sem conexão com a nuvem — modo local (cópia de dd/mm hh:mm)"*. Dá para desligar a internet e continuar trabalhando (inclusive o login, com as mesmas contas e senhas).
- **Cópia local do banco:** a cada início (em segundo plano) e sempre que a nuvem volta, o banco local é posto igual ao da nuvem: projetos, obras, REC, regras (leitor, vinculação, domínio, conversão, IA), modais/barra, contextos, ajustes, prompts, históricos, **contas de usuário** (o login offline usa o mesmo `user_id` do online) e a base de orçamento (na ordem da planilha). Ficam só no computador (não são copiadas nem enviadas): sessões, logs, a cópia pessoal da base e `configuracoes` (tem a chave da API).
- **O trabalho feito offline não se perde:** na cópia seguinte o programa compara local × nuvem × a "foto" da última cópia. O que você criou/editou/apagou só aqui **é enviado à nuvem**; o que mudou só na nuvem vem para cá; se **os dois mudaram a mesma linha**, a nuvem prevalece e a sua versão fica guardada na tabela `conflitos_locais` do banco local (motivo e data). Na **primeira** cópia de cada tabela a nuvem prevalece e o que só existia aqui também fica em `conflitos_locais`. Se o envio falhar, a mudança continua no local e é reenviada na próxima cópia.
- **Desligar:** `REDE_ADAPTATIVA=0` no `.env` volta ao comportamento antigo. Fora de escopo: proxy com usuário e senha e arquivo PAC (`AutoConfigURL`).
- **Ver o estado:** `GET /api/rede/estado` (online, caminho, proxy sem credenciais, data da cópia, pendências). `POST /api/rede/copiar` dispara uma cópia agora.

## Nome do ativo, descrição e parâmetros (TASK-122)
- **A descrição não resolve nome:** o cálculo procura o nome do ativo só em **ATIVO (e apelidos) → COMPONENTE → CÓDIGO**. A coluna DESC ATIVO é só texto de apoio: um nome que só existe como descrição (ex.: `FLABIT`) **não é encontrado**, com ou sem parâmetros. Para usá-lo, cadastre-o como ativo, apelido (`IT_2=IT_2(...), IT-2, FLABIT`), componente ou código. A Totalizadora continua mostrando a descrição do ativo achado, e o modal "Pesquisar ativos" continua buscando por descrição.
- **Validação separada em nome e parâmetros:** o aviso **C1-BASE** avalia só o **nome** (antes do `(`): `FLABIT(D,35)` → "Ativo 'FLABIT' não encontrado na base técnica". Quando o nome existe mas os parâmetros não conferem, o aviso é o **C1-PARAM**: valor não cadastrado, parâmetro desconhecido, valores a mais/menos, ou "ativo sem parâmetros declarados".
- **Ajuda ao digitar `(` na tabela Outros:** ao digitar `(` depois de um nome (`1-IT_2(`, `2 IT-2(`, `9400D, IT_2(`), uma caixa mostra os **parâmetros do ativo** na ordem da declaração, com o padrão e os valores cadastrados (mais os padrões `%…`, que só valem como condição). O parâmetro da posição atual fica destacado e dá para **clicar num valor** para inseri-lo; depois de uma vírgula a caixa passa para o parâmetro seguinte (ou o da `nome=` digitado). Parâmetros que só aparecem em fórmula vêm marcados "informar por nome". Sem declaração: "**ativo sem parâmetros declarados**"; nome desconhecido: "**ativo não encontrado na base**". Esc fecha. Usa a base já carregada na tela (funciona offline).

## Totalizadora: nome e parâmetros avaliados separados (TASK-123)
Ao digitar/editar o ATIVO (ou a ORIGEM) de uma linha da Totalizadora, a tela avalia o **nome** (antes do `(`) e depois os **parâmetros**, em separado:
- **Vermelho = o nome não existe** na base (ativo/apelido → componente → código, para o projeto e a origem da linha; a descrição não vale). A descrição fica vazia.
- **Âmbar = o nome existe, mas os parâmetros têm problema**; a mensagem aparece na coluna DESCRIÇÃO: **fora do escopo** (`POSTE=XX não cadastrado em IT_2 (existem: D, %D, %R…)`), **não declarado** (`parâmetro CABOT desconhecido`, `ativo sem parâmetros declarados`) ou valores a mais (`parâmetros demais para IT_2 (esperado: POSTE, CABOBT)`).
- Sem cor = nome e parâmetros ok; a descrição é a da linha do nome (`FLABIT(D,35)` mostra a descrição do `FLABIT`). Antes a edição comparava o texto inteiro (`FLABIT(D,35)`) com a base e ficava vermelha.
- A **validação** (janela "Problemas encontrados nas planilhas") agora confere contra a **mesma base do cálculo** (a oficial); antes lia a cópia pessoal de "Salvar Alterações", que ficava desatualizada e dizia "ativo não encontrado" para o que a tela via. O aviso **C1-PARAM** também vale para pedidos escritos com **apelido** (`FLABIT(D,A,35)`).



## Poste com nome compacto no cadastro (TASK-124)
O poste da linha de Outros é reconhecido em qualquer destes nomes: `DT11/600`/`RC12/1000`/`CV…`/`FV…` (antigo), `R12/1000` (abreviado) e `121000R` (compacto, inclusive quando é o ATIVO cadastrado na base). Qualquer um deles preenche o parâmetro `poste` das estruturas à direita na mesma linha (`121000R, N1` → `N1(121000R,A)`), no formato compacto e no antigo (`1-121000R 1-N1`). Não precisa renomear o poste na planilha.


## Modal de vinculação: o que entra da tabela Outros (TASK-125)
O modal lista as linhas de Outros classificadas como ESTRUTURA/POSTE **e** as ainda sem classificação (entidade `0`) que tenham ativo. Linhas com outra entidade escolhida (CHAVE, TRAFO, RAMAIS, IP…), sem ativo, ou com operação R/*R não aparecem. Linha sem regra na Tabela de Estruturas aparece como "sem regra".


## Modal de vinculação sem cabos (TASK-126)
As estruturas da tabela Outros aparecem no modal mesmo sem nenhum cabo na tabela de cabos, para receber cabo fictício (botão direito na estrutura).


## Publicar uma nova versão do programa (TASK-127)
1. Faça o commit das alterações e rode `python scripts/release.py 2.1.0` (atualiza a versão e cria a tag `v2.1.0`).
2. `git push origin main --tags`. O GitHub Actions compila o `Leitor_Projetos.exe` no Windows e publica a Release (com o arquivo `.sha256`).
3. Os programas instalados consultam a Release ao abrir: aparece "Atualização Disponível" e, em **Instalar e Reiniciar**, o programa baixa, confere o SHA-256, troca o `.exe` e reabre sozinho. O banco e o `.env` ficam como estão.
`ATUALIZADOR=0` no `.env` desliga a verificação.


## Programa desktop: telas dentro da janela (TASK-128)
No programa desktop nada abre em nova aba do navegador: **Montar Orçamento**, **Tabela de orçamento** e o **Manual** abrem por cima da tela atual (botão **Fechar**/Esc ou **Voltar**; o Resumo continua como estava) e **Processar** no Leitor segue para o Resumo na mesma janela. Na versão online (navegador) continua abrindo nova aba. Ao abrir, o programa mostra "Iniciando…" até o servidor responder, e as chaves de API são campos mascarados que o navegador não oferece para salvar.


## Parêntese automático e descrição/filtro do REC (TASK-129)
- Na coluna Ativo de Outros, ao terminar de digitar parâmetros (`N1(D, X`), **Enter** ou **sair do campo** fecha o `)` que falta (`N1(D, X)`).
- No Resultado do Orçamento, ao **somar/subtrair/carregar um REC**, colar itens ou usar o **modo comparação**, a descrição e o **filtro** de cada código vêm da Base de Orçamento (a mesma do cálculo; as linhas do projeto selecionado têm preferência). Código que não existe na base fica com "Descrição indisponível" e filtro vazio.


## Composto com parâmetros e grupo que chama ativos (TASK-134)
O ativo composto (`NOME:=item + qtd*item`) aceita **parênteses no nome**:
- **Composto com parâmetros** — `CF(cabo, cabo2=S04):=CHAVE_FACA(cabo,cabo2) + TERM(cabo) + TERM(cabo2)`. `CF` não é um ativo da base. O pedido `CF(S04,A35P)` (ou `CF`, com os padrões, ou por nome `CF(cabo2=A35P,cabo=S04)`) passa os valores aos itens: **no item, o nome do parâmetro vira o valor** (`TERM(cabo)` → `TERM(S04)`). Os parâmetros do composto aceitam qualquer valor; o vínculo cabo⇄estrutura preenche `cabo`/`cabo2` nele como num ativo comum.
- **Grupo que chama ativos** — `CHAVE_FACA(cabo=%S04%):=TERM(cabo)` quando `CHAVE_FACA` já é um ativo da base. Quando o pedido da `CHAVE_FACA` casa as condições (as mesmas dos grupos: `%`, `|`, `!=`), os itens entram no cálculo, com os valores pedidos no lugar dos nomes. Várias linhas para o mesmo ativo somam. Os valores usados nas condições também contam como cadastrados no ativo.
- Regras: linha só de definição (CÓDIGO e COMPONENTE em branco); os itens podem ser outros compostos (sem ciclo); condição só com parâmetros que o ativo declara; use nomes diferentes (`cabo`, `cabo2`) para chamar cada cabo.
- Compacto: terminais dos dois cabos da chave = `CHAVE_FACA(cabo=%S04%):=TERM(cabo)` e `CHAVE_FACA(cabo2=%S04%):=TERM(cabo2)` (um par de linhas por terminal), no lugar de uma linha de código por combinação.


## Sugestão de ativos ao digitar e ativo de outro projeto (TASK-136)
- Ao digitar na coluna **Ativo** (Demais ativos e Totalizadora) aparecem as sugestões — só **ativos e apelidos** que existem na Base de Orçamento para o **projeto atual** (e os genéricos). ↑/↓ escolhem, **Tab** aceita (o 1º, se nada estiver destacado), **Enter** aceita o destacado, **Esc** fecha. O botão **💡 Sugestões**, na barra do Resumo, liga/desliga (a escolha fica guardada).
- Um ativo cadastrado só em **outro projeto** não vale: não é sugerido, não é reconhecido e, no cálculo, aparece em "não encontrados". Ativos sem projeto (genéricos) valem em qualquer projeto.
