# Manual: regras de validação e ajustes

Este manual explica, passo a passo, como cadastrar **regras de validação** e **ajustes** no OrçaCore. Não é preciso saber programar: quase tudo é feito pelo editor visual. Os exemplos em formato de código (JSON) são para quem preferir editar direto, e cada um deles é testado automaticamente.

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

No **Resumo do Orçamento**, clique em **Regras de validação**. A gaveta lateral tem as abas:

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

Na tela de resultado há um alternador **REC | SAP**. O **REC** é a tabela única de sempre. Na **SAP** os mesmos itens aparecem em duas tabelas: **ODI** (tudo que está instalando) e **ODD** (tudo que está removendo), com TIPO MDO abreviado (**M** = material, **N** = mão de obra) e uma coluna **CENTRO** (**1600** para material, **1000** para mão de obra — acompanha o TIPO MDO). A ordem das colunas é TIPO MDO, CÓDIGO, QUANTIDADE, CENTRO e DESCRIÇÃO (o Copiar e o Excel seguem a mesma ordem). As tabelas têm os mesmos campos editáveis, filtros, busca e ordenação; **⇄** move uma linha entre ODI e ODD; **Copiar** (acima de cada tabela) copia só aquela tabela. **Exportar Excel** (só na SAP) gera um arquivo com as planilhas ODI e ODD. Os botões **Copiar** de ODI e ODD levam TIPO MDO, CÓDIGO, QUANTIDADE e CENTRO (sem a descrição). Também dá para **selecionar células** (arrastando ou com Shift+clique) e copiá-las de uma vez com **Ctrl+C** ou **Copiar seleção**; o que já foi copiado fica **marcado em verde-azulado com ✓** (a marca some se o valor mudar; **Limpar marcações** recomeça). **Estilo Excel (TASK-164):** 1 clique seleciona a célula, **2 cliques / F2 / Enter** editam (Enter confirma, Esc cancela; digitar já substitui), arrastar, Shift+clique ou Shift+setas ampliam, o número da linha (`#`) seleciona a linha inteira (não é copiado) e Ctrl+A seleciona a tabela. Cada tabela mostra quantas linhas estão selecionadas e quantas faltam. **Copiar em blocos:** o botão **▶ Copiar próximas N** copia as primeiras N linhas (padrão 10, o número é editável), depois as N seguintes, até o fim, com barra de progresso (◀ volta um bloco; ↺ Recomeçar). Em **Colunas da cópia** escolha o que vai (padrão TIPO MDO, CÓDIGO, QUANTIDADE e CENTRO) e em **Formato da cópia** escolha Excel (recomendado; igual ao que o Excel copia, com quebra de linha em todas as linhas), Só texto ou LF (como era antes). A SAP é apenas outra forma de ver: salvar no banco e exportar o `.rec` continuam exatamente como no REC.

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


### Equipamentos que leem os cabos do poste (RL, CFU, para-raio, chave-faca)

Na **Tabela de Estruturas** (modal Vinculação), uma linha com **Ocupa desmarcado** é um **equipamento**: não consome ponta de cabo, não conta capacidade e não disputa cabo com a estrutura que segura os cabos (B4MT_RL, N4, CE-4…). Com **Herda = POSTE**, ele recebe `cabo`/`cabo2` (os parâmetros da coluna *Parâmetros*) dos **cabos do poste**: os que chegam no poste pelo desenho, os ligados às estruturas da linha e os fictícios. Em `12600D, B4MT_RL, RL`, a B4MT_RL liga os cabos como sempre e o RL fica `RL(poste=12600D, cabo=…, cabo2=…)` sem digitar nada.
- **Nível** (MT/BT) filtra os cabos; **Ordem = MAIOR_PRIMEIRO** põe o cabo de maior seção em `cabo` (tronco) e o menor em `cabo2` (ramal); **Distintos** junta cabos iguais (o mesmo cabo dos dois lados vira só `cabo`); **Se faltar** = AVISAR (parâmetro no padrão da base) ou BLOQUEAR (vermelho). *Distintos* e *Ordem* também valem para estruturas que ocupam (ex.: CE2.CE3 com tronco ×2 + ramal).
- No modal, o equipamento aparece em **cinza tracejado** ("lê do poste 12600D: A185P, A35P") e não recebe cabo por clique. As regras prontas do 081 estão em `data/vinculacao_estruturas_081.json` (importar na tabela).

### Avisos "grupos que podem valer juntos" (importação da base)

O aviso *"Código X aparece nos grupos A e B, que podem valer juntos: a quantidade soma duas vezes"* só sai quando os dois grupos podem mesmo valer no mesmo pedido. Dois padrões do tipo **contém** (`%30KVA%` × `%10KVA%`, `%A35%` × `%S1/0%`) **não** valem juntos (o valor é um código só); `%A35%` × `%A35P%` valem. Quando um ativo tem **mais de 5** avisos assim, a importação e o modal mostram **um aviso resumo** (ex.: *"RL: 61 códigos aparecem em grupos que podem valer juntos…"*), com a lista completa recolhida. Se a soma é de propósito (material por lado, como no RL), ignore.

### Estrutura fictícia do cabo: componentes avulsos (alças, conectores)

Quando o cabo é ligado a uma **estrutura fictícia** (o cabo novo ancora numa N4 existente, que não está na obra), o modal **Vinculação** mostra, **abaixo do cabo**, o que essa estrutura usa para aquele cabo: `N4 (fictícia) → 3 × ALCA(cabo=3XA35P…)  [+ Outros]`. São os itens dos grupos da base que chamam ativos (`N4(cabo=%,cabo2=%):=3*ALCA(cabo) + 3*ALCA(cabo2)`), **só do lado do cabo novo** (a alça do cabo existente fica de fora). Itens que dependem do cabo existente também (ex.: `CON(cabo, cabo2)`) aparecem com `cabo2=?` e **"complete o cabo existente"** — não vão para Outros.
- **+ Outros** cria uma linha em Outros por item (você pode editar ou apagar); o botão vira **adicionado ✓**. A quantidade vai no texto (`3-ALCA(cabo=…)`) para o cálculo contá-la. Nada entra no cálculo sem o clique.
- Mudou o cabo ou a estrutura fictícia? A sugestão volta e as linhas antigas ficam em laranja ("avulso desatualizado"). Estrutura sem linhas de chamada na base (ex.: N1): sem sugestão. A marca fica no `.obra`: reabrir mantém o "adicionado".

### Regras novas da Tabela de Estruturas (derivação, cabos por vão e parâmetros automáticos)

- **Cabo fictício e nível:** o cabo fictício de uma estrutura é MT ou BT conforme a primeira regra da **Tabela de Cabos** que casa com o código **e aceita aquela estrutura** (S04 numa LN3 = BT nua; numa N3 = CAA MT). Por isso o RL/para-raio (só MT) não leem um fictício BT.
- **Os lados = DERIVACAO** (SDBIT, SDBIM): 1 cabo de **derivação** (obrigatório) + até 2 do **tronco**, **iguais entre si**. O 1º parâmetro (`cabobt`) é sempre a derivação, em qualquer ordem de ligação. Só tronco = "falta o cabo da derivação"; tronco diferente = vermelho; 2 cabos diferentes = o 1º ligado vira a derivação e avisa.
- **Cabos por vão** (IBIT, IBIM): uma regex por vão (`["^PB\\d+$", "^C\\d+$"]`); cada vão precisa de **um cabo diferente** (um multiplex e uma BT nua; um 2º multiplex é recusado). Os parâmetros seguem a ordem dos vãos.
- **Parâmetros automáticos** (`define`): `{"linha_tem": "RL, RL_MT", "cabo": "^C120$", "param": "tipo=RL_CU120"}` — preenche `tipo` quando a linha do poste tem um RL e/ou o 1º cabo casa (o 1º item que casa vence; os mais específicos vêm antes). Não sobrescreve o que você digitou nem o que veio do vínculo do cabo. No balão do ativo aparece como **automático ‹linha: RL›** ou **‹cabo C35›**; sem RL na linha, o parâmetro fica no padrão da base (ou "não definido (a linha não tem RL)").
- Regras prontas do 081: `data/vinculacao_estruturas_081.json` (importar na Tabela de Estruturas). **RL_CAA40 = cabo S4/0**: confirmar com a planilha modular do religador.

### Ao abrir o programa

O programa **sempre abre no projeto default**: o projeto que estava aberto da última vez não volta sozinho. Abra o projeto em **Projeto → Abrir** (ou pelo plugin do nanoCAD). Enquanto o programa está aberto, recarregar ou trocar de tela mantém o projeto aberto.

### Parâmetros do ativo: o que está preenchido e o que falta

Passe o mouse no **ativo** (tabela Demais Ativos ou Totalizadora): o balão lista, por estrutura, os parâmetros **preenchidos** (✓, com a origem: *do pedido*, *vínculo*, *da obra* ou *padrão*) e os que **faltam** (⚠, com o motivo, ex.: "nenhum cabo vinculado a esta estrutura"). Valor que vem só do padrão da declaração conta como preenchido.
- **Cores (opcional):** em **Configurações › Aparência › ▸ Avançado**, o interruptor **Colorir ativos por parâmetros** pinta o ativo na tabela Demais Ativos: **verde** = todos os parâmetros preenchidos, **laranja** = falta algum (numa linha com várias estruturas, basta uma faltar). Ativo sem parâmetros não é pintado. Vem desligado; vale só na tela (não altera dados nem relatórios) e fica guardado neste computador.

### Vinculação por poste (obra vinda do plugin do nanoCAD)

Quando você abre uma obra que veio do plugin (tem a geometria do desenho), o programa **vincula sozinho pelo que chega em cada poste**: cada vão (cabo) tem 2 pontas e cada ponta termina em um poste; cada **estrutura do poste** procura, entre os cabos que chegam nele, os que cumprem as regras dela (Tabela de Estruturas: quantidade, cabos iguais, tipo de cabo, operação…) e liga em cada um, consumindo uma ponta. Ex.: uma **B1MT** (2 cabos MT iguais) liga nos dois cabos MT iguais que chegam no poste.
- Cabo **M** usa 1 ponta; cabo **I** usa 2. Estruturas que disputam o mesmo cabo: vale a **ordem da Tabela de Estruturas** (depois a ordem na linha do poste).
- Chegou **mais cabo que a estrutura aceita**: o programa escolhe os primeiros (ordem da Tabela de cabos) e **avisa**. Chegou **menos**: a estrutura fica incompleta e aparece no aviso/Validação.
- Vínculo **manual** (feito por você no modal) é mantido; os automáticos são refeitos a cada abertura. Cabo sem geometria e retirada (R) não são tocados.
- No modal **Vinculação › Script**, o botão **🏗️ Vincular por poste (desenho)** refaz na hora (Ctrl+Z desfaz) e lista o que sobrou, pontas soltas e estruturas incompletas.

### Parâmetros de obra (`OBRA.*`)

Parâmetro cujo nome começa com `OBRA.` é preenchido pelo **programa**, não pelo pedido: `OBRA.MUNICIPIO`, `OBRA.BAIRRO`, `OBRA.UTD`, `OBRA.UTEP` (do cabeçalho da obra, em **Projeto → Informações do relatório SAP**), `OBRA.REGIAO` (a regional, derivada) e `OBRA.CONTEXTO` (o contexto selecionado). Outros nomes `OBRA.*` são recusados ao importar a base.
- Na base, declare-os como qualquer parâmetro, de preferência por último: `FLBIM=FLBIM(poste=D, obra.regiao=CENTRO)` e use nos grupos: `FLBIM(obra.regiao=CENTRO|CENTRO_OESTE|OESTE)`. O valor da declaração é o **padrão** quando a obra não tem o dado (o resultado mostra o aviso *"OBRA.REGIAO não informado na obra; usando CENTRO"*).
- No pedido **não se informa**: `1-FLBIM(obra.regiao=SUL)` dá erro, e `1-FLBIT(121000R,50)` liga só poste e cabo (os `OBRA.*` são pulados na posição).
- Os valores são comparados sem acento e sem diferença de maiúsculas, com espaço virando `_` (`Centro Oeste` = `CENTRO_OESTE`, `34,5kV` = `34,5KV`).
- **`OBRA.REGIAO`**: vem da UTEP do cabeçalho (se for uma regional), senão da UTD, senão do município (+ bairro), pela tabela da ICP (`data/regionais_pernambuco.json`, só no projeto PERNAMBUCO).
- O **contexto** selecionado passa a ser gravado no `.obra` (e volta ao abrir; se o contexto não existir no projeto o programa avisa e fica em Nenhum; **Mesclar** mantém o atual).

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
2. `git push origin main --tags`. O GitHub Actions compila o `OrcaCore.exe` no Windows e publica a Release (com o arquivo `.sha256`).
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


## Relatórios ODI e ODD em PDF (TASK-140)

Na tela **Resultado do Orçamento**, na visão **SAP**, o botão **Relatórios PDF** abre um diálogo para preencher o cabeçalho (título, nº da nota, nº SAP, bairro, município, UTD, UTEP, projetista, data e ID de investimento) e salva dois arquivos, `NOTA_<nº>_ODI.pdf` e `NOTA_<nº>_ODD.pdf`, no mesmo formato da nota do SAP (tabela com título verde, Calibri 11, rodapé "Internal Use"). Entram os itens **visíveis** de cada tabela (filtros valem). O ID de investimento do ODD, se vazio, repete o do ODI. Projetista, UTD, UTEP, município e IDs ficam lembrados neste computador.

Na visão SAP, enquanto você não clicar em nenhum título de coluna, cada tabela (ODI e ODD) é exibida com **material (M) primeiro e mão de obra (N) depois**, e em cada grupo o **código em ordem crescente** (os numéricos como número, ex.: `500109` antes de `3481046`; os com letra em ordem natural). Os PDFs e o Excel seguem essa ordem. Clicar no título da coluna ordena de outro jeito; CÓDIGO também ordena como número.

## Executável: baixar e importar arquivos (TASK-142)

No programa instalado (janela própria, fora do navegador) os botões de baixar e de importar arquivo usam as caixas nativas do Windows (**Salvar como** e **Abrir**). Ao baixar, o programa avisa onde o arquivo foi salvo. Na versão online, no navegador, nada mudou. Ao trocar de projeto em **Configurações**, a barra de modais passa a mostrar os modais (e a ordem) do projeto escolhido. Ao fechar a **Tabela de orçamento**, o Resumo recarrega a base: um ativo editado e sincronizado com a nuvem já vale na hora.

## Base de Orçamento: pesquisa e ficha do ativo (TASK-146)

Na tela **Base de Orçamento**, a **barra de pesquisa fixa no topo** (atalho **Ctrl+K**, que leva o cursor para ela) procura ativos e códigos: digite parte do ativo, do apelido, do código ou da descrição, use ↑/↓ e **Enter** para abrir a **ficha do ativo**. A ficha mostra: **Itens** (comuns e por parâmetro, com os códigos clicáveis e, ao lado, onde cada código é usado), **Parâmetros e variantes** (padrões e máscaras usadas), **Chama / É chamado por** (links para as fichas dos outros ativos) e **Simulador** (digite `ATIVO(parâmetros)` e veja os códigos e quantidades, calculados com a base já salva). Na lateral, as **verificações** apontam: chamada a ativo que não existe, código sem descrição, código com descrições diferentes em ativos diferentes e ativo sem itens. **Ctrl+clique** numa célula ATIVO ou CÓDIGO da tabela abre a ficha/o uso do código; **Ir para a tabela** destaca a linha do ativo. A ficha mostra as edições da tela mesmo antes de salvar. A **Consulta Rápida** (canto da tela) usa a mesma pesquisa.

## Base de Orçamento: cadastro rápido de ativos (TASK-147)

**+ Novo ativo** abre o assistente: **1) Identidade e parâmetros** (nome, apelidos, descrição, parâmetros e padrões — ou escolha um **Modelo**), **2) Itens** numa grade como a do Excel e **3) Revisar**. Na grade, um **código que já existe** na base preenche sozinho a descrição, o MDO e o filtro (e sugere códigos ao digitar); **Enter** desce; dá para **colar do Excel** (CÓDIGO ⇥ FATOR I [⇥ FATOR R ⇥ LV], ou as 13 colunas da base) e **colar combinações** (valor ⇥ valor ⇥ código, como o CON). A coluna **Condição** usa a sintaxe dos grupos (`cabo=%S336%`); no **Código**, `:= 3*TERM(cabo)` faz o ativo chamar outros. Ao revisar, o programa mostra as linhas exatas, valida com as mesmas regras da importação da base (só aponta o que o ativo novo causaria), simula o ativo e permite **Copiar linhas (Excel)** ou **Adicionar à tabela** (no fim da tabela do projeto; depois use **Salvar Alterações** e, se for o caso, **Sincronizar**). O rascunho fica guardado até você usar **Limpar**.

## Base de Orçamento: publicar só o ativo (TASK-148, administradores)

No passo **3 · Revisar** do cadastro rápido (ou ao **Editar** um ativo pela ficha), o botão **Publicar no <projeto>** troca **só as linhas daquele ativo**, na base local e na nuvem, sem reenviar a tabela inteira. Antes, a revisão mostra o que muda (**novas**, **alteradas** com o valor antigo → novo, **removidas**), a validação (a mesma da importação; erro recusa a publicação), os **ativos que chamam este** e o simulador. Ao publicar, fica guardado **um passo anterior** daquele ativo: **Desfazer esta publicação** (no painel de sucesso ou na ficha: **↶ Desfazer última publicação**) volta ao que era antes; usar de novo refaz. Se outro administrador alterar o ativo depois que você o abriu, o programa **avisa** (quem e quando) e você escolhe **recarregar a versão atual** ou **publicar mesmo assim**. Se a nuvem falhar, nada é alterado. A tabela da tela é recarregada depois de publicar (edições ainda não salvas nela se perdem — o programa pergunta antes).

## Tabela Outros: escolher parâmetros pelo teclado (TASK-149)

Ao digitar `(` depois de um ativo, a caixa de parâmetros agora funciona pelo teclado: **↓/↑** realçam um valor do parâmetro da vez, **Enter ou Tab** inserem o valor, acrescentam a vírgula e já mostram as opções do **próximo** parâmetro; no **último**, o `)` é fechado sozinho. Sem valor realçado, Enter e Tab fazem o de sempre. Clicar num valor também avança. As máscaras simples cadastradas (`%S336%`) aparecem como valor (`S336`).

## Resumo: inverter a ordem das tabelas (TASK-150)

No grupo **Layout**, o botão **⇄ Inverter** põe **Demais ativos (postes) à esquerda** e **Cabos à direita** (nas empilhadas: Demais ativos em cima). Clique de novo para voltar. A escolha fica guardada neste navegador.

## OrçaCore: novo nome (TASK-156; antes GridR, TASK-152)

O programa passou a se chamar **GridR** e tem ícone próprio (grade com o caminho laranja ligando os nós). Quem já usa o programa **não precisa fazer nada**: a atualização automática continua funcionando, e o arquivo no seu computador pode seguir com o nome antigo (`Leitor_Projetos.exe`) — é o mesmo programa. (Na TASK-152 os downloads passaram a vir como `GridR.exe`; hoje vêm como `OrcaCore.exe`, ver abaixo.) As senhas salvas da janela e os demais dados são preservados.

## Projeto de trabalho (TASK-154, aplicativo desktop)

No aplicativo (não na versão online) a barra do Resumo tem o menu **📁 Projeto**:
- **Novo projeto…**: informe o nome, o número de obra e a pasta, e clique em **Criar**. O programa cria `<nome>.obra` na pasta e passa a gravar tudo nele sozinho (Aplicar e OK só ficam ativos depois de criar). O botão **Informações do relatório SAP…** guarda o cabeçalho dos relatórios ODI/ODD no projeto.
- **Abrir projeto…**: escolha um `.obra` (os `.obra.json` antigos também abrem). Se o Resumo já tiver dados, escolha **Abrir** (apaga a tela e usa o projeto), **Mesclar** (junta) ou **Sobrescrever** (a tela vale e vai por cima do projeto).
- **Propriedades do projeto…**: muda nome, número e cabeçalho SAP. **Voltar ao projeto default**: volta a trabalhar como antes, sem arquivo.
- O número de obra vem do projeto; ao exportar o REC aparece um pop-up para confirmar o número (vale só para aquela exportação). Relatórios e exportações abrem a caixa "Salvar como" na pasta do projeto, e você pode escolher outro destino.
- Importar/exportar JSON continua: dentro de um projeto só carrega ou exporta as tabelas do Resumo; título e demais informações só vêm ao abrir o `.obra`.
- A gravação é automática (2 s depois da última edição, ao ir para o Resultado/Tabela de orçamento/Totalizadora e ao voltar ao default). Se outro computador alterar o mesmo arquivo, o programa avisa antes de gravar.
- **📋 Colar SAP**: copie o JSON do Atlas e clique no botão na janela do projeto (preenche nome, número de obra e dados do SAP), na janela "Informações do relatório SAP" ou na janela dos relatórios ODI/ODD. Se o sistema não deixar ler a área de transferência, cole o texto na caixinha que abre. O ID de investimento vem do `odiCodigo` (e `oddCodigo` para o ODD). No relatório, se a nota colada for diferente do número do projeto, o programa avisa e usa a colada só naquele relatório.

O programa agora se chama **OrçaCore** (o ícone continua o mesmo). Quem já usa **não precisa fazer nada**: a atualização automática continua funcionando e o arquivo no seu computador pode seguir com o nome antigo (`GridR.exe` ou `Leitor_Projetos.exe`) — é o mesmo programa. Os novos downloads vêm como `OrcaCore.exe` (sem acento, por ser nome de arquivo). As senhas salvas da janela e os demais dados são preservados.

## Resultado do orçamento: Comparar por formato e Importar PDFs SAP (TASK-157)

- **⇄ Comparar** (e **Arquivo…** na barra de comparação) pergunta o formato da referência: **.REC** (como sempre), **.obra** (usa só o resultado do orçamento que ficou gravado no projeto; se o arquivo não tiver, o programa explica como gerar), **PDFs PB** (Orçamento + Lista de materiais, os mesmos do Importar PDFs) ou **PDF SAP** (escolha o ODI, o ODD ou os dois; vale o PDF gerado pelo programa e o feito por fora no mesmo formato).
- **Importar PDFs** tem o seletor **PB | SAP**. Em SAP, as linhas do ODI entram como instalação (I) e as do ODD como retirada (R), mescladas na tabela do Resultado. O programa descobre se é ODI ou ODD pelo "TIPO DE ORDEM" dentro do PDF, não pelo nome do arquivo. PDF escaneado como imagem não pode ser lido.

## Salvar, Salvar como e salvamento automático (TASK-158, aplicativo desktop)

Na barra do Resumo ficam o nome do projeto, o botão **💾 Salvar ▾** e uma etiqueta de estado (Salvo às 14:32 · Não salvo · Gravando… · Rascunho gravado…).
- **Salvar** (Ctrl+S) grava tudo no `.obra`: Resumo, resultado do orçamento (com as edições), dados do SAP, nome e número. **Salvar como…** (Ctrl+Shift+S, no ▾) grava uma cópia com outro nome ou pasta e passa a trabalhar nela.
- **Salvamento automático** (chave no ▾): **novo projeto** nasce ligado; **projeto aberto** vem desligado (você decide quando gravar; a etiqueta mostra "Não salvo" e o Salvar fica laranja). Ao ligar a chave, o que estava pendente é gravado na hora.
- **Projeto default**: o automático fica ligado e guarda um **rascunho** do Resumo na pasta `rascunhos` ao lado do programa. Para trazê-lo de volta: ▾ → **Recuperar rascunho do default…** (o mais recente ou o da sessão anterior). Salvar ou Salvar como no default abre direto "Novo projeto".
- Ao voltar ao default, abrir outro projeto ou criar um novo com alterações não salvas (automático desligado), o programa pergunta se quer salvar.
- O aviso vermelho "Arquivo alterado por fora" só aparece quando o programa vai gravar e o arquivo mudou no disco; clique nele para escolher Sobrescrever, Recarregar ou Decidir depois.

## Atualização automática: progresso e cancelar (TASK-159)

Ao clicar em **Instalar e Reiniciar** no aviso de atualização, aparece uma **barra de progresso** com a porcentagem e os MB baixados (ex.: "Baixando… 36% (54,0 MB de 150 MB)"). O botão **Cancelar** para o download a qualquer momento antes da troca: o arquivo parcial é apagado e o programa continua exatamente como estava. Se a conexão ficar parada por 20 segundos ou o download passar de 30 minutos, o programa avisa o erro. Se você recarregar a página no meio, a barra volta. Depois de conferir o arquivo (SHA-256), o programa fecha sozinho e abre na versão nova.

## Etiqueta de estado do projeto e PDF SAP com ODI e ODD juntos (TASK-160)

- A etiqueta ao lado do **Salvar** sempre diz o estado do salvamento automático: "Salvo às 14:32 · Automático ligado", "Não salvo · Automático desligado", "Rascunho gravado às … · Automático ligado" (projeto default). Se o arquivo do projeto mudar no disco, aparece à parte a etiqueta vermelha **⚠ Arquivo alterado por fora — resolver**; ao clicar, a janela mostra quando o arquivo foi modificado e deixa escolher Sobrescrever, Recarregar ou Decidir depois.
- Ao ler um PDF SAP (Importar PDFs e Comparar), um mesmo arquivo pode trazer o **ODI e o ODD** (como o relatório do ICP): o programa separa as duas partes pelo "TIPO DE ORDEM" de cada página (páginas sem cabeçalho continuam a parte anterior).

## Relatório SAP em PDF único (TASK-161)

O botão **Relatórios PDF** do Resultado agora gera **um PDF só**, `NOTA_<nº>_ODI_ODD.pdf`: o ODI primeiro e o ODD depois, cada um começando em página nova com o cabeçalho (se só um deles tiver itens, o arquivo leva só ele: `NOTA_<nº>_ODI.pdf` ou `_ODD.pdf`). Como no relatório do ICP, as páginas de continuação não repetem o cabeçalho nem a linha verde, a primeira página comporta 48 linhas e as quantidades saem sempre com duas casas. A ordem é a da visão SAP (material antes de mão de obra, por código).

## Base de Orçamento: barra de pesquisa fixa (TASK-162)

A pesquisa da Base de Orçamento agora é uma **barra fixa no topo da tela** (ela acompanha a rolagem). **Ctrl+K** leva o cursor para ela de qualquer lugar da tela. Os resultados aparecem logo abaixo da barra: ↑/↓ escolhem, **Enter** abre a ficha do ativo (ou o uso do código), **Esc** fecha a lista e o **✕** limpa o texto. O painel flutuante "Consulta Rápida" foi removido: a barra o substitui.

## Resultado do orçamento: Montar sempre recalcula (TASK-180)

**Montar Orçamento** sempre calcula de novo com a base técnica atual, mesmo com projeto aberto e resultado gravado nele. Se você tinha editado o resultado à mão, o log informa que as edições foram substituídas. O resultado gravado no arquivo `.obra` (com as edições) só volta quando você **abre o projeto** com a tela do Resultado à vista: a tabela mostra o que está no arquivo, sem recalcular. Se o projeto não tem resultado gravado, a tabela fica vazia com a orientação de voltar ao Resumo e clicar em Montar Orçamento. Recarregar a página do Resultado também recalcula.

## Dicionário de cabos (TASK-173)

No modal **Vinculação** há a aba **Dicionário de cabos**: cada linha diz o que um cabo **é** (família, valor de cabo, seção em mm² para o "maior primeiro", nível MT/BT e fases quando o texto não traz, apelidos e como montar o neutro). A primeira linha que casa vale. **Sem linhas, o programa usa as regras de Pernambuco** (C, A…P, S/W, PB, RAMA, CAZ). Use **Testar um texto** (ex.: `3xA35P(7,94)-13,8kV 45 m`) para ver código, fases, neutro, nível, seção, valor e a linha usada. Marque **Código solto** quando o código sozinho (ex.: `ALU-50 30 m`) já deve valer como notação do LABEL. A Tabela de Cabos continua dizendo em quais estruturas o cabo entra.

## Convenção de postes (TASK-174)

No modal **Vinculação**, a aba **Postes** diz o que é poste na linha de Outros: os **tipos** (código e letra da notação compacta), as **formas de escrever** o poste (a primeira é o nome oficial; use `{tipo}`, `{sufixo}`, `{altura}`, `{esforco}`), a notação compacta e as faixas de altura e de dígitos do esforço. **Sem tabela, vale a de Pernambuco** (`DT11/600`, `11600D`). Exemplo de outra concessionária: formato `{tipo}{altura}-{esforco}` com o tipo `P` reconhece `P11-400, N3` e passa o parâmetro `poste` para a N3. Use **Testar um nome** para conferir. Se altura e esforço estiverem colados (`11600D`), o poste que existe na base decide a separação; sem base, valem as faixas.

## Leitura do desenho: layers e cores de rede (TASK-175)

No modal **Vinculação**, a aba **Desenho** diz quais **layers** e **cores** do DXF são rede **MT** e **BT** (os mesmos campos do plugin do nanoCAD). Layer: o nome **contém** o trecho; `=NOME` exige o nome exato; `MT_*` é curinga. Cor: número ACI (1 = vermelho), `#RRGGBB` ou faixa RGB. **Regra:** `COR_MANDA` (padrão: a layer confirma a cor e divergência descarta a linha), `LAYER_MANDA` (a layer decide; sem layer de rede, vale a cor) ou `SO_LAYER` (a cor é ignorada). **Sem configuração, vale Pernambuco** (REDE_MT/REDE_BT, vermelho = MT, verde = BT). Use **Testar** (layer + cor) e **Salvar leitura do desenho**.

## Perfil de exportação (TASK-176)

No modal **Vinculação**, a aba **Exportação** diz como saem os relatórios do projeto: os **campos do cabeçalho** (rótulo no PDF, coluna do cabeçalho, se viram parâmetro `OBRA.<CHAVE>` da base e se o programa lembra o valor), o **centro** de material e de mão de obra (padrão 1600/1000), a cor do título e o rodapé do PDF, as **colunas do Excel** e quais **arquivos** o projeto oferece. O `.obra` grava as chaves do cabeçalho. Exemplo: acrescente o campo `contrato` marcado `OBRA.*` e a base pode usar `OBRA.CONTRATO` como parâmetro. **Pré-visualizar uma página** mostra o PDF de exemplo com o perfil em edição. **Sem perfil, vale o de Neoenergia.**

## Pacote de projeto (TASK-177)

No modal **Vinculação**, a aba **Pacote** exporta e importa tudo de um projeto num arquivo `.zip`. **Exportar:** marque as partes (base técnica, vinculação, dicionário de cabos, postes, desenho, exportação, regras do leitor, tabela regional) e baixe o pacote. **Importar** (só administrador): escolha o `.zip`, veja a **prévia** — para cada parte, se é *novo*, *substitui* ou *igual*, quantos itens há hoje e no pacote e os avisos do validador — marque o que quer importar e confirme. Nada é gravado antes da confirmação; se uma parte tiver erro, nada é gravado. A base de outro projeto vira a do projeto onde você importa. Pacote de uma versão de formato mais nova que a do programa é recusado. Um projeto novo, sem configuração, começa pelo **pacote de partida** dele (campo do cadastro do projeto, ou o pacote que declara o código/nome dele): hoje existem o de Pernambuco e o de Rondônia, em `data/pacotes/`.

## Casos de referência (TASK-178)

No **Resultado do Orçamento**, **Guardar como caso de referência** guarda o pedido e o resultado desta tela. No modal **Vinculação → Casos** você vê os casos do projeto e usa **Conferir base**: o programa roda todos e diz quais ficaram **iguais**, quais **mudaram** (código a código, antes → depois) e quais deram **erro**. Para testar uma base antes de importar, escolha o CSV em *conferir um CSV antes de importar* — nada é gravado. Se a mudança era esperada, clique em **Mudança esperada** e diga o motivo: o resultado novo passa a ser o esperado e a nota fica no histórico. Ao importar um pacote com base, a prévia também mostra quantos casos mudariam. Os casos fazem parte do pacote do projeto.

## Guia e modelo para criar uma base (TASK-179)

O **guia da base** está em `docs/base_guia.md` (sintaxe do simples ao avançado, armadilhas e a lista de códigos do validador). Para começar de algo pronto, abra **Vinculação → Pacote**, escolha **Exemplo** em *Começar de um pacote pronto*, veja a prévia e importe num projeto novo: vêm uma base pequena (poste, N1, N3 com parâmetro, N4 com linha viva, cabos, alça, conector e um composto), a vinculação e 5 casos de referência para conferir. Na **ficha da base**, a coluna **O que faz** traz a frase em português de cada linha. Os erros e avisos da importação da base trazem um código (`B-005`…) que aponta para a seção do guia; o aviso **B-013** pega o equipamento (trafo, chave, religador…) colocado nas linhas gerais de um ativo que tem parâmetro desse tipo.

## Novo projeto: a nota é o nome (17/11/2026)

Ao criar um projeto (**Projeto → Novo projeto…** ou **Salvar como…**) e usar **📋 Colar SAP**, o **nome do projeto** (e o nome do arquivo `.obra`) e o **número da obra** passam a ser a **nota** do SAP (ex.: `9101578720.obra`). O **título** do SAP não vira o nome: fica guardado no projeto e é sugerido como **título dos relatórios PDF** (sem título guardado, o relatório usa o nome do projeto). Em projeto já criado, o Colar SAP continua como antes.

## Login: "Lembrar de mim" (17/11/2026)

Na tela de login, com **Lembrar de mim** marcado o programa **abre direto, sem pedir o login**, por até **30 dias** (ou até você clicar em **Sair**). Com a caixa **desmarcada**, o login vale só enquanto o programa está aberto: ao fechar e abrir de novo, a tela de login volta. Se a sessão vencer (30 dias, ou 24 h sem a marca), o programa avisa "Sessão expirada" e leva ao login. Quem já estava logado antes desta versão continua entrando direto até o token vencer.

## Ativo fora da base, linhas congeladas e bloqueadas (17/11/2026)

- **Ativo que não está na base:** ligue **Configurações › Aparência › Avançado › Colorir ativos por parâmetros**. Em **Cabos e Outros**, o ativo que não existe na base (nem como ativo, apelido, componente ou código do projeto) fica **vermelho**; passe o mouse para ver a dica. Em Outros, verde = parâmetros completos e laranja = falta algum, como antes. A cor confere o texto digitado; o botão **Validar** confere de novo depois das regras de conversão.
- **Congelar linha** (botão direito na linha › *Congelar linha*; vale para as linhas selecionadas): a linha fica na tela, **editável**, mas **fora de tudo** — Totalizadora, orçamento, validação, ajustes, cores, resumo da rede e vínculos. Serve para anotar (por exemplo, um limite de seção). Um **cabo congelado** pode continuar sendo a **fonte de um vínculo que você fez à mão**: a estrutura ligada a ele segue funcionando normalmente. Os vínculos automáticos (por poste, mapa e script) não ligam cabo congelado, e uma estrutura congelada não entra em vínculo.
- **Bloquear linha**: a linha **não pode ser editada nem excluída** (colar e limpar em bloco também a pulam), mas continua nos cálculos. Uma linha pode ser congelada e bloqueada ao mesmo tempo. **Desbloquear** é a única ação aceita numa linha bloqueada. Ctrl+Z desfaz.
- As marcas ❄ (congelada) e 🔒 (bloqueada) aparecem ao lado do número da linha e ficam gravadas no projeto `.obra`.

## Base de Orçamento mais rápida (17/11/2026)

- Salvou, publicou ou sincronizou na **Base de Orçamento**? As outras telas (Resumo/Orçamento) são avisadas e **recarregam a base na hora**: o ativo novo já aparece nas sugestões e nas cores, sem esperar. Aparece o aviso "Base de orçamento atualizada".
- A base da nuvem não é mais baixada inteira a cada abertura: o servidor só atualiza a cópia local a cada **30 segundos**.
- A **pesquisa** (Ctrl+K) ficou mais ágil: usa a base já carregada e se atualiza sozinha em segundo plano.
- Uma **barra de status** no canto inferior esquerdo mostra o que está acontecendo (Carregando, Validando, Salvando, Publicando, Recarregando) e o resultado (✔ Pronto ou ✖ erro, que fica até você fechar). Os botões de salvar/sincronizar ficam desabilitados enquanto a operação roda.
- A tabela grande é montada em lotes (a barra mostra "Montando linhas… X de Y").

## Renomear um ativo ao editar (17/11/2026)

- Na ficha do ativo, **Editar** abre o cadastro rápido. No passo 1 o campo **Nome do ativo** agora pode ser trocado: aparece "Renomeando “A” → “B”".
- Por padrão o **nome antigo vira apelido** do ativo novo (marcado como recomendado): as chamadas `:=` de outros ativos e os orçamentos que ainda usam o nome antigo continuam funcionando. Desmarque só se quiser que o nome antigo deixe de existir (a tela lista quem o chama).
- Não dá para renomear para o nome de **outro** ativo do mesmo projeto (aparece o erro na revisão). O ativo é renomeado só no projeto escolhido.
- Ao publicar, as linhas do nome antigo saem e as do novo entram de uma vez (base local e nuvem). **Desfazer** volta ao nome antigo; desfazer de novo refaz.

## Comentário, operação antes do ativo e ativo fora da base (17/11/2026)

- **Comentário:** botão direito na linha › *Marcar como comentário* (💬). Serve para anotar na tabela (por exemplo, o limite de uma seção): fica fora de todos os cálculos, como a linha congelada, e continua editável. *Tirar comentário* devolve a linha ao normal.
- **Operação antes do ativo:** nas linhas em branco de Cabos e Outros, escolha a **Operação** na coluna antes de digitar o ativo; a nova linha nasce com ela. Se não escolher, vale a operação da linha de cima, como antes.
- **Ativo que não existe na base:** com *Colorir ativos por parâmetros* ligado, passe o mouse no ativo vermelho: o balão mostra "**ATIVO NÃO EXISTE NA BASE**". Um nome que é **componente** ou **código** da base não é marcado.

## Ocultar parâmetros ao sair da linha (18/11/2026)

- **Configurações › Aparência › Avançado › Ocultar parâmetros ao sair da linha** (desligada por padrão). Ligada, na tabela **Demais Ativos** o ativo que tem parâmetros mostra só o nome quando você sai da linha: `RL, PR(D,X), ATRINT` aparece como `RL, PR, ATRINT`. Ao entrar na linha o texto completo volta para você editar.
- **Passe o mouse** no ativo: o balão lista os parâmetros e de onde veio cada valor. O que você preencheu à mão aparece como **manual** (os outros mostram vínculo, da obra, padrão ou automático).
- É só aparência: o texto da linha, o cálculo, a cópia e o projeto continuam com os parâmetros.

## Linhas congeladas, bloqueadas e comentários no projeto; "#" = comentário (18/11/2026)

- **Congelar, bloquear e comentar** ficam gravados no projeto: ao reabrir, as linhas voltam com ❄, 🔒 e 💬. **Ctrl+Z** também desfaz e refaz essas marcas.
- **"#" no começo do ativo** (em Cabos ou Demais Ativos) transforma a linha em **comentário** ao sair do campo: `# conferir poste`. Apagar o `#` devolve a linha ao normal; *Tirar comentário* pelo menu também tira o `#`. Linha bloqueada não muda.
- **Botões ❄ 💬 🔒:** no cabeçalho de **Cabos** e **Demais Ativos**; agem nas linhas selecionadas (clique no número da linha; Ctrl+clique seleciona várias) ou, sem seleção, na linha onde está o cursor. O menu do botão direito continua igual. Clicar de novo desfaz.

## Operação digitada e seletor de parâmetros (18/11/2026)

- **Operação:** digite (`*r` vira `*R`) ou escolha. Ao entrar no campo aparece a lista com **todas** as operações (instalar, remover, manter e as versões em linha viva); ↑↓ e Enter escolhem, clique também, digitar filtra, Esc cancela. Só vale `I`, `*I`, `R`, `*R`, `M` ou `*M`; outra coisa deixa a célula vermelha ("Operação inválida…") e o valor anterior volta. Vale em Cabos, Demais Ativos, nas linhas em branco (escolha antes do ativo) e na Totalizadora. **Enter** na operação leva ao ativo da mesma linha; **←** no começo do ativo volta à operação.
- **Parâmetros do ativo (Demais Ativos):** ao digitar `(` depois de um ativo com parâmetros, o programa já escreve o `)` e abre a caixa, no estilo do VS Code: no topo a assinatura `ATIVO(param1, param2…)` com o parâmetro da vez em destaque e, embaixo, a lista de valores terminando na linha **✎ digitar outro valor…**. **↑↓** navegam (↑ vai direto à linha de digitar); **Enter** ou **Tab** aceitam o valor realçado e **avançam para o próximo parâmetro** (a caixa continua aberta; no último ela fecha com o cursor depois do `)`); **Shift+Tab** volta; **Esc** fecha mantendo o texto (Ctrl+Z desfaz). Digitar filtra a lista; você também pode clicar no valor.
- O valor entra **pelo nome**: `PR(CABO2=S04)`. Dá para mudar só um parâmetro sem passar pelos anteriores. Se o texto já tinha valores por posição (`CE_4(12600D,A35P,A35P)`), ao mudar um deles o pedido é reescrito por nome. Parênteses vazios somem (`PR()` vira `PR`). A caixa mostra ao lado de cada parâmetro o valor que vale e de onde vem (**manual**, vínculo, da obra, padrão ou automático).
- O preenchimento ficou rápido também em bases grandes: a base inteira deixou de ser reprocessada a cada tecla.

## Menu pelo teclado, Duplicar linha e Enter (18/11/2026)

- **Botão direito na linha:** com o menu aberto, **↑↓** navegam pelas opções, **Enter** executa e **Esc** fecha.
- **Duplicar linha:** duplica a linha (ou as selecionadas) e coloca a cópia **logo abaixo**; a cópia sai sem bloqueio e sem o vínculo da original. Ctrl+Z desfaz.
- **Enter no ativo:** cria uma linha nova logo abaixo da atual e vai para ela; numa linha em branco o Enter só desce.
- **Sugestões ao digitar:** só aparecem **ativos** (da base) e **parâmetros**; as linhas já digitadas na tabela não são mais oferecidas.
- **Onde ficam as opções de tela:** botão **⚙ Configurações** (barra de cima) › **Aparência** › **Avançado**: *Colorir ativos por parâmetros* e *Ocultar parâmetros ao sair da linha*.

## Tabelas e vinculação rápidas com muitas linhas; aba «Cabos ⇄ Postes» (18/11/2026)

- **Tabelas grandes:** com mais de 60 linhas, o Resumo desenha primeiro só o que aparece na tela e completa o resto em segundo plano. Criar, excluir, duplicar e desfazer ficam em ~50 ms mesmo com 400+ linhas em cada tabela. Se você clicar ou apertar uma tecla antes de terminar, o resto é montado na hora (uma vez).
- **Modal «VINCULAR CABOS» › Vínculo:** compacto — **um cabo por linha** e **um poste por linha** (as estruturas do poste ficam lado a lado como `N1 1/2`). Passe o mouse na estrutura para ver o motivo/regra. A rolagem e os cliques continuam fluidos com centenas de linhas.
- **Aba «Cabos ⇄ Postes»:** só mostra cabos e postes. Clique no cabo e no poste a que ele chega (ou o contrário; um cabo tem até 2 pontas); clique na linha para desfazer; botão direito no cabo: «Tirar este cabo dos postes». O **algoritmo por poste** escolhe sozinho, em cada poste, a estrutura que recebe o cabo (mesmas regras da Tabela de Estruturas). Com «Vincular as estruturas automaticamente» ligado isso acontece a cada ligação; senão use «Vincular estruturas agora». Ctrl+Z (ou «Desfazer») desfaz. Vínculos feitos à mão na aba Vínculo continuam mandando; cabo congelado não é alterado.

## Aba «Cabos ⇄ Postes» pelo teclado, avisos do algoritmo e curinga nas layers (18/11/2026)

- **Teclado na aba Cabos ⇄ Postes:** **↑↓** escolhem o cabo; **→** passa para os postes (**↑↓** escolhem, **Enter** liga ou desliga o cabo àquele poste); **←** volta aos cabos.
- **Avisos do algoritmo:** sob a lista aparecem as estruturas que ficaram incompletas, os cabos que sobraram e as pontas soltas — é onde você sabe o que ainda precisa ligar.
- **Layers com curinga:** nas regras do Leitor, o campo *Layers* aceita `*` (por exemplo `POSTE_*_LV` pega todas as variantes `_LV`; `*_LV` pega qualquer layer terminada em `_LV`). Sem `*`, o nome continua exato (sem diferenciar maiúsculas).
- **Tabelas com 1000+ linhas** continuam fluidas: só a parte perto da tela é montada; o resto entra ao rolar.

## Ocultar parâmetros: clicar na linha não mostra os parâmetros; "(" abre o balão (18/11/2026)

Com *Ocultar parâmetros ao sair da linha* ligado, clicar na linha mantém só o nome do ativo (`PR`). Ao **digitar `(`** (ou levar o cursor para dentro do parêntese com **End**/**→**), os parâmetros **preenchidos** aparecem no **balão** junto ao campo — não na linha. Passar o mouse mostra o mesmo balão.

## Ocultar parâmetros, balão dos ativos e Tab em Outros (18/11/2026)

- **Opção desligada:** o campo do ativo mostra o texto inteiro (com os parênteses) e se edita normalmente.
- **Opção ligada (⚙ Configurações › Aparência › Avançado):** o campo mostra só os nomes (`12600D, N1, N4`) e você escreve, seleciona, cola e copia normalmente. Os parâmetros ficam **presos a cada ativo** enquanto ele existir; se você apagar ou renomear um ativo, ele passa a ser um ativo novo (sem parâmetros). Digitar `(` não escreve nada: abre os parâmetros do ativo onde está o cursor.
- **Balão dos ativos:** **Ctrl+Enter** no ativo, **duplo clique** ou **mouse parado** ~1,5 s na linha. Mostra os ativos da linha e o que já está preenchido/falta. **↑↓** ou **Tab** navegam; **Enter** (ou clique) abre os parâmetros do ativo. Nos parâmetros: **Tab** anda entre eles, **↑↓** escolhem o valor, **Enter** grava e vai ao próximo (no último, ao próximo ativo com pendência), **Del** limpa o valor, **Esc** volta à lista; **Esc** de novo fecha e devolve o cursor ao ativo.
- **Tab na tabela Outros:** Operação → Ativo → Operação da linha de baixo (na última linha cria uma nova); **Shift+Tab** volta. **Enter** no ativo cria linha nova abaixo, como antes.

## Ajustes do balão, sugestões e menu de contexto (18/11/2026)

- **Balão pelo mouse:** deixe o mouse parado ~1,5 s sobre a linha que **já está selecionada**; ele abre com o teclado ativo (↑↓/Tab navegam, Enter edita, Esc fecha). Em outra linha não abre.
- **Sugestão de ativos:** digitou o nome completo (ex.: `PR`), ele aparece na lista. **Tab** com a sugestão igual ao texto segue para a próxima célula.
- **Seta para baixo** na última linha preenchida vai para a primeira linha em branco.
- **Menu de contexto** (botão direito ou tecla de menu): depois de escolher uma ação (congelar, bloquear…) ou apertar Esc, o cursor volta para a célula em que estava.

## Leitor: comprimento do cabo pela posição no desenho (19/11/2026)

- **Por que mudou:** a lista do DXF é ordenada por altura; a regra «próximos N itens» não acha (ou acha o de outro cabo) o comprimento que está logo abaixo do rótulo. Agora a vizinhança pode usar a **posição**.
- **Como usar (regras do Leitor › Processamento › regra do cabo com label sozinho):** em *Vizinhança*, marque **Por posição no desenho (DXF)**, escolha o **Raio** (unidades do desenho; 12 cobre o padrão da CTW — o comprimento costuma ficar a 2–7 unidades) e **Onde** (abaixo, acima ou qualquer lado). A regex da vizinhança continua sendo a do comprimento (`^(\d+(?:[.,]\d+)?)\s*m$`).
- Rótulos riscados/duplos e o de cima, que dividem o mesmo comprimento, passam a ler o mesmo valor. Com coordenadas e nada dentro do raio, o cabo fica sem comprimento (aumente o raio se for o caso). Em PDF (sem coordenadas) vale a ordem da lista, como antes.
- Rótulos com neutro (`3xC16(16)+16`, `3xC35(35)+35`) também são reconhecidos.

## Cabo com condutor extra: `3xC16(16)+16` (19/11/2026)

No label do cabo, `+N` (sem kV) depois do neutro é **mais um condutor da mesma família** com a seção N. `3xC16(16)+16 35 m` vira, na Totalizadora, **5 linhas** `C16` de 35 m (3 fases + neutro + o `+16`). `+20kV` continua sendo a tensão da rede. A coluna «Qtd ativos» da tabela Cabos conta só as fases (3).

## Assistente de IA: ferramentas, plano e skills (19/11/2026)

O assistente (painel ✨ Assistente IA) agora **consulta e calcula de verdade**: busca ativos na Base de Orçamento do projeto, lê as suas tabelas, confere linhas, totaliza e calcula o orçamento. Exemplos: «monte uma rede com 3 postes DT11/300 e um trafo de 25», «tire o PR220 do poste 2», «o que não foi encontrado na base?», «calcule o orçamento disso», «e se eu adicionar mais 2 CFU?».

- **Nada muda sem você aceitar.** Quando o assistente quer mudar a tabela, aparece um **cartão do plano** no chat (＋ adicionar, ✎ editar, − remover, com avisos como «ativo não encontrado na base»). **Aplicar** faz tudo de uma vez (Ctrl+Z ou **Desfazer** volta); **Descartar** ignora. Se você mexer numa linha depois do pedido, a alteração daquela linha é ignorada. Obras salvas também passam pelo cartão e pela janela de confirmação de sempre.
- **Atalho rápido:** «adicionar 2 CON» / «remover 1 TR125» continua sem IA e aplica na hora.
- **Skills (📚 no topo do painel):** guias em texto que o assistente lê antes de agir. Vêm `formato_ativos` (como escrever Outros e Cabos), `montar_projeto` (regras técnicas e passo a passo), `orcamento` (como calcular e conferir) e `projeto` do Pernambuco. O administrador pode **editar, criar e apagar** skills (as suas valem mais que as que vêm com o programa) e usar **Gerar catálogo de ativos**, que lê a Base de Orçamento do projeto selecionado e grava a skill `ativos` (famílias, códigos e descrições). Gere de novo quando a base mudar.
- **Como escrever uma skill:** a 1ª linha pode ser `descricao: quando usar` (é o que o assistente vê antes de abrir); o resto é texto livre, com regras curtas. Nome em minúsculas, números, `_` e `-`. «Deste projeto» vale só para o projeto selecionado; «Todos os projetos» vale para todos. No .exe elas ficam na pasta `skills/` ao lado do programa.
- **Regras aprendidas:** continuam valendo («lembre-se…», «regra:»). Skills servem para procedimentos e catálogos maiores.
- Precisa de chave de IA (Gemini, OpenAI ou Claude) como antes. Sem chave, as skills continuam abrindo e editando.

## Provedor reserva do assistente (19/11/2026)

Para o assistente ficar sempre disponível, configure um **segundo provedor de IA** que entra sozinho quando o principal falha (sobrecarga/503, cota esgotada, chave inválida, sem rede).

- **Onde:** botão **Chave API** → **Configuração de IA** → seção **Provedor reserva (automático)**. Marque **Usar provedor reserva**, escolha um atalho (**DeepSeek**, **Qwen**, **OpenRouter**, **Ollama**, Gemini ou Claude) ou preencha **URL base**, **nome do modelo** e **chave**, e clique **Salvar reserva**. **Testar** salva e faz uma chamada mínima, dizendo se funcionou e em quanto tempo. Só o administrador configura.
- **DeepSeek, Qwen, OpenRouter e Ollama** usam o formato «OpenAI-compatível»: o que muda é a URL e o nome do modelo (os nomes mudam com o tempo: confira na documentação do serviço). Ollama roda no seu PC e não usa chave (digite qualquer texto). Modelos pequenos erram mais nas ferramentas do assistente.
- **Na prática:** se o principal falhar, o chat mostra «Usando o provedor reserva…» e a pergunta continua no reserva. O principal fica em pausa por cerca de 90 segundos (você não espera a falha a cada mensagem) e depois volta a ser tentado. Se os dois falharem, a mensagem mostra o motivo de cada um.
- A chave do reserva fica guardada no programa e nunca é mostrada de volta. Deixe o campo em branco para manter a chave salva.
- Sem chave no provedor principal mas com o reserva ativo, o assistente usa o reserva.

## Assistente: rápido, com andamento e botão Parar (19/11/2026)

- **Atalho instantâneo (sem IA):** comandos só de quantidade e ativo aplicam na hora, mesmo sem chave: `adicionar 2 CON`, `adicione 3 PR220, 2 CFU e 1 TR125`, `remova 1,5 IT; 2x CON`, `tire 2-CON`. Frases de verdade («adicionar 2 postes de 11 metros») vão para a IA.
- **Enquanto a IA trabalha** o chat mostra o que está fazendo e há quantos segundos («Consultando a base · 4 s»). O botão de enviar vira **■** (Parar): clique para interromper. Se uma chamada passar de 45 segundos (20 s com provedor reserva), o assistente desiste dela e troca de modelo/provedor; a pergunta inteira tem teto de 2 minutos.
- **Para ficar mais rápido:** use o modelo Flash Lite do Gemini (padrão) ou um provedor de resposta rápida; peça tudo numa frase só; deixe as skills curtas (a skill `formato_ativos` e as linhas da tela entram em TODA pergunta, então skills enormes deixam o assistente lento).

## Resumo da rede: como conta (19/11/2026)

O quadro **Postes · Média · Baixa · Equip.** (e o modal «Resumo da rede») soma só as linhas **instalando**: operação `I` ou `*I` (linha viva). `R`, `*R`, `M`, `*M`, linhas vazias e congeladas não entram.

- **Postes:** itens que são poste. Vale o nome `DT…`/`RC…`/`CV…`/`FV…` e também qualquer poste que a **convenção de postes do projeto** reconheça (modal Vinculação → Postes): em Pernambuco, `12600D` = DT12/600, `121000R` = RC12/1000, mesmo que a base guarde o nome compacto como oficial. Quantidade vale (`2 9600R` = 2 postes); parâmetros são ignorados (`12600D(fases=TRI)`).
- **Equipamentos:** itens que começam por `TR`, `CFU` ou `CFA` (ex.: o `TR` de um poste de trafo).
- **Média / Baixa:** cabo com tensão no rótulo (`3xA35P(7,94)-13,8kV 25 m`) = média (o vão conta uma vez e o **neutro não entra**); sem tensão (`3xC16(16) 20 m`) = baixa (cabo + neutro). No formato clássico: `CAA`/`CAL`/`P…` = média, `M2X`/`M3X` = baixa. Cabo sem comprimento usa o da linha de baixo.
- Se um projeto novo tiver outro tipo de poste, cadastre-o na **convenção de postes** do projeto; o resumo passa a contá-lo sozinho.

## Login lembrado e abertura de arquivos no programa (19/11/2026)

- **Lembrar de mim:** com a caixa marcada no login, o programa abre direto na tela inicial por 30 dias (sem a tela de login). Se o token vencer, ou se você sair (**Sair**), o login volta. Sem marcar, o login aparece a cada vez que abre o programa.
- **Escolher arquivo (Ler projeto, importar…):** assim que a caixa «Abrir» do Windows fecha, aparece «Abrindo <arquivo> (N MB)…» no pé da tela; o arquivo chega à tela logo em seguida (os bytes não passam mais pela ponte lenta da janela).

## Grupos de linhas e mover linhas (19/11/2026)

Nas tabelas **Demais Ativos** e **Cabos** você pode agrupar linhas e recolhê-las.

- **Criar:** digite uma linha começando com `@` (ex.: `@P1 Rua das Flores`). Ela vira o título do grupo (fica congelada: não entra em nenhum cálculo). `#` continua sendo comentário.
- **Enter / Shift+Enter** na linha do grupo (ou numa linha dele): **Enter** cria uma linha nova já dentro do grupo (no fim); **Shift+Enter** cria uma linha **solta** logo abaixo do grupo. Nas demais linhas nada mudou.
- **Botão direito → Grupo ▸:** *Adicionar ao grupo* (mostra a lista de grupos), *Criar grupo…* (pede o nome e leva as linhas selecionadas para ele), *Tirar do grupo*, *Recolher/Expandir*, *Desfazer o grupo* (as linhas ficam na tabela), *Mover para cima/baixo*. Funciona com uma ou várias linhas selecionadas.
- **+ / −** no número da linha do grupo recolhe ou expande (Alt+← / Alt+→ também); o número ao lado é a quantidade de linhas. O estado é salvo com a obra.
- **Mover:** **Alt+↑ / Alt+↓** move a linha (ou as selecionadas); na linha do grupo move o grupo todo. Passar da borda tira a linha do grupo; uma linha solta que desce sobre o título entra nele. Ctrl+Z desfaz.
- O menu do botão direito agora abre para cima (ou para a esquerda) quando não há espaço embaixo, em vez de ficar cortado.
- **Congelar / bloquear o grupo:** com a linha do grupo (`@…`) selecionada ou com o cursor nela, **Congelar grupo** e **Bloquear grupo** (botão direito ou botões ❄ 🔒) valem para todas as linhas dele; repetir desfaz.

## Pernambuco: estrutura solta (20/11/2026)

No projeto Pernambuco, uma linha de Outros que seja só estrutura (`N3`, `2 N3`, `N1, N4`, `B2`, `CE1`, `ISOL`…) classifica como **ESTRUTURA**. Linha com poste (`11300D, N3`) continua **POSTE**, cabo continua **CABO**, e materiais (`CON`, `PR220`, `FLABIT`) seguem sem classificação (`0`). Projetos cujas regras do leitor foram editadas à mão não mudam sozinhos: acrescente a regra em Regras do leitor › Classificação (ou importe o pacote).

## Mapa: escala 1:500 e mapa mais leve (22/11/2026)

No mapa de vinculação (Vinculação → Mapa), o seletor no canto inferior direito troca entre **Ajustar à tela** (cabe a rede inteira, como sempre) e **1:500** (o desenho em tamanho de prancha: 1 m = 2 mm; é aproximado, depende do monitor). No canto inferior esquerdo aparece **Escala ≈ 1:N** com uma barra em metros, que acompanha o zoom. Em 1:500 a rede pode ser maior que a tela: arraste o fundo para andar e use a roda para o zoom; **Enquadrar** volta a centralizar. A escolha fica lembrada.

Arrastar postes e curvar cabos em 1:500 guarda o ajuste em metros, separado do ajuste feito em «Ajustar»; trocar de escala não bagunça nenhum dos dois. Em redes grandes o mapa mostra só o que está perto da janela e some com os rótulos durante o movimento, para o arrastar e o zoom ficarem leves.

## Mapa: pontos como grupos, nome, comentário e Rearranjar (22/11/2026)

- **Tudo é ponto:** poste, estrutura avulsa e poste fictício. Cada ponto é um **grupo** em Demais Ativos com uma linha `@nome` e as linhas dos ativos. Se você preencher o comentário, a tabela também recebe uma linha `# comentário`. Para removê-la, limpe o campo e salve no painel do ponto. Deixe o nome em branco para usar o **número automático** (`P13`, `P14`…) ou digite um nome único.
- Abra o ponto (clique nele): a caixa **Ponto** mostra o nome, o comentário, os botões **Congelar grupo** e **Bloquear grupo** (valem para todas as linhas do ponto) e as linhas que o ponto tem em Demais Ativos e em Cabos.
- **Agrupar pontos (N):** transforma de uma vez os pontos que ainda não são grupo (os do CAD, por exemplo) em `@P<n>`. Nada na tabela muda sem você clicar.
- **As posições não mudam** quando você adiciona pontos: o novo ponto aparece num lugar livre perto do centro da tela.
- **Rearranjar:** organiza os pontos de forma uniforme e limpa (vãos do mesmo tamanho, alinhados, sem sobreposição; vãos retos). Grupos bloqueados ficam onde estão. Em 1:500 o botão devolve os pontos à posição do CAD. Ctrl+Z desfaz.

## Mapa: revisar vínculos e editar o poste

- Abra o ponto para ver os cabos à esquerda e as estruturas à direita. As linhas mostram os vínculos neste ponto. Clique no código de uma estrutura para destacar os cabos dela.
- Use **Editar poste** para abrir o cadastro preenchido com anotação, operação, nome e comentário. O formulário altera o mesmo poste e preserva a posição e os cabos. Se a edição alterar uma estrutura interna já vinculada, revise esse vínculo antes.
- Ao confirmar a chegada de um cabo, adicionar uma estrutura ou editar uma linha no mapa, o sistema aplica os vínculos inequívocos pelas regras do projeto. Use **Só pendências** para ver as pontas que precisam de decisão. **Desfazer** reverte a ação e os vínculos gerados por ela.

## Mapa: posicionar o ponto novo e copiar (22/11/2026)

- Ao confirmar um **ponto novo** (poste, estrutura ou poste fictício), o mapa pede que você o **posicione**: o ponto acompanha o mouse e o **clique** fixa o lugar. **Esc** mantém o lugar sugerido.
- **Copiar ponto:** na caixa **Ponto** (clique no ponto) → **Copiar ponto**. A cópia é um grupo novo com as mesmas linhas, sem vínculos com cabos; depois você a posiciona com o mouse.
- **Copiar cabo:** no cartão do cabo (lista à direita) → **Copiar**. O cabo novo vai para o fim de Cabos sem pontas ligadas: clique num ponto destacado para ligar cada ponta.

## Mapa: nível MT/BT e cabos que não aparecem (22/11/2026)

- O nível do cabo é o mesmo na lista e no mapa: vale o **Nível** digitado na linha; senão a **tensão escrita no rótulo** (ex.: 13,8kV); senão o padrão do dicionário de cabos; só então a **linha do desenho** do CAD; por fim BT.
- Quando o rótulo diz um nível e a linha do desenho associada é do outro, o cartão do cabo avisa «⚠ Nível MT pelo rótulo, mas a linha do desenho é BT» e o vão fica tracejado: normalmente o CAD ligou o texto à linha errada (redes MT e BT paralelas). Confira no desenho.
- Cabo que não aparece: o cartão diz o motivo — sem linha do desenho e sem ponta ligada (não é desenhado; ligue pelas «Ponta livre»), sem linha do desenho (só a ponta ligada aparece) ou linha do desenho que termina a mais de 5 m de um poste. O filtro **A conferir** lista só esses cabos.

## Mapa: cabo com ponta longe do poste (23/11/2026)

A ponta de um cabo só se **liga** a um poste a menos de 5 m. Quando o desenho tem a ponta mais longe (o cabo termina ao lado do poste), o mapa avisa («⚠ A linha do desenho termina a mais de 5 m de um poste»), mas desenha o cabo junto do ponto mais próximo (até 25 m): se você mover ou **Rearranjar** os pontos, essa ponta acompanha. A ligação para o orçamento continua exigindo as pontas nos postes.

## Mapa: voltar às coordenadas do CAD (23/11/2026)

O mapa coloca cada ponto na coordenada do desenho. Quando você **arrasta** um ponto, cria/copia um ponto, usa **Agrupar pontos** ou **Rearranjar**, a posição fica gravada na obra e passa a valer no lugar da coordenada. O aviso do mapa mostra «N ponto(s) com posição ajustada à mão». O botão **Posições do CAD** descarta esses ajustes e volta cada ponto à coordenada do desenho (Ctrl+Z desfaz). Em **1:500**, se a rede for maior que a janela, o aviso lembra que só parte dela aparece: arraste para andar ou escolha **Ajustar à tela**.


## Mapa: montagem rápida da rede (23/11/2026)

- **Poste sem bloco no CAD:** o ponto aparece onde está a anotação, com o marcador **tracejado laranja** («posição aproximada»). Ele não liga sozinho as pontas dos cabos (só o bloco do CAD faz isso a 5 m).
- **Fundo do CAD** (botão **Fundo**): mostra, esmaecidos, os trechos de rede e os **blocos de poste** do desenho que ainda não viraram linha da obra. Clique num bloco para criar o poste **na coordenada do CAD**. Vem do projeto aberto pelo plugin do nanoCAD (ou do DXF lido no Leitor).
- **Pendências** (topo da lista à direita): barras de **vãos vinculados**, **pontas ligadas** e **pontos com coordenada do CAD**, e a lista do que falta, cada item com o botão que resolve: *Posicionar*, *Usar bloco livre*, *Confirmar posição*, *Ligar a P…*, *Escolher ponto*, *Abrir ponto*. **Ligar N ponta(s) ao ponto mais perto** resolve todas as pontas soltas de uma vez (Ctrl+Z desfaz). Clique no texto do item para ir até ele no mapa.
- **Arrastar a ponta do cabo:** selecione o cabo; aparecem duas alças amarelas perto das pontas. Arraste uma até um ponto (ele acende) e solte para ligar ou trocar a ponta. O vínculo com a estrutura é aplicado quando não há dúvida.
- **Traçar vão (V):** clique no ponto de origem e no de destino. O formulário já vem com o último cabo criado e o **comprimento medido**; confirme com Enter. O destino vira a próxima origem, então dá para seguir a rede clicando ponto a ponto. **Shift+clique** no destino cria sem perguntar. Esc encerra.
- **Atalhos** (fora de campos de texto): **P** poste, **E** estrutura, **C** cabo, **V** traçar vão, **L** ligar a ponta livre do cabo selecionado, **Del** exclui os pontos selecionados (ou o cabo selecionado), **Esc** limpa a seleção/ferramenta.
- **Seleção:** **Shift+arrastar** no fundo do mapa seleciona por caixa; **Shift+clique** num ponto põe/tira da seleção. Arrastar um ponto selecionado move todos juntos.
- **Plugin do nanoCAD:** a anotação de poste que fica longe do bloco e sem linha guia agora é ligada ao **bloco livre mais perto** (cada bloco recebe uma só; até 25 × a altura do texto) e vai para **Conferir**. A que continua sem bloco leva a posição da anotação para o mapa. (Plugin 0.15.0 ou mais novo.)

## Cabo M com duas pontas (23/11/2026)

- O vínculo **automático** liga só as pontas da coluna **Pontas I · M** da Tabela de Cabos (padrão: I = 2, M = 1). Com elas, o cabo fica completo.
- O vínculo **à mão** pode ir além, até a coluna **Máx. à mão I · M** (padrão: I = igual às pontas, M = 2). Assim, num cabo M, o automático nunca liga a 2ª ponta: ela é sempre decisão sua.
- No painel do poste, o cabo M já completo aparece como **Opcional** na outra ponta. Abra **Opções**, escolha a estrutura e clique em **Vincular**.

## Mapa: o ponto com linhas e operação (23/11/2026)

- A barra do mapa tem só **+ Ponto** (P), **+ Cabo** (C) e **Traçar cabo** (V).
- **+ Ponto:** nome e comentário na mesma linha — vira o grupo **«@P12 - comentário»** em Demais Ativos. Embaixo, as **linhas do ponto** (ex.: `12600D`, `N3`, `CE2`), cada uma com a operação: **F** (fictícia), **I M R** e **\*I \*M \*R**. A linha nova vem com a operação da linha de cima. Se alguma linha tem poste, o ponto é **Poste**; senão é **Estrutura avulsa** (ex.: uma N3 solta removendo = ponto com a linha `N3` em **R**).
- **F (fictícia):** a linha fica congelada — **fora do orçamento** —, mas continua recebendo cabos no vínculo. Ponto só com linhas F não entra no orçamento. O poste de uma linha F preenche o parâmetro **POSTE** das estruturas do mesmo ponto (ex.: `12600D` em F + `N3` em I → `N3(POSTE=12600D)`). Cabo com **F** é cabo fictício (congelado).
- **Painel do ponto:** clique numa linha para trocar o texto e a operação ali mesmo (Enter grava, Esc cancela, × tira a linha, **+ Linha** acrescenta). Trocar a estrutura **mantém os vínculos** quando a nova aceita os mesmos cabos; se não aceitar, o painel diz quais vínculos cairiam e pede **Trocar mesmo assim**.
- **Vincular por clique:** no painel, clique no cabo e depois na estrutura (ou na estrutura e depois no cabo).
- **Formato antigo** (ponto sem grupo ou com comentário em linha «#»): o mapa pergunta **Converter todos**; ao abrir um ponto antigo, **Converter este ponto**. Várias linhas «#» viram um comentário só, juntas por «; ». Ctrl+Z desfaz.

## PDF no mapa (23/11/2026)

- **Ler projeto** com um **PDF exportado do CAD** monta o mapa como o DXF: as linhas **vermelhas** viram rede MT e as **verdes**, rede BT (mesma regra de cor do projeto); as pontas da rede valem como postes; cada texto vai para a rede/poste mais perto (ou pela linha-guia).
- **Escala:** o programa escolhe sozinho entre **1:500** e **1:1000** comparando a metragem escrita nos cabos («… 38 m») com o comprimento desenhado. O aviso mostra a escala e quantos cabos confirmaram. Sem nenhum cabo com metragem ligado à rede, usa 1:500 (não confirmada).
- PDF **escaneado** (imagem) não tem posições: os pontos ficam na coluna ao lado e você monta pelo mapa.

## Plugin do nanoCAD: montar no formato do mapa (23/11/2026)

- Em **Montar manualmente**, escolha **as regras de um projeto** ou **sem regras** (botão **Regras / sem regras** para trocar).
- **Selecionar poste:** clique no símbolo do poste (Enter = sem símbolo), selecione os textos da anotação e tecle Enter. Na janela, cada texto é um **campo** com a sua linha (**L**). Marque a operação de cada campo: **F**, **I**, **\*I**, **M**, **\*M**, **R**, **\*R**. **+ Novo campo** e **+ Nova linha** separam o texto. Ctrl+Enter confirma.
- **Selecionar cabo:** clique na linha, escolha **Inteira** ou **Trecho** e selecione os textos. Vale a metragem anotada; a medida na linha aparece ao lado.
- **Com regras**, o ativo e a operação vêm do Leitor do projeto; o texto que não se encaixa entra como **F** (fictício). **Sem regras**, o texto entra como está.
- **Anotação de várias linhas** (um MTEXT com mais de um parágrafo): as linhas de continuação entram no mesmo campo (ex.: «12600D, CEJ2, FLABIT, IT-R, FLABIM, IT-R»); a metragem do cabo não vira campo. Para separar, use **+ Novo campo**. (Plugin 0.16.2.)
- Os pontos e cabos aparecem no painel do plugin, nas abas **Postes e outros** e **Vãos e cabos**. **Duplo clique** (ou **Ajustar ponto/cabo…**) mostra o item no desenho e reabre a janela para corrigir ou **Excluir**. (Plugin 0.16.1.)
- **Adicionar anotações ao OrçaCore** manda os pontos e cabos para a obra aberta: cada poste aparece em Demais Ativos como **«@P1 - comentário»** com as suas linhas, e cada cabo em Cabos. Faça os vínculos no **mapa**. (Plugin 0.16.0 ou mais novo.)

## Mapa: um ponto só com todas as operações (24/11/2026)

- Todas as linhas de um ponto «@P1» aparecem como **um único ponto** no mapa, qualquer que seja a operação (I, \*I, M, \*M, R, \*R, F), mesmo sem coordenada do CAD.
- No painel do ponto, a lista **Estruturas** mostra todas as estruturas com a operação. As que estão **instalando ou mantendo** (I, \*I, M, \*M) ficam em **destaque verde** e recebem cabos. As que estão **removendo** (R, \*R) aparecem apagadas e riscadas, fora do vínculo.
- Quando o ponto tem um poste em **F** (existente) e outro instalando, o nome do ponto mostra o poste instalando.

## Leitura do DXF: cabos sem traçado (24/11/2026)

- **Rede em outra layer:** se a rede de MT (ou BT) do desenho está numa layer que não é REDE_MT/REDE_BT — por exemplo o tracejado vermelho na **ANOTACOES** —, inclua essa layer em **⚙ Leitura → Layers da rede MT** (aba **Desenho** do vínculo, ou ⚙ Leitura no plugin). Para só a layer «0», escreva **=0**.
- **Anotação dentro de moldura:** a linha-guia que sai da borda do retângulo em volta do texto agora é usada.
- **MT e BT lado a lado:** o cabo que diz a tensão (kV, A35P → MT; PB, RAMA, C16( → BT) só pega a rede do seu tipo.
- **Marcas de X** riscando anotações não contam como rede.
- **Cartão do cabo:** os botões das pontas mostram o mesmo nome do ponto que o mapa (P1, P2…).
- Cabo que continua sem traçado: no cartão, **Ponta livre** → clique no ponto, para cada ponta.

## Mapa: cabos lado a lado (24/11/2026)

- Vários cabos entre os mesmos dois pontos aparecem **lado a lado**, nunca um em cima do outro — com ou sem traçado do CAD, e em qualquer sentido. A MT fica de um lado, a BT do outro.
- O espaço entre eles é de pelo menos 9 px na tela, mesmo com o zoom afastado.
