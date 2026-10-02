# Roteiro de testes no Windows — Modo autônomo (TASK-031)

**Objetivo:** confirmar, no **seu** Windows e com o **executável** (`Leitor_Projetos.exe`), o que só foi provado em Linux: o motor `quickjs`, o processo auxiliar do leitor, mover arquivos, o fluxo completo e a comparação com o modo manual.
**Tempo:** cerca de 1h30 (A–D) + a comparação com um projeto real (E).
**Regra de ouro:** só marque **OK** se o resultado for **exatamente** o esperado. Em qualquer dúvida, marque **FALHA**, anote o que viu e continue — não tente consertar no meio.

## O que levar de volta para a conversa
1. O bloco **"COPIE ISTO"** do diagnóstico (Parte A).
2. A tabela do final deste roteiro com OK / FALHA / N/A em cada teste e, nas falhas, o que apareceu na tela (print) e o conteúdo do `.erro.txt` ou do botão **Detalhes** do histórico.
3. Para a Parte E: as duas planilhas lado a lado (veja E3).

---

## Parte 0 — Preparação (10 min)

| # | O que fazer | Esperado |
|---|---|---|
| 0.1 | **Feche** o Leitor de Projetos. Copie `banco_resumo.db` (ao lado do `.exe` ou do `app.py`) para um lugar seguro, por exemplo `banco_resumo_ANTES.db`. | Backup feito. É a sua volta atrás. |
| 0.2 | Crie uma pasta de testes **fora** do OneDrive/Dropbox, ex.: `C:\TesteAutonomo`. Dentro dela crie `amostras`. | Pastas criadas. |
| 0.3 | Anote: versão do Windows, se tem antivírus além do Defender, e se o PC é de rede corporativa. | Para o relatório. |
| 0.4 | Se o programa usa a nuvem (Supabase), prefira testar com um projeto/usuário de **teste** — as obras geradas vão para a nuvem como as manuais. | — |

## Parte A — Ambiente de desenvolvimento (Python) (15 min)

Abra o **Prompt de Comando** na pasta do projeto (onde está o `app.py`).

| # | Comando / ação | Esperado |
|---|---|---|
| A1 | `python --version` | `Python 3.11.x` |
| A2 | `pip install -r requirements.txt` | Termina sem erro, instalando o `quickjs`. Se aparecer *"Microsoft Visual C++ 14.0 is required"*, o `pip` tentou **compilar** em vez de baixar o pacote pronto → **FALHA** (anote a versão do Python/pip). |
| A3 | `python scripts\autonomo_verificar_ambiente.py` | Todas as linhas `[OK]` (as de "node" e "mover arquivo aberto" podem ser `AVISO`/`OK`). Copie o bloco **COPIE ISTO**. |
| A4 | `pip install -r requirements-dev.txt` e depois `python -m pytest tests -q` | Anote a linha final (esperado: **0 failed**; no Linux foram 647 aprovados). Falhas de teste = copie o nome do teste e a mensagem. Se o `node` não estiver instalado, os testes de "paridade" são **pulados** (normal; só importa o número de falhas = 0). |
| A5 | `python app.py` — com o console aberto, deixe rodando e abra o modo autônomo (Parte C, C1–C3) em modo desenvolvimento **antes** de testar o `.exe`. | A tela abre. O console mostra os logs (útil se algo falhar). |

**Se A3 mostrar FALHA em "Processo auxiliar..." ou "Dependência quickjs": pare aqui e me envie o bloco — o `.exe` não vai funcionar sem isso.**

## Parte B — Gerar o executável (15 min)

| # | O que fazer | Esperado |
|---|---|---|
| B1 | Abra o `build.bat` e confira se o caminho do `pyinstaller.exe` (primeira linha longa) existe no seu PC. Já inclui `--hidden-import quickjs --hidden-import _quickjs`. | Caminho válido. |
| B2 | Execute `build.bat`. | Termina com *"Building EXE ... completed successfully"* e cria `dist\Leitor_Projetos.exe`. Procure no texto da saída por `quickjs` ou `_quickjs` em **WARNING** (ex.: *"Hidden import '_quickjs' not found"*) → anote. |
| B3 | Copie `dist\Leitor_Projetos.exe` para `C:\TesteAutonomo\app\` (junto, copie a pasta `static`? **não** — o `.exe` já embute tudo). | O `.exe` está na pasta de teste. O banco será criado/lido **ao lado do .exe** (`banco_resumo.db`). Para testar com seus dados reais, copie o seu `banco_resumo.db` para essa pasta. |
| B4 | Dê um duplo clique no `.exe`. | A janela abre na tela de login (pode demorar alguns segundos na 1ª vez — anote quantos). |

## Parte C — Testes no executável (50 min)

Faça login como **administrador** → **Painel Admin** → botão **Abrir o modo autônomo**.

Gere as amostras: `python scripts\autonomo_arquivos_teste.py gerar C:\TesteAutonomo\amostras`

**Antes dos testes T5–T6** crie, no **projeto 027 (PARAIBA)** ou no projeto de teste, um ajuste que **exclui um item**: Resumo → **Regras de validação** → aba **Ajustes** → *Novo ajuste* → ação **Remover ativo da linha**, tabela `outros`, ativo `U3`, marque **ativo** → **Salvar**. (Instruções no manual, seção 4. Sem esse ajuste os arquivos "com exclusão" não geram confirmações.)

| # | Ação | Esperado | OK? |
|---|---|---|---|
| **C1** | Na tela do modo autônomo: em *Obras geradas em nome de* escolha **Eu mesmo**; *Pasta base* `C:\TesteAutonomo\autonomo`; intervalo `3`; espera `3` → **Salvar configuração**. | Mensagem verde "Configuração salva." As pastas ainda não existem (são criadas ao ligar). | |
| **C2** | Clique **Ligar**. | Selo **Ligado**. Em `C:\TesteAutonomo\autonomo\` aparecem `entrada`, `processados`, `erros`, `saida`. **Nenhuma janela de console pisca.** | |
| **C3** | Crie `C:\TesteAutonomo\autonomo\entrada\027\` e **copie** `01_normal.dxf` para dentro. Aguarde ~10 s. | O arquivo some da entrada e aparece em `processados\027\`. No histórico: situação **Concluída** ou **Com pendências** (o importante é ter gerado resultado). Em `saida\027\01_normal-xxxx\` existem `cabos.csv`, `outros.csv`, `tabelas.json`, `orcamento.json`, `orcamento.csv`, `relatorio.json`. Abra os CSV no Excel (acentos corretos, colunas separadas por `;`). | |
| **C4** | **Gerenciador de Tarefas** (Ctrl+Shift+Esc) → procure `Leitor_Projetos`. | Há **2** processos `Leitor_Projetos.exe` (o programa e o auxiliar do leitor, que sobe ao processar o 1º arquivo). | |
| **C5** | Em **Detalhes** da execução do C3, confira a linha **ramais** (a amostra tem "TROCAR 2 RS M AC") e "Ajustes". | Etapas listadas sem `erro`; ramais com 1 item e 2 linhas geradas. | |
| **T3** | **Cópia lenta:** `python scripts\autonomo_arquivos_teste.py copiar-devagar C:\TesteAutonomo\amostras\07_grande.dxf C:\TesteAutonomo\autonomo\entrada\027\07_grande.dxf --segundos 20` | **Enquanto copia**, o arquivo **não** é processado. Só depois que termina e fica parado ~3 s ele vai para `processados`. Nenhum erro. | |
| **T4** | **Arquivo em uso:** `python scripts\autonomo_arquivos_teste.py copiar-e-travar "C:\TesteAutonomo\amostras\08_nome com espaço e acentuação ção.dxf" C:\TesteAutonomo\autonomo\entrada\027\08.dxf --segundos 60` (copia o arquivo e o **mantém aberto**). Enquanto isso, tente movê-lo no Explorer. | O Explorer **recusa** mover/apagar. O programa processa o arquivo **uma única vez** (aparece **uma** linha no histórico), mas **não consegue movê-lo** e **não dá erro**: o arquivo fica na entrada, sem virar "duplicado" nem ser processado de novo. Quando o script termina (60 s), o arquivo vai sozinho para `processados\027` e o histórico passa a apontar para o novo local (botão **Reprocessar** funciona). | |
| **T5** | **Exclusão — "Sim" e "Sim para todos":** copie `02_com_exclusao.dxf` para `entrada\027`. | Em ~10 s aparece o cartão **Aguardando a sua confirmação** com 2 itens ("Editar OUTROS-0 … → …"). O arquivo já está em `processados`, mas **não há obra nem orçamento** (`saida\...` só tem `pendencias.json` e as tabelas). Clique **Sim** no 1º item → confirma o aviso → sobra 1. Clique **Sim para todos** → conclui. A obra e o orçamento passam a existir. | |
| **T6** | **Exclusão — "Não para todas":** copie `03_com_exclusao_b.dxf`. Clique **Não para todas**. | Termina como **Com pendências**; a linha com `U3` continua nas tabelas (`outros.csv`). | |
| **T7** | **Duplicado:** copie `01_normal.dxf` de novo com outro nome (`01_copia.dxf`). | Vai para `processados\027\duplicados\`. Não cria nova obra. | |
| **T8** | **Erros:** copie `05_quebrado.dxf`, `06_nota.txt` para `entrada\027`; copie `04_sem_nada_a_excluir.dxf` direto na raiz de `entrada` (solto); e crie a pasta `entrada\ZZZ` com um DXF. | Todos vão para `erros\` (`erros\027`, `erros\sem_projeto`, `erros\ZZZ`), cada um com um `.erro.txt` explicando. Aparecem no histórico como **Erro**. O programa continua normal. | |
| **T9** | **Reverter e reprocessar:** no histórico, **Reverter** na execução do C3; depois **Reprocessar**. | Reverter → **Revertida**. Reprocessar → nova linha no histórico. | |
| **T10** | **Religar sozinho:** com o modo **Ligado**, feche o programa (X da janela) e abra de novo. | Ao abrir e ir em Modo autônomo, já está **Ligado**. Um arquivo novo na entrada é processado. | |
| **T11** | **Encerrar sem órfão:** com o modo ligado e após processar algo, feche o programa. Olhe o Gerenciador de Tarefas por ~10 s. | **Nenhum** `Leitor_Projetos.exe` sobra. (Se sobrar, é o processo auxiliar órfão → **FALHA**.) | |
| **T12** | **Caminho com acento/espaço/rede:** mude a pasta base para `C:\Users\<você>\Área de Trabalho\Autônomo Teste` e repita C3. Se puder, repita com uma pasta de **rede/OneDrive**. | Funciona igual (em rede/OneDrive podem ocorrer atrasos: anote). | |
| **T13** | **Uso manual junto:** com o modo ligado e um arquivo sendo processado, abra o **Resumo** manual com um arquivo e use Ajustar/Montar Orçamento. | O modo manual funciona normalmente, sem travar. | |
| **T14** | **Tempo:** anote o tempo do 1º arquivo (inclui o início do auxiliar) e dos seguintes. | Informativo. Se o 1º demorar mais de ~30 s, anote (antivírus escaneando o auxiliar). | |
| **T15** | **Desligar:** clique **Desligar**. Copie um arquivo na entrada. | Nada é processado; o arquivo fica na pasta. | |

## Parte D — Operador e segurança (10 min)

| # | Ação | Esperado | OK? |
|---|---|---|---|
| D1 | Entre com um usuário **operador** e abra `/autonomo` (digite o endereço ou use o link). | É barrado e volta para a página inicial. | |
| D2 | O operador usa o Resumo normalmente (Validar, Ajustar, Montar orçamento). | Tudo como antes. | |
| D3 | Veja se os **ajustes** e **regras** continuam iguais para o modo manual depois dos testes. | Sem mudanças. | |

## Parte E — Comparação com o modo manual (a mais importante)

**Objetivo:** provar que, para o **mesmo arquivo**, o autônomo fecha **igual** ao seu processo manual.
Use **um projeto real seu**, com o orçamento que você considera correto. Faça com cuidado para comparar nas mesmas condições:

1. **Desligue todos os ajustes** do projeto (aba **Ajustes** da gaveta: nenhum "ativo") — o manual não aplica ajustes sozinho, o autônomo sim. (Depois você repete com os ajustes ligados para ver as diferenças.)
2. **Modo manual:** abra o arquivo no Leitor → **Processar** → Resumo → **sem editar nada** → anote/copie as tabelas **Cabos** e **Outros** e clique **Montar Orçamento** (anote as linhas e totais por código).
3. **Modo autônomo:** copie o **mesmo** arquivo em `entrada\<código-do-projeto>\`. Compare `cabos.csv`, `outros.csv` e `orcamento.csv` (em `saida\…`).

| # | Comparar | Esperado |
|---|---|---|
| E1 | `cabos.csv` × tabela **Cabos** do Resumo | Mesmas linhas, mesma ordem, mesmas **quantidades** (`qtdAtivos`). |
| E2 | `outros.csv` × tabela **Outros** | Iguais, **mais** as linhas "RAMAIS (GERADO)" no fim, se houver ramais (no manual elas só entram quando você clica em "Adicionar" no modal RAMAIS). |
| E3 | `orcamento.csv` × resultado do orçamento | Mesmos códigos, operações e **totais**. Anote qualquer diferença com: o ativo, o valor manual, o valor autônomo. |
| E4 | `relatorio.json` → `orcamento.nao_encontrados` | Igual à lista de "não encontrados" do manual. |
| E5 | Repita com os ajustes **ligados** e confira se as diferenças são exatamente os ajustes (Detalhes → "Ajustes aplicados"). | Diferenças = só o que os ajustes fazem. |

**Atenção:** se o seu projeto **não tem** regras de conversão salvas, o manual (em um navegador novo) e o autônomo aplicam a regra padrão **CABOS × 1,05**. Se tiver regras salvas, ambos usam as salvas. Se a diferença estiver nos cabos, confira isso primeiro.

## Se algo falhar — o que me enviar
- O bloco "COPIE ISTO" da Parte A.
- O teste que falhou (ex.: **T11**) e o que apareceu (print).
- O `.erro.txt` da pasta `erros`, o **Detalhes** da execução e o `relatorio.json`.
- Para falha no `.exe` que não aparece em `python app.py`: diga se o processo auxiliar (C4) aparece no Gerenciador de Tarefas.
- Possíveis causas já conhecidas: (a) o `quickjs` não empacotou (procure `quickjs` nos avisos do B2); (b) o auxiliar não inicia dentro do `.exe` — o sintoma é a execução ficar em **Erro** com a mensagem *"O motor do leitor não conectou"*; (c) antivírus bloqueando o auxiliar.

## Desfazer tudo
1. Clique **Desligar** e feche o programa.
2. Apague `C:\TesteAutonomo\autonomo` (e as obras `[Autônomo] …` pela tela "Carregar obra", se quiser).
3. Se mexeu no banco real e algo ficou estranho, restaure `banco_resumo_ANTES.db` (com o programa fechado).

## Tabela de resultados (preencha)

| Teste | OK / FALHA / N/A | Observações |
|---|---|---|
| A1–A5 | | |
| B1–B4 | | |
| C1–C5 | | |
| T3 | | |
| T4 | | |
| T5 | | |
| T6 | | |
| T7 | | |
| T8 | | |
| T9 | | |
| T10 | | |
| T11 | | |
| T12 | | |
| T13 | | |
| T14 (tempo do 1º arquivo) | | |
| T15 | | |
| D1–D3 | | |
| E1–E5 | | |
