# Plugin nanoCAD — Orçamento (Leitor de Projetos)

Subproduto do **Leitor de Projetos Pro**: um plugin C# (.NET) que roda **dentro do nanoCAD**, lê o desenho da rede
(postes, vãos e anotações) e produz uma **obra totalmente parametrizada**, gravada como **projeto de trabalho do OrçaCore**
(`<desenho>.obra`, na pasta do desenho — o mesmo arquivo do 📁 Projeto do programa), que a aplicação
importa em **Carregar Obra** (`POST /api/obras/importar`) para gerar o orçamento. O cálculo **não** roda no CAD.

> Documentação para IA/desenvolvimento: [`.ai/`](.ai/) — requisitos, estado, arquitetura, regras e tarefas do plugin.
> Mockups do resultado esperado: [`referencias/mockups/`](referencias/mockups/).

## Usar no nanoCAD

| nanoCAD | Arquivo para o APPLOAD | Como obter |
|---|---|---|
| **26** (.NET 6+) | [`appload/nanoCAD26/OrcamentoNanoCAD.dll`](appload/nanoCAD26/) | já vem pronto nesta pasta |
| **5** (.NET Framework 3.5) | [`appload/nanoCAD5/OrcamentoNanoCAD.dll`](appload/nanoCAD5/) | DLL x86 para o runtime CLR 2. A confirmação do carregamento no nanoCAD 5 está pendente. Alternativa com o SDK instalado: `build_nanocad5.bat` no Windows com o nanoCAD 5 instalado (compila contra as DLLs dele) |

**Junto com o programa principal:** o `OrcaCore.exe` traz as duas DLLs e as instala em `%APPDATA%\OrcamentoNanoCAD\plugin\<nanoCAD26|nanoCAD5>\`
(botão **Plugin nanoCAD** na tela inicial do programa: caminho, versão, instalar agora). No nanoCAD, uma vez: `APPLOAD` → **Inicialização** → adicionar esse arquivo.

1. No nanoCAD: `APPLOAD` (ou `NETLOAD`) → escolha a DLL. Para carregar sempre, adicione-a à lista de inicialização do APPLOAD.
2. Comando **`ORCACORE`** ou atalho **`OC`** → *Montar manualmente*, *Ler projeto* ou *Continuar obra salva*; escolha as regras (ex.: Pernambuco 081).
3. Confira no painel — **CABOS** (vãos) e **POSTES E ESTRUTURAS** em listas separadas, como no Resumo da aplicação; filtros
   Ok / Conferir / Erro / Ignorados. Na leitura só aparece o que as regras do leitor classificaram (rede existente e poste sem
   anotação ficam ocultos); em *Montar manualmente* aparece tudo o que você montou. Ajuste no pop-up (duplo clique ou `ORC_EDITAR`): corrija o ativo
   com o autocompletar da Base de Orçamento, exclua (a linha lida fica em **Ignorados** e volta com "Restaurar") ou crie linhas digitadas
   (**+ Cabo** / **+ Outros**). Tudo isso sobrevive ao ↻ Atualizar. Depois **Salvar**.
4. **🌐 Abrir na aplicação** (`ORC_ABRIR_APP`): salva o projeto e o abre no OrçaCore deste computador (aberto: a janela vem para a frente
   com o projeto; fechado: o programa abre já com ele). Ou, no OrçaCore: 📁 Projeto → Abrir → `<desenho>.obra`. Nada é gravado no banco.

### Intercâmbio de desenhos

Abra **`ORCACORE`** (ou **`OC`**) e escolha **Intercâmbio de desenhos**. As mesmas ações também estão disponíveis na linha de comando.

- **`ORC_EXPORTAR_JSON`** gera o JSON do Model Space. **`ORC_IMPORTAR_JSON`** acrescenta a geometria de um JSON ao desenho ativo.
- **`ORC_GERAR_DXF`** cria um DXF separado para revisão. Abra o arquivo com **`ORC_ABRIR_DXF`**.
- **`ORC_SINCRONIZAR_JSON`** aplica alterações incrementais de um patch v1 no desenho de origem, com confirmação e backup. A opção é experimental. Confira o resultado em uma cópia do DWG.
- Para recuperar a geometria visual de objetos proxy, abra uma vez o OrçaCore atualizado e execute **`ORC_CONVERSOR_DXF`** no nanoCAD para conferir o registro do conversor local. A recuperação usa os gráficos que o desenho contém. Objetos sem gráfico recuperável geram um aviso.
- **`ORC_AJUDA_DESENHO`** lista os comandos. **`ORC_SUBSTITUIR_DESENHO`** e **`ORC_ATUALIZAR_DESENHO`** trabalham com geometria nativa e pedem confirmação e backup.

### Conferir uma leitura

- Abra **Pendências** para ver avisos sobre destino, vínculos, duplicidades, operação e textos de interesse fora da obra. Use **Buscar** ou selecione o poste pela placa.
- Um clique destaca a anotação e a geometria. Duplo clique aproxima a vista. **Conferido** avança a fila. **Só mudanças** compara o resultado com a leitura anterior após **Atualizar leitura**.
- **Escolher destino** permite confirmar um dos candidatos exibidos. **Ajustar** abre os campos de ativo, operação, metros e vínculos para o item selecionado.
- Use Ctrl ou Shift para selecionar várias linhas e abrir **Operação em lote**. A prévia mostra o que mudará na obra e oferece uma opção explícita para trocar a layer do texto no DWG.
- Em **Conferir → Cores por operação**, vermelho indica instalação, magenta indica remoção e azul indica reaplicação. O asterisco e o rótulo LV identificam linha viva. O modo usa marcas temporárias.
- Em **Configurações → Leitura do desenho**, configure a tolerância de metragem em metros para o projeto. O valor inicial zero desativa esse aviso. O detalhe do vão sempre pode mostrar as três medidas disponíveis.

### Anotações e auditoria de layers

- **Anotações → Criar nota** (`ORC_NOTA`) cria um MTEXT com operação, altura, alinhamento e guia opcional. **Editar nota** e **Mover nota** trabalham com MTEXT. No movimento, você pode selecionar textos acompanhantes, alinhar o grupo ou limitar o deslocamento à vertical.
- **Reapontar guia** (`ORC_REAPONTAR_NOTA`) associa uma nota a outro objeto. **Atualizar guias** recalcula os extremos. A atualização automática usa os eventos de movimento do nanoCAD quando disponíveis.
- **Auditoria de layers** (`ORC_AUDITAR_LAYERS`) lista as notas por operação e as que estão fora. Ctrl+A seleciona a aba. Um clique destaca e duplo clique aproxima. Alterar layers, alternar LV e excluir usam os comandos de desenho. A exclusão pede confirmação.
- A aba **Análise** permite atribuir operações apontando no desenho. O filtro de visibilidade do 26 atua na viewport. No 5, o filtro controla as marcas de análise.
- Use **Atualizar leitura** após editar o DWG para incorporar as mudanças à obra e enviar a versão revisada ao OrçaCore.

### Comandos

| Comando | O que faz |
|---|---|
| `ORCACORE` / `OC` | tela inicial. O alias `ORCAMENTO` continua disponível. |
| `ORC_LER` / `ORC_RELER` / `ORC_ATUALIZAR` | lê o desenho pelas regras (layer + cor) · reler do zero (descarta ajustes) · reler mantendo os ajustes |
| `ORC_CONSTRUIR`, `ORC_POSTE`, `ORC_VAO`, `ORC_VINCULAR` | construir do zero: apontar poste (bloco + anotação) e vão (linha + anotação); refazer vínculos automáticos |
| `ORC_MODO_MANUAL`, `ORC_PONTO_MAPA`, `ORC_CABO_MAPA`, `ORC_MAPA_ITEM`, `ORC_MAPA_ZOOM`, `ORC_MAPA_EXCLUIR`, `ORC_MAPA_OBJETO` | montagem manual no formato do mapa: regras de um projeto ou sem regras · Selecionar poste (símbolo + textos → campos com operação F, I, \*I, M, \*M, R, \*R) · Selecionar cabo (linha inteira ou o trecho clicado da polilinha + textos) · ajustar/excluir o item escolhido no painel |
| `ORC_EDITAR` | pop-up do vão/poste: entrada (lida do desenho), saída (para a obra), ajustes — **ativo editável com o autocompletar da Base de Orçamento** (sugestões, parâmetros, `)` automático; não altera o desenho), operação, metros, estrutura do vínculo, troca opcional da layer do texto, **Excluir da obra / Restaurar** |
| `ORC_NOVO_CABO`, `ORC_NOVO_OUTROS` | linha nova (+ Cabo / + Outros no painel): digitada com o autocompletar, ou **lida de um TEXT/MTEXT escolhido no desenho** pelas regras do leitor. Em + Cabo o texto é lido como cabo e o comprimento vem da mesma linha ou do texto "35 m" ao lado/embaixo no desenho |
| `ORC_VINCULOS`, `ORC_ATIVOS` | tabela de vínculos · ativos com parâmetros (`N4(9400D, P, 3XA35P_7.94_-13.8KV, ?)`) |
| `ORC_SALVAR` | grava o **projeto `<desenho>.obra`** na pasta do desenho (formato do 📁 Projeto do OrçaCore: linhas + a obra do plugin; SAP e resultado do programa ficam) e a cópia **no DWG**. Se as linhas foram alteradas no OrçaCore, pergunta antes de gravar por cima |
| `ORC_PROJETO` | nome e número de obra do projeto (o número vai para o REC) |
| `ORC_ABRIR_APP` | salva o projeto e o abre no OrçaCore deste computador: aberto, recebe o projeto e vem para a frente (Abrir/Mesclar/Sobrescrever se a tela já tem dados); fechado, abre já com ele — 🌐 no painel. Sem o programa instalado, o projeto fica na pasta do desenho |
| `ORC_LEITURA`, `ORC_DIAGNOSTICO` | onde ficam rede, postes e guias neste padrão de desenho · o que o plugin leu em cada texto (arquivo + tabela) |
| `ORC_FECHAR` | fecha a obra (⌂ Início no painel volta à tela inicial) |
| `ORC_ATUALIZAR_PLUGIN`, `ORC_VOLTAR_VERSAO` | conferir agora se há versão nova do plugin · voltar para a versão anterior |
| `ORC_CONECTAR`, `ORC_SINCRONIZAR` | ligar à aplicação aberta (localhost/online): regras e base de orçamento ao vivo · baixar e conferir agora |
| `ORC_CONTINUAR`, `ORC_PAINEL`, `ORC_ZOOM`, `ORC_MARCAS`, `ORC_REGRAS`, `ORC_AJUDA` | abrir a obra salva · reabrir o painel · localizar · ligar/desligar números e vínculos na tela · recarregar regras · ajuda |

### Atualização automática

A partir da 0.6.0 o plugin se atualiza sozinho: 20 s depois de abrir o nanoCAD e a cada 4 h ele pergunta à aplicação (o endereço do
`ORC_CONECTAR`, sem precisar de login) se há versão nova, baixa, confere a impressão digital (SHA-256) e troca a DLL. A versão nova
**vale quando você reabrir o nanoCAD**. `ORC_ATUALIZAR_PLUGIN` confere na hora; `ORC_VOLTAR_VERSAO` volta para a anterior. Desligar:
`ORC_CONECTAR` → "Atualizar o plugin sozinho". Deixe a DLL numa pasta onde você pode gravar (não em "Arquivos de Programas").
Publicar uma versão: `./compilar.sh` (DLL + `appload/nanoCAD26/versao.json`), mesclar no main e publicar a aplicação.

### Ligado à aplicação (regras e base ao vivo)

Com a aplicação aberta (`http://127.0.0.1:8000` ou o endereço online): `ORC_CONECTAR` → e-mail e senha → Entrar. A tela
inicial passa a listar **todos os projetos cadastrados** ("Aplicação: PERNAMBUCO (081)"…): as regras do projeto escolhido são baixadas na hora,
junto com a **configuração de leitura** guardada na aplicação (⚙ Leitura salva lá e vale em qualquer computador). O plugin consulta a aplicação a cada 5 s (ajustável): regras
do vínculo mudaram → recalcula os vínculos; base de orçamento mudou → confere de novo. Use `ORC_ATUALIZAR` para reler o DXF com regras novas. Estrutura, cabo
ou poste que **não está na base** vira *Conferir* no painel ("não está na base de orçamento"). Sem a aplicação, segue com a
última cópia das regras do mesmo projeto, se já baixada. A senha não é guardada e o acesso vale até 30 dias.

Para **Paraíba**, digite `027` no campo **Projeto na aplicação** ou selecione **Aplicação: PARAIBA (027)**. A configuração de layers, blocos e vínculo pode vir do pacote da aplicação ou da última cópia baixada. Quando a aplicação não tiver salvo um perfil de geometria, configure o desenho em **⚙ Leitura**.

### Usar a mesma leitura do DXF

Em **ORCACORE → Ler projeto**, o plugin lê o desenho ativo e consulta o OrçaCore conectado pelo **ORC_CONECTAR**. Um DXF salvo e sem alterações é enviado diretamente. Para um DWG ou um DXF com alterações, o nanoCAD gera um DXF temporário com o estado atual do desenho, incluindo as edições ainda não salvas. O plugin remove essa cópia após a leitura. **ORC_RELER**, **ORC_ATUALIZAR** e **ORC_DIAGNOSTICO** usam o mesmo fluxo.

O OrçaCore aplica o extrator e o motor JavaScript do Leitor com as regras salvas para o código selecionado. O plugin recebe as linhas prontas e associa postes, vãos e vínculos. Uma linha reconhecida fica na obra mesmo quando não encontra geometria. O log e o diagnóstico registram o projeto e a identificação das regras. A leitura automática exige que o OrçaCore conectado disponibilize a rota de leitura DXF.

### Quando a leitura não acha a rede, os postes ou as guias

1. **⚙ Leitura** (`ORC_LEITURA`): informe as layers da rede MT/BT, o nome (ou a layer) do bloco de poste — círculo também vale — e,
   se quiser, as layers das linhas guias. A tela mostra as layers e os blocos do desenho. Fica guardado por projeto neste computador.
2. **🩺 Diagnóstico** (`ORC_DIAGNOSTICO`): grava `<desenho>.diagnostico.json` e mostra, para cada texto, o que as regras fizeram e
   para onde ele foi (poste, vão, avulso ou ignorado + motivo). Mande esse arquivo para ajuste fino; o log fica em
   `%APPDATA%\OrcamentoNanoCAD\log.txt`.
3. **Dicas** que o `ORC_LER` e o diagnóstico mostram (💡): linhas vermelhas/verdes de poste a poste numa layer que não é de rede
   (inclua a layer em MT/BT) e blocos nas pontas soltas da rede (inclua em Blocos de poste). `=NOME` exige o nome exato
   (ex.: `=0` = só a layer 0; sem `=` vale qualquer layer que contenha o trecho).
4. Várias anotações no mesmo vão ou poste (recondutagem, troca de poste) viram **irmãos** (`V7b`, `P3b`), cada um com sua linha
   na obra. A **placa** do poste (ex.: B039571) é lida e aparece no painel.

### Regras

As regras do leitor e do vínculo são **os mesmos arquivos** de `../data/` do projeto principal, embutidos na DLL
(Pernambuco 081, regras de fábrica do leitor de Paraíba 027 e padrão genérico). Para usar outras, crie ao lado da DLL a pasta `regras\<codigo>\` com
`processamento.json`, `classificacao.json`, `vinculacao.json` (exportados da aplicação) e, opcionalmente, `nome.txt`;
depois `ORC_REGRAS`.

## Pastas

```
plugin_nanocad/
├── .ai/                  requisitos, estado, arquitetura, regras, decisões e tarefas DO PLUGIN
├── appload/nanoCAD26/    DLL pronta para o APPLOAD (nanoCAD 26)
├── appload/nanoCAD5/     saída do build_nanocad5.bat
├── exemplos/             .obra (projeto) e .obra.json gerados pelo núcleo (os testes de contrato da aplicação abrem/importam)
├── referencias/          mockups e arquivos do resultado esperado (DWG/DXF, imagens…)
├── src/Nucleo/           lógica SEM CAD (leitura, regras, vínculo, .obra.json) — testável no Linux
├── src/Plugin/           camada do nanoCAD (comandos, telas, DWG, marcas na tela)
├── tests/Nucleo.Testes/  testes xUnit do núcleo (inclui paridade com o JS/Python da aplicação)
├── tests/Net40/          verificação adicional do núcleo em .NET 4 / C# 5
├── src/nanoCAD5/         projeto da DLL do nanoCAD 5 (.NET 3.5), mesmo código
├── src/ApiNanoCAD5/      referências da API (só o que o plugin usa), GERADAS por ferramentas/GerarApiNanoCAD5
├── tests/oraculo/        oráculos: rodam o código da APLICAÇÃO para gerar o esperado
├── build_nanocad5.bat    compila para o nanoCAD 5 no Windows
├── compilar_nanocad26.bat recompila a DLL do nanoCAD 26 (precisa do .NET SDK)
└── compilar.sh           desenvolvimento: testes + verificação .NET 4 + DLL do nanoCAD 26
```

## Desenvolvimento

```bash
cd plugin_nanocad
./compilar.sh                         # testes do núcleo + net40 + appload/nanoCAD26/OrcamentoNanoCAD.dll
ATUALIZAR_EXEMPLOS=1 ./compilar.sh    # também regrava exemplos/Projeto_Rua_das_Flores.obra(.json)
cd .. && python -m pytest tests/test_plugin_nanocad_contrato.py   # contrato com a aplicação
```

Precisa do .NET SDK 8. A API do nanoCAD 26 vem do pacote NuGet `ModPlus.nanoCAD.API.26` (só para compilar).
O ambiente de desenvolvimento não roda o nanoCAD: o teste dentro do CAD é feito pelo usuário, no Windows.

### Ler uma região e acrescentar anotações

Em **Montar manualmente**, escolha o projeto e clique em **Ler por região**. Clique nos vértices no nanoCAD, confirme com Enter e revise a grade. **Limpar** esvazia a montagem da sessão. Use o editor para mudar texto, operação e marcas da linha. Abra a obra de destino no Resumo e clique em **Adicionar anotações ao OrçaCore**. O envio usa o mesmo projeto selecionado nas duas telas. A adição pode ser desfeita com Ctrl+Z no Resumo.

### Montar no formato do mapa (0.16.0)

Ao entrar em **Montar manualmente**, escolha **as regras de um projeto** ou **sem regras** (troque depois em **Regras / sem regras**, `ORC_MODO_MANUAL`).

- **Selecionar poste** (`ORC_PONTO_MAPA`): clique no bloco ou círculo do poste (Enter = ponto sem símbolo), selecione os textos da anotação (vários de uma vez) e tecle Enter. Na janela, cada texto é um campo com a sua linha (L). Marque a operação de cada um: **F** (fictício), **I**, **\*I**, **M**, **\*M**, **R**, **\*R**. **+ Novo campo** separa um texto em mais campos na mesma linha; **+ Nova linha** cria a linha seguinte. Ctrl+Enter confirma.
- **Selecionar cabo** (`ORC_CABO_MAPA`): clique na linha e selecione os textos. LINE vale inteira; na polilinha vale o **trecho clicado**, de vértice a vértice. A janela mostra o **medido na linha** (informação) e o **comprimento usado** (vai para a anotação; padrão = medido arredondado, sem casas decimais). Cada linha do cabo tem o seu **m**: a metragem do texto («PB70 32 m» ou «PB70/32 m») fica naquela linha; as outras usam o comprimento usado.
- Com regras, o ativo e a operação vêm do Leitor do projeto; o texto que não se encaixa entra como **F**. Sem regras, o texto entra como está.
- Os pontos e cabos aparecem nas abas **Postes e outros** e **Vãos e cabos** do painel e, na tela, o que já foi coletado fica marcado (pontos em ciano, cabos em laranja, enviados em cinza) enquanto o painel está aberto. Clique num item: o desenho enquadra e destaca o item com o poste/linha dele; **Editar**, **Excluir** e **Definir poste/cabo** trocam os campos, apagam o item ou escolhem outro objeto base (`ORC_MAPA_ITEM`, `ORC_MAPA_EXCLUIR`, `ORC_MAPA_OBJETO`).
- O tipo do cabo (MT/BT) vem da **cor** da linha escolhida; sem cor de rede, da layer e, por fim, do texto.
- **Adicionar anotações ao OrçaCore** manda os pontos («@P1 - comentário» com as linhas) e os cabos para a obra aberta. O vínculo é feito no mapa do OrçaCore.
