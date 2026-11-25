# ESTADO DO PLUGIN

> Retrato do estado atual do `plugin_nanocad/`. ✅ = testado sem CAD · 🟡 = implementado, aguardando teste no nanoCAD
> pelo usuário · ⏳ = pendente. Requisitos detalhados: [`REQUISITOS.md`](REQUISITOS.md).

## Plugin 0.16.5

TASK-039 (etapa 1): nanoCAD 5 sem a consulta a `Document.CommandInProgress` (abria o «API Logger» e o Salvar da montagem manual não seguia). ✅ 153 testes do núcleo; DLLs 0.16.5; DLL do nanoCAD 5 sem `get_CommandInProgress`. 🟡 Validar no nanoCAD 5. ⏳ Etapa 2: nanoCAD 26 trava no Abrir no OrçaCore / Adicionar anotações.

## Plugin 0.16.4

TASK-038 (ajuste): Selecionar cabo sem «Inteira/Trecho» — polilinha = trecho clicado (vértice a vértice), LINE = inteira; janela com Medido na linha + Comprimento usado (medido arredondado) e metragem por linha do cabo («PB70 32 m», «PB70/32 m»). ✅ 153 testes do núcleo; DLLs 0.16.4. 🟡 Validar no nanoCAD (clique no trecho e janela do cabo).

## Plugin 0.16.3

TASK-038 (ajuste): tipo do cabo pela cor; coletados marcados na tela (ciano/laranja/cinza) com o painel aberto; item do mapa no painel: clique = zoom + destaque com o objeto base; Editar, Excluir, Definir poste/cabo. ✅ 151 testes do núcleo; DLLs 0.16.3. 🟡 Validar no nanoCAD (marcas transitórias e zoom pelo painel).

## Plugin 0.16.2

TASK-038 (ajuste): campos da janela sem ativo dividido — parágrafo de continuação do mesmo MTEXT junta ao campo de cima; texto já juntado pela regra (metragem) não vira campo. ✅ 149 testes do núcleo; DLLs 0.16.2. 🟡 Validar no nanoCAD.

## Plugin 0.16.1

TASK-038 (ajuste): pontos e cabos do mapa listados nas abas do painel; duplo clique/Ajustar reabre a janela (com Excluir). ✅ 138 testes do núcleo; DLLs 0.16.1. 🟡 Validar no nanoCAD.

## Plugin 0.16.0

TASK-038: montagem manual no formato do mapa — Selecionar poste/cabo → textos → janela com campos (L) e operação por campo (F, I, \*I, M, \*M, R, \*R); regras de um projeto (texto fora delas = F) ou sem regras; cabo inteiro ou trecho, metragem anotada com a medida ao lado. ✅ 137 testes do núcleo + contrato com a aplicação; DLLs do nanoCAD 26 e 5 geradas (0.16.0). 🟡 Validar no nanoCAD. [TASK-038](tasks/TASK-038-23-11-2026.md).

## Plugin 0.15.0

TASK-037: 3ª passada liga a anotação de poste que sobrou ao bloco livre mais perto (1:1, até 25 × a altura, Conferir) e o poste sem bloco leva `_posteAprox` no `.obra`. ✅ 130 testes do núcleo; DLLs do nanoCAD 26 e 5 geradas (manifestos 0.15.0 com SHA-256). 🟡 Validar no nanoCAD com a nota real. [TASK-037](tasks/TASK-037-23-11-2026.md).

## Plugin 0.14.0

A montagem manual lê por região, agrupa as anotações de cada poste, mantém as linhas dos vãos separadas e oferece um editor por linha. O envio incremental preserva a obra aberta no Resumo. DLLs geradas para nanoCAD 26 e 5. Validação nativa em Windows pendente. [TASK-036](tasks/TASK-036-22-11-2026.md).

## Plugin 0.13.1

TASK-035 lê o DXF ou DWG ativo em Ler projeto, releitura e diagnóstico. Um DXF salvo e sem alterações é enviado diretamente. O plugin exporta o restante para um DXF temporário da sessão atual e remove a cópia após o envio. 🟡 Validar a exportação e a leitura no nanoCAD 5 e 26 no Windows.

## Plugin 0.13.0

TASK-034 usa o DXF e o motor canônico do OrçaCore na leitura automática. A comparação com a saída do Leitor passou nos testes. A execução no nanoCAD 5 e 26 com a nota real aguarda validação no Windows.

## Plugin 0.12.4

TASK-033 registra `ORCACORE` e `OC` para a tela inicial nos builds do nanoCAD 5 e 26. `ORCAMENTO` permanece disponível como alias. A execução no nanoCAD aguarda confirmação no Windows.

## Plugin 0.12.3

TASK-032 corrige a seleção das regras de Paraíba (027) na leitura do nanoCAD. As regras da aplicação, o cache e a base de fábrica são resolvidos pelo código do projeto. A leitura de um DWG real da Paraíba aguarda validação no nanoCAD 5 e 26.

## Plugin 0.12.2

TASK-031 corrige a abertura do painel de obra antes da seleção da primeira aba. A execução de `ORCAMENTO` no CAD aguarda confirmação.

## Plugin 0.12.0

TASK-029 integra os comandos de notas e a auditoria de layers. O build do 5 usa referências geradas. A execução no 5 e no 26, incluindo atualização após comandos nativos e desfazer, aguarda validação no Windows.

## Plugin 0.11.0

TASK-028 implementa a fila de pendências, revisão por poste, alterações entre leituras, destaque temporário e operação em lote. Validação da interface e dos gráficos no nanoCAD 5 e 26 pendente.

## Distribuição 0.10.4

Build do 5 em x86/.NET 3.5 conferido pelos metadados do binário. A compatibilidade dos comandos com a API instalada segue pendente. O fornecedor distribui a versão 5 mediante login.

## Validação pendente em 19/11/2026

TASK-025 e TASK-026: plugin 0.10.3, N3 removido na tabela Outros e teste de paridade incluindo chamadas múltiplas. A abertura no Windows depende também da correção no OrçaCore (TASK-197). Testar com a nota real nas duas versões do nanoCAD.

## Trabalho de 19/11/2026

TASK-022 (leitura), TASK-023 (alocação por poste) e TASK-024 (painel) implementadas para teste. Conferência operacional pendente no nanoCAD 5 e 26. A comparação de vínculos cobre o exemplo de Pernambuco; regras avançadas de derivação, parâmetros herdados e elementos fictícios exigem casos adicionais.

## Versão

`0.14.0` — `src/Nucleo/ObraArquivo.cs:VersaoPlugin` e `src/Versao.props` (e `<Version>` do `.csproj`)

Branch de trabalho: `codex/plugin-desenho-ativo`

## Entregáveis

| Arquivo | Situação |
|---|---|
| `appload/nanoCAD26/OrcamentoNanoCAD.dll` | ✅ compila contra a API real do nanoCAD 26 (`hostmgd`/`hostdbmgd` net6.0) · 🟡 carregar no nanoCAD 26 |
| `build_nanocad5.bat` → `appload/nanoCAD5/OrcamentoNanoCAD.dll` | ✅ build .NET 3.5 x86/C# 5 com referências geradas · 🟡 teste no host real |
| `exemplos/Projeto_Rua_das_Flores.obra.json` | ✅ importado pela API no teste de contrato |

## Implementado

- [x] TASK-023: o núcleo distribui cabos pelas estruturas de cada poste em ordem de regra. Vínculo manual ocupa a unidade e conserva as pontas escolhidas.

### Núcleo (sem CAD) — ✅
- [x] Motor do leitor portado (`MotorLeitor`), paridade com o JS em 39 casos (Pernambuco 081 + padrão)
- [x] Cor/MTEXT como o leitor de DXF (`Cores`): ACI, layer `TXT_RRGGBB`, parágrafos com cor herdada/local
- [x] Anotação do cabo (`Anotacoes.LerCabo`), paridade com o Python em 13 rótulos; poste compacto → oficial
- [x] Ler projeto: redes por cor + layer, corte em vãos nos postes, guias, associação por raio, avulsos
- [x] Numeração seguindo a rede (mesma ordem dos mockups)
- [x] Vínculo automático pelas regras de vínculo (capacidade, lados iguais, I/M, amarração), vínculo manual
- [x] Parâmetros por estrutura (`tipocabo`, `cabo`, `cabo2`; maior C > P > A) — iguais aos da aplicação
- [x] Conferência Ok / Conferir / Erro com mensagens
- [x] Construir do zero (poste/vão apontados, sugestão das pontas, excluir)
- [x] Ajustes (operação, metros, vínculo) que sobrevivem ao Atualizar; Voltar ao lido; Reler do zero
- [x] `.obra.json` (leitor-obra v1 + seção `cad`), leitura de volta (Continuar)
- [x] Tabelas: conferência, vínculos (com slot), ativos com parâmetros (`?` pendente)

### Camada nanoCAD — 🟡 (compila; falta o teste no CAD)
- [x] Comandos `ORCAMENTO`, `ORC_LER/RELER/ATUALIZAR/CONTINUAR/CONSTRUIR/POSTE/VAO/VINCULAR/EDITAR/ZOOM/VINCULOS/ATIVOS/SALVAR/PAINEL/MARCAS/REGRAS/AJUDA`
- [x] Leitura do DWG: TEXT, MTEXT (por parágrafo), atributos, blocos de poste (nome com "poste", bloco dinâmico pelo nome efetivo), LINE, LWPOLYLINE, POLYLINE 2D, LEADER
- [x] Telas WinForms: início (A), painel de conferência (C), pop-up (F), tabelas (D, E)
- [x] Marcas transitórias (P1/V1 com moldura, tracejado ciano do vínculo, destaque do elemento)
- [x] Gravar/ler a obra no DWG (NOD) e em `<desenho>.obra.json`
- [x] Trocar a layer da anotação ao mudar a operação (opção do pop-up; layer tirada das regras)
- [x] Regras embutidas (081 e padrão) + pasta `regras\<codigo>\` ao lado da DLL

### "+ Cabo": cabo com comprimento da linha ou da vizinhança (TASK-021, 0.10.1) — ✅ núcleo · 🟡 CAD

### Projeto de trabalho do OrçaCore (TASK-020, 0.10.0) — ✅ núcleo, aplicação e contrato · 🟡 CAD/Windows
- [x] Salvar grava `<desenho>.obra` (formato do 📁 Projeto), com aviso se as linhas foram alteradas no OrçaCore; número de obra (`ORC_PROJETO`)
- [x] 🌐 entrega o caminho ao OrçaCore (aberto ou fechado); sem gravação no banco
- [ ] Testar no Windows: salvar no nanoCAD, 🌐 com o OrçaCore aberto e fechado, editar no programa e salvar de novo no plugin (aviso)

### Cabo lido de texto: comprimento da linha de baixo sem vírgula (TASK-019, 0.9.2) — ✅ núcleo · 🟡 CAD

### Linha nova lida de um texto do desenho (TASK-018, 0.9.1) — ✅ núcleo · 🟡 CAD
- [x] "Selecionar texto no desenho (TEXT/MTEXT)…" na tela de + Cabo / + Outros; regras do leitor preenchem ativo, operação e tabela

### Edição no plugin (TASK-017, 0.9.0) — ✅ núcleo e API · 🟡 CAD
- [x] Ativo editável com o autocompletar da aplicação; Excluir/Restaurar qualquer linha; + Cabo / + Outros sem desenho
- [ ] Testar no nanoCAD: digitar no campo do ativo (lista, ↑/↓, Enter, "(" mostra os parâmetros), excluir e restaurar, + Cabo

### "Abrir na aplicação" no GridReader deste computador (TASK-016, 0.8.0) — ✅ · 🟡 Windows

### nanoCAD 5 pronto + projetos e leitura da aplicação (TASK-014/015, 0.7.0) — ✅ núcleo, API e build · 🟡 CAD
- [x] `appload/nanoCAD5/OrcamentoNanoCAD.dll` (.NET 3.5 x86) gerada aqui contra `src/ApiNanoCAD5` (gerador `ferramentas/GerarApiNanoCAD5`); `Compat` desliga o que faltar
- [x] Tela inicial com todos os projetos da aplicação; regras sempre do projeto; configuração de leitura na aplicação (vale em qualquer computador)
- [x] Programa principal traz e instala as duas DLLs (TASK-131 do principal)
- [ ] Testar no nanoCAD 5 de verdade: carregar a DLL, `ORC_AJUDA` mostra os recursos indisponíveis; mandar o log se não carregar

### Atualização automática (TASK-013, 0.6.0) — ✅ núcleo e API · 🟡 CAD
- [x] Consulta/baixa/confere/troca (`Nucleo/Atualizacao`, testes `AtualizacaoTests`); `AtualizacaoPlugin` (timer, auxiliar PowerShell, estado em `atualizacao.json`)
- [x] `ORC_ATUALIZAR_PLUGIN`, `ORC_VOLTAR_VERSAO`; opção em `ORC_CONECTAR`; manifesto `appload/nanoCAD26/versao.json` (compilar.sh)
- [ ] Testar no nanoCAD: instalar a 0.6.0 à mão uma vez; publicar uma 0.6.1 e ver a troca ao reabrir

### Painel em duas listas + ícones (TASK-012, 0.5.0) — ✅ núcleo · 🟡 CAD
- [x] CABOS (vãos) e POSTES E ESTRUTURAS (outros) separados, como a Tabela de Cabos e a Tabela de Outros do Resumo (`Relatorios.Cabos/Outros`)
- [x] Ler: só o que as regras do leitor classificaram (sem anotação = oculto, também nas marcas e nos vínculos); Construir/manual: tudo (`Relatorios.Visivel`)
- [x] Ícones desenhados (GDI+, sem depender de emoji): estado, rede MT/BT, operação, poste/solto; barra de ícones com dica (`Plugin/Icones.cs`)

### Desenho real NOTA 9101578729 (TASK-011, 0.4.0) — ✅ núcleo · 🟡 CAD
- [x] Diagnóstico real em `referencias/desenhos/` como teste (`LeitorRealTests`)
- [x] Guias em pedaços / mesma layer; irmãos (várias anotações por vão/poste); MT×BT; redes em pedaços; troca de poste; placas
- [x] Sugestões de layers de rede (linhas vermelhas/verdes de poste a poste) e de blocos de poste (pontas soltas) no ORC_LER e no diagnóstico

### Correções do primeiro uso (TASK-004 a 010, 0.3.0) — ✅ núcleo e aplicação · 🟡 CAD
- [x] Desenho bruto + montagem no núcleo; paridade texto a texto com o Leitor da aplicação num DXF de teste (`LeitorDxfTests`)
- [x] MLEADER (texto + linhas de chamada), LEADER, contorno do texto; guia usada mostrada no destaque
- [x] Configurar leitura por projeto (`ORC_LEITURA`): layers da rede, blocos/layers de poste (bloco ou círculo), layers de guia, raios
- [x] Ignorados no painel; `ORC_DIAGNOSTICO` (`<desenho>.diagnostico.json` + tabela); log em `%APPDATA%\OrcamentoNanoCAD\log.txt`
- [x] Botões do painel com Enter + foco no desenho; seleção manual: bloco ou círculo, TEXT/MTEXT/MLEADER, linha/polilinha 2D/3D
- [x] ⌂ Início, ✕ Fechar obra (`ORC_FECHAR`), voltar ao painel; tema da aplicação em todas as telas
- [x] 🌐 Abrir na aplicação (`ORC_ABRIR_APP`) + `/api/plugin/obra` + Resumo abre `?obra=` + login volta ao link

### Ligação com a aplicação (TASK-003) — ✅ núcleo e API · 🟡 CAD
- [x] `ORC_CONECTAR` (endereço, e-mail, senha → token; DPAPI), `ORC_SINCRONIZAR`, situação no painel e na tela inicial
- [x] Regras do leitor/vínculo baixadas da aplicação ("Aplicação: PERNAMBUCO (081)"), com cache em `%APPDATA%\OrcamentoNanoCAD`
- [x] Sincronização automática: versão consultada a cada 5 s (ajustável); regras mudaram → obra refeita do último desenho lido, ajustes mantidos
- [x] Conferência com a base de orçamento (`/api/plugin/conferir`): ativo fora da base → Conferir ("não está na base", "fica fora do orçamento")
- [x] Ponta a ponta conferido contra a aplicação rodando (uvicorn): login, versão, regras, conferência, renovação do token

### Aplicação (TASK-103 do projeto principal) — ✅
- [x] `services/obras_arquivo.py`: importar/exportar mantém `vinculoEstruturas`, `textoOriginal`, `_x/_y/_geo/_geoTipo/_poste`, fictícios…
- [x] `tests/test_plugin_nanocad_contrato.py`: importa o exemplo pela API e compara o vínculo/parâmetros com os da aplicação

## Pendente / próximas tarefas

- [ ] NOTA 9101578729 no nanoCAD com a configuração sugerida (⚙ Leitura: MT = REDE_MT, ANOTACOES, =0; postes = pg, KMLP) e novo diagnóstico
- [ ] Testar no nanoCAD as TASK-004 a 010 (painel, seleção manual, guias, Abrir na aplicação, visual)
- [ ] Teste da ligação com a aplicação dentro do nanoCAD (TASK-003, passos na tarefa)
- [ ] **Teste no nanoCAD 26** pelo usuário: APPLOAD, ORCAMENTO, ler um DWG real de Pernambuco, salvar, importar na aplicação (TASK-002)
- [ ] **nanoCAD 5**: rodar `build_nanocad5.bat` e mandar os logs se falhar
- [ ] Desenhos reais em `referencias/desenhos/` para calibrar raios/tolerâncias (hoje: 4×, 1,5×, 40× a altura típica)
- [ ] Fictícios (`cabosFicticios`, `estruturasFicticias`) no plugin (RQ-25)
- [ ] Textos dentro de blocos aninhados e redes dentro de blocos (a aplicação varre até 3 níveis; o plugin só o espaço do modelo e atributos)
- [ ] Notação antiga de Outros (`DT11/600 1-N1`) e apelidos da base no vínculo (o plugin entende a notação compacta `9400D, N1, N4`)
- [ ] Modo autônomo + motor de cálculo gravando o resultado na pasta do desenho (RQ-40/41) — tarefa da aplicação
- [ ] Pontos em aberto do briefing: `REQUISITOS.md` › Pontos em aberto

## Limitações conhecidas

- O motor do leitor e parte do vínculo são **portes** (C#) do JS/Python da aplicação: mudança lá exige repetir aqui (`RULES.md` P3; o teste de paridade avisa).
- `\d` no motor: a regex em C# roda em modo ECMAScript (só dígitos ASCII, como o JS).
- Ponto de poste na rede: o bloco precisa estar a até 1,5 × altura típica da linha; postes desenhados longe da rede não cortam o vão.
- Ligação com a aplicação: "ao vivo" = consulta a cada N segundos (padrão 5); no modo online cada consulta lê regras/base do Supabase.
  Mudança no DESENHO continua pedindo Atualizar (o plugin não relê o DWG sozinho).
- Sem o SDK/DLLs do nanoCAD 5, a camada do CAD do 5 não foi compilada aqui; se a API do 5 não tiver algo, o `.bat` tenta sem gráficos transitórios.

## Testes

| Suíte | Comando | Resultado (31/10/2026) |
|---|---|---|
| Núcleo (xUnit, net8.0) | `dotnet test tests/Nucleo.Testes` | 91 ✅ (ponta a ponta com `APP_URL`/`APP_EMAIL`/`APP_SENHA`) |
| Núcleo em .NET 4 / C# 5 | `dotnet build tests/Net40` | compila ✅ |
| DLL nanoCAD 26 | `dotnet build src/OrcamentoNanoCAD.csproj -c Release` | compila ✅ (18 comandos, ExtensionApplication, 5 recursos de regras) |
| Contrato com a aplicação | `pytest tests/test_plugin_nanocad_contrato.py` (raiz) | 5 ✅ |
| API do plugin na aplicação | `pytest tests/test_plugin_cad_api.py` (raiz) | 8 ✅ |

## Última atualização

2026-11-11 — TASK-014 (nanoCAD 5 pronto) e TASK-015 (projetos/leitura da aplicação). Antes: TASK-013 (atualização automática), TASK-012 (painel Cabos/Postes, filtro do leitor, ícones), TASK-011, TASK-004 a 010, TASK-003 e TASK-001.
