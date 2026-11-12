# ESTADO DO PLUGIN

> Retrato do estado atual do `plugin_nanocad/`. ✅ = testado sem CAD · 🟡 = implementado, aguardando teste no nanoCAD
> pelo usuário · ⏳ = pendente. Requisitos detalhados: [`REQUISITOS.md`](REQUISITOS.md).

## Versão

`0.8.0` — `src/Nucleo/ObraArquivo.cs:VersaoPlugin` e `src/Versao.props` (e `<Version>` do `.csproj`)

Branch de trabalho: `ccr-7ee2beaa-e39rup`

## Entregáveis

| Arquivo | Situação |
|---|---|
| `appload/nanoCAD26/OrcamentoNanoCAD.dll` | ✅ compila contra a API real do nanoCAD 26 (`hostmgd`/`hostdbmgd` net6.0) · 🟡 carregar no nanoCAD 26 |
| `build_nanocad5.bat` → `appload/nanoCAD5/OrcamentoNanoCAD.dll` | 🟡 núcleo conferido em .NET 4/C# 5; a camada do CAD só compila no Windows com o nanoCAD 5 |
| `exemplos/Projeto_Rua_das_Flores.obra.json` | ✅ importado pela API no teste de contrato |

## Implementado

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

### nanoCAD 5 pronto + projetos e leitura da aplicação (TASK-014/015, 0.7.0) — ✅ núcleo, API e build · 🟡 CAD
- [x] `appload/nanoCAD5/OrcamentoNanoCAD.dll` (.NET 4) gerada aqui contra `src/ApiNanoCAD5` (gerador `ferramentas/GerarApiNanoCAD5`); `Compat` desliga o que faltar
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
| Núcleo (xUnit, net8.0) | `dotnet test tests/Nucleo.Testes` | 51 ✅ (ponta a ponta com `APP_URL`/`APP_EMAIL`/`APP_SENHA`) |
| Núcleo em .NET 4 / C# 5 | `dotnet build tests/Net40` | compila ✅ |
| DLL nanoCAD 26 | `dotnet build src/OrcamentoNanoCAD.csproj -c Release` | compila ✅ (18 comandos, ExtensionApplication, 5 recursos de regras) |
| Contrato com a aplicação | `pytest tests/test_plugin_nanocad_contrato.py` (raiz) | 5 ✅ |
| API do plugin na aplicação | `pytest tests/test_plugin_cad_api.py` (raiz) | 8 ✅ |

## Última atualização

2026-11-11 — TASK-014 (nanoCAD 5 pronto) e TASK-015 (projetos/leitura da aplicação). Antes: TASK-013 (atualização automática), TASK-012 (painel Cabos/Postes, filtro do leitor, ícones), TASK-011, TASK-004 a 010, TASK-003 e TASK-001.
