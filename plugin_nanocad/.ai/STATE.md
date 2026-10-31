# ESTADO DO PLUGIN

> Retrato do estado atual do `plugin_nanocad/`. ✅ = testado sem CAD · 🟡 = implementado, aguardando teste no nanoCAD
> pelo usuário · ⏳ = pendente. Requisitos detalhados: [`REQUISITOS.md`](REQUISITOS.md).

## Versão

`0.2.0` — `src/Nucleo/ObraArquivo.cs:VersaoPlugin` (e `<Version>` do `.csproj`)

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
| Núcleo (xUnit, net8.0) | `dotnet test tests/Nucleo.Testes` | 31 ✅ (ponta a ponta com `APP_URL`/`APP_EMAIL`/`APP_SENHA`) |
| Núcleo em .NET 4 / C# 5 | `dotnet build tests/Net40` | compila ✅ |
| DLL nanoCAD 26 | `dotnet build src/OrcamentoNanoCAD.csproj -c Release` | compila ✅ (18 comandos, ExtensionApplication, 5 recursos de regras) |
| Contrato com a aplicação | `pytest tests/test_plugin_nanocad_contrato.py` (raiz) | 5 ✅ |
| API do plugin na aplicação | `pytest tests/test_plugin_cad_api.py` (raiz) | 5 ✅ |

## Última atualização

2026-10-31 — TASK-003 (ligação com a aplicação: regras e base ao vivo). Antes: TASK-001.
