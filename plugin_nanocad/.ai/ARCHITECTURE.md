# ARQUITETURA DO PLUGIN

> Arquitetura do subproduto `plugin_nanocad/`. A da aplicação está em `../../.ai/ARCHITECTURE.md`.

## 1. Camadas

```
┌──────────────────────────── nanoCAD (26: .NET 6+ · 5: .NET Framework 4) ────────────────────────────┐
│  src/Plugin  (depende de hostmgd.dll / hostdbmgd.dll e WinForms)                                      │
│   Aplicacao.cs        IExtensionApplication, regras disponíveis, Sessao por documento (doc.UserData)  │
│   Comandos.cs         [CommandMethod] ORCAMENTO, ORC_* — todo acesso ao desenho passa por comandos      │
│   ExtratorDesenho.cs  DWG → Desenho (textos/MTEXT por parágrafo, blocos de poste, linhas), cor resolvida│
│   PersistenciaDwg.cs  NOD "ORCAMENTO_LEITOR" → Xrecord "OBRA" (JSON em pedaços) + <desenho>.obra.json  │
│   Marcas.cs           P1/V1 e vínculos como gráficos TRANSITÓRIOS (TransientManager; fallback DrawVector)│
│   Telas.cs, PainelObra.cs  WinForms sem designer: início, painel (não modal), tabelas, pop-up         │
├───────────────────────────────────────────────────────────────────────────────────────────────────────┤
│  src/Nucleo  (SEM CAD — C# 5, .NET 4.0+, sem pacotes)                                                  │
│   Modelo.cs        Desenho* (entrada) · Obra, Poste, Vao, Avulso, Vinculo (saída)                       │
│   MotorLeitor.cs   porte de static/regras_leitor_engine.js                                             │
│   Cores.cs         ACI → RGB, layer TXT_RRGGBB, MTEXT por parágrafo (como services/dxf_service.py)       │
│   Anotacoes.cs     LABEL do cabo, poste compacto (11300D = DT11/300), itens "[qtd ]ativo, …"              │
│   RegrasVinculo.cs porte de services/vinculacao_regras.py (casa, info, pode_vincular, parâmetros)       │
│   LeitorProjeto.cs Ler projeto: textos → redes → vãos → associação → numeração                          │
│   Vinculador.cs    vínculo automático + parâmetros + Conferencia (Ok/Conferir/Erro)                     │
│   Construtor.cs    Construir do zero + ajustes (operação/metros/vínculo manual) + reaplicar no Atualizar │
│   Relatorios.cs    tabelas: conferência, vínculos, ativos com parâmetros                                │
│   ObraArquivo.cs   envelope leitor-obra v1 (CONTRATO) + seção "cad"                                     │
│   RegrasProjeto.cs regras de um projeto (JSON) + catálogo (embutidas / pasta "regras")                  │
│   Json.cs          JSON mínimo (sem System.Text.Json — não existe no .NET 4)                             │
└───────────────────────────────────────────────────────────────────────────────────────────────────────┘
          │ .obra.json (pasta do desenho)                                    ▲ data/*.json (regras, embutidas na DLL)
          ▼                                                                  │
   Aplicação: POST /api/obras/importar → Carregar Obra → orçamento      projeto principal
```

Regra de dependência: **Nucleo não referencia nada do CAD nem de WinForms**. Tudo que é decisão (o que é vão, quem
vincula com quem, estado, o que vai para o arquivo) fica no núcleo e é testado no Linux.

## 2. Ler projeto (LeitorProjeto.Ler)

1. **Textos** ordenados como o leitor de DXF (Y decrescente, depois X; ordenação estável → parágrafos do mesmo MTEXT
   ficam juntos, mesma posição/layer = "mesmo MTEXT" para a vizinhança). Cada um passa por `MotorLeitor.ProcessarEClassificar`.
   Blocos entram como `[BLOCO: nome]` (o leitor de DXF da aplicação também lista).
2. **Redes**: linha nas layers `REDE_MT`/`REDE_BT` cuja cor confirma a layer (vermelho = MT, verde = BT; cor neutra vale
   a layer; divergente descarta). Linhas fora das redes com até 40 × altura típica viram **guias** (linha da anotação).
3. **Vãos**: cada rede é cortada nos postes que estão sobre ela (distância ≤ 1,5 × altura típica). Pedaço entre dois
   postes = vão; ponta sem poste = vão com ponta solta (**Erro**). Uma polilinha pode virar vários vãos (`Trecho`).
4. **Associação**: anotação de CABO → vão mais próximo; de POSTE/ESTRUTURA → poste mais próximo; raio = 4 × altura do
   texto, seguindo a guia quando houver. Segunda anotação no mesmo elemento vira **avulso** (Conferir). Anotação de poste
   sem bloco por perto vira poste "só anotação" (Conferir). Outras entidades (≠ 0, CABO, RAMAIS, IP) vão para Outros como avulsos.
5. **Numeração** (`Numerar`): DFS a partir do poste mais à esquerda; em cada poste segue primeiro o vão que mais continua
   a direção de chegada. Resultado no desenho dos mockups: V1 P1-P2, V2 P2-P3, V3 P3-P4, V4 P2-P5, … (teste `Numeracao_segue_a_rede_como_no_mockup`).
   A ponta A de cada vão é a do poste de número menor.
6. **Vínculo** (`Vinculador.Vincular`): para cada vão I/M (na ordem dos números), em cada ponta, a 1ª estrutura do poste
   (na ordem da anotação) que aceita o cabo pelas regras (`PodeComOcupantes`: regra do cabo × regra da estrutura,
   capacidade, lados iguais, I/M) — até completar as `pontas` do cabo (I: 2, M: 1, ou o que a regra disser). Vínculos manuais ficam.
7. **Parâmetros** por estrutura = os da aplicação (`_analisar_base`): `tipocabo` (VALOR DE CABO; consolidado C > P > A),
   `cabo`/`cabo2` (descrição completa), na ordem cadastrada.
8. **Conferência**: Erro = vão sem poste na ponta; Conferir = sem anotação, cabo sem regra, vínculo incompleto,
   estrutura com menos cabos que o mínimo, cabos diferentes, anotação extra/sem bloco. R/*R: Ok ("R não vincula").

## 3. Contrato — `.obra.json` (ADR-002)

```json
{"formato": "leitor-obra", "versao": 1, "exportado_em": "…", "tipo": "obra", "nome": "Projeto_Rua_das_Flores", "projeto": "081",
 "dados": {
   "cabos":  [{"entidade": "CABO", "operacao": "I", "ativo": "3xA35P(7,94)-13,8kV 45 m", "textoOriginal": "…",
               "_x": 215, "_y": -278, "_geo": [[105,-305],[325,-275]], "_geoTipo": "MT", "vinculoEstruturas": [0, 1]}],
   "outros": [{"entidade": "POSTE", "operacao": "I", "ativo": "9400D, N1, N4", "textoOriginal": "…", "_x": …, "_y": …, "_poste": [325, -275]}],
   "totalizadora": []},
 "cad": {"versao": 1, "origem": "plugin-nanocad", "plugin": "0.2.0", "desenho": "…dwg", "regras": "Pernambuco (081) — Leitor",
         "modo": "ler", "salvo_em": "…", "altura_texto": 8,
         "postes": [{"numero": "P1", "handle": "…", "x": …, "y": …, "texto": "<handle>", "operacao": "I", "operacaoLida": "I",
                     "estruturas": [{"nome": "N1", "cabos": 1, "parametros": {"TIPOCABO": ["P"]}}], "estado": "CONFERIR", …}],
         "vaos": [{"numero": "V1", "linha": "<handle>", "trecho": 0, "pontos": […], "posteA": "P1", "posteB": "P2",
                   "metrosLidos": 45, "metros": 45, "vinculoA": {"poste": "P1", "estrutura": "N1", "k": 0}, …}],
         "avulsos": […]}}
```

- Outros: uma linha por poste com anotação, na ordem P1, P2…; depois os avulsos. Cabos: uma por vão com anotação, na ordem V1, V2…
- `vinculoEstruturas` = ids das unidades de estrutura em Outros: `linha` para a 1ª estrutura da linha, `linha + k/100` para as seguintes (igual a `id_unidade` da aplicação). Ausente em vão R/*R.
- `cad` é ignorada pela aplicação (chave desconhecida); serve para **Continuar obra salva** e para conferência (`estruturas[].parametros` é só informativo — recalculado ao abrir).
- O DWG guarda o MESMO texto (`PersistenciaDwg`).

## 4. Persistência e sessão (ADR-003)

- `Sessao` por documento em `doc.UserData`; o painel é um só e mostra a sessão do documento ativo.
- Gravar: `ORC_SALVAR` → NOD `ORCAMENTO_LEITOR`/`OBRA` (Xrecord, pedaços de 250 caracteres, `DxfCode.Text`) + arquivo ao lado do DWG (UTF-8 sem BOM). Desenho sem arquivo: só no DWG, com aviso.
- Abrir: `ORC_CONTINUAR`/tela inicial → DWG primeiro, senão o arquivo → `ObraArquivo.Ler` (recalcula parâmetros e conferência com as regras atuais).
- Marcas (números e vínculos) são transitórias: não vão para o DWG nem para a obra.

## 4b. Ligação com a aplicação (TASK-003, ADR-005)

```
nanoCAD (thread da tela)                         segundo plano (ThreadPool)              aplicação (FastAPI)
Sincronizador.Tique (timer 1 s) ──────────────▶ ClienteAplicacao ──── HTTP + JWT ────▶ /api/plugin/versao   (impressões digitais)
  ▲  aplica o resultado na obra                   (TransporteHttp)                       /api/plugin/regras   (leitor + vínculo)
  └───────────────────────────────────────────── resultado pronto ◀──────────────────── /api/plugin/conferir (Totalizadora → cálculo)
```
- A cada intervalo (padrão 5 s): versão. `regras` mudou e a sessão usa regras da aplicação → baixa, guarda no cache e refaz
  a obra a partir de `Sessao.UltimoDesenho` (sem tocar no DWG) + `ReaplicarAjustes`. `base`/`conversao`/`regras` mudou → reconfere.
- Cada mudança na obra sobe `Sessao.Revisao`; a conferência vale para a revisão enviada (resposta de revisão velha é descartada).
- `ConferenciaBase.Aplicar`: linha `cabos#i`/`outros#i` → elemento pela ordem de `ObraArquivo.ElementosDasLinhas`; ativos não
  encontrados → `ForaDaBase` + `MensagensBase` → `Conferencia.Avaliar` (Conferir).
- Erros: sem conexão → espera crescente (até 60 s) e segue com o que tem; 401 → para até `ORC_CONECTAR`.
- Guardado em `%APPDATA%\OrcamentoNanoCAD`: `conexao.json` (endereço, e-mail, intervalo; tokens cifrados com DPAPI do usuário)
  e `cache\<hash do endereço>\regras_<projeto>.json`.

## 5. Compilação (ADR-004)

| Alvo | Como | Referências do CAD |
|---|---|---|
| nanoCAD 26 | `dotnet build src/OrcamentoNanoCAD.csproj -c Release` (Linux ou Windows) → `appload/nanoCAD26/` | NuGet `ModPlus.nanoCAD.API.26` (só compilação, `ExcludeAssets=runtime`) |
| nanoCAD 5 | `build_nanocad5.bat` no Windows → `appload/nanoCAD5/` | DLLs do nanoCAD 5 instalado; csc do .NET Framework 4 (C# 5); se faltar `TransientManager`, recompila com `SEM_TRANSIENTES` |
| testes | `dotnet test tests/Nucleo.Testes` (net8.0) · `dotnet build tests/Net40` (net40, C# 5) | — |

- `LangVersion` = 5 no plugin e no núcleo: o mesmo código compila no csc do Windows (nanoCAD 5).
- WinForms/WPF entram como **assemblies de referência** do pacote `Microsoft.WindowsDesktop.App.Ref` (dá para compilar no Linux sem o SDK "Windows Desktop").
- Regras embutidas como recurso, direto de `../data/*.json` (LogicalName `regras.<codigo>.<tabela>.json`).

## 6. Paridade com a aplicação

| Porte | Oráculo | Teste |
|---|---|---|
| `MotorLeitor` ⇄ `static/regras_leitor_engine.js` | `tests/oraculo/oraculo_motor.js` (Node) sobre `casos_motor.json` | `ParidadeTests.Motor_do_leitor_igual_ao_JS` |
| `Anotacoes.LerCabo` ⇄ `anotacao_cabos.py` + `vinculacao_regras.derivar_cabo/descricao_cabo` | `tests/oraculo/oraculo_anotacoes.py` | `ParidadeTests.Anotacao_de_cabo_igual_ao_Python` |
| `Vinculador`/`RegrasVinculo` ⇄ `vinculacao_regras.parametros_por_estrutura/analisar` | a própria aplicação, no pytest | `tests/test_plugin_nanocad_contrato.py` |

`ParidadeTests.Esperados_versionados_batem_com_o_codigo_atual_da_aplicacao` roda os oráculos (se houver node/python) e
falha se o esperado versionado ficou velho — sinal de que o código da aplicação mudou e o porte precisa acompanhar.
