# REGRAS DE TRABALHO — PLUGIN nanoCAD

> Valem **todas** as regras do projeto principal (`../../.ai/RULES.md`: entender antes de alterar, não inventar
> comportamento, alteração mínima, testes, documentação, segurança, português…). Abaixo, só o que é específico do plugin.

## Regra P1 — O núcleo não conhece o CAD

`src/Nucleo/` não usa `HostMgd.*`, `Teigha.*`, WinForms nem pacotes NuGet. Decisões (o que é vão, quem vincula com quem,
estado, o que vai para o arquivo) ficam no núcleo e têm teste. A camada `src/Plugin/` só lê o desenho, mostra e grava.

## Regra P2 — C# 5 e .NET Framework 4.0

Todo código de `src/` compila com o csc do .NET Framework 4 (nanoCAD 5): **sem** interpolação `$""`, `?.`, `nameof`,
membros `=>`, inicializador de auto-propriedade, `out var`, tuplas, `using static`, pattern matching. O `.csproj` fixa
`LangVersion=5`; `tests/Net40` confere o núcleo em `net40`. Testes (`tests/`) podem usar C# moderno.

## Regra P3 — Paridade com a aplicação (porte do motor)

As REGRAS são dados (`../data/*.json`) e nunca são copiadas. O MOTOR é portado; ao mudar na aplicação qualquer um de
`static/regras_leitor_engine.js`, `services/anotacao_cabos.py`, `services/vinculacao_regras.py` (e pares JS):
1. regerar os esperados (`node tests/oraculo/oraculo_motor.js > tests/oraculo/esperado_motor.json`,
   `python tests/oraculo/oraculo_anotacoes.py > tests/oraculo/esperado_anotacoes.json`, a partir da raiz do repo);
2. ajustar o porte até `dotnet test` passar;
3. rodar `pytest tests/test_plugin_nanocad_contrato.py` (compara o vínculo do plugin com o da aplicação).
`ParidadeTests.Esperados_versionados_batem_com_o_codigo_atual_da_aplicacao` falha quando o passo 1 foi esquecido.

## Regra P4 — Contrato do `.obra.json` (ADR-002)

- O arquivo é o envelope `leitor-obra` versão 1 da aplicação (`services/obras_arquivo.py`). Mudança incompatível = nova
  versão nos dois lados, nunca silenciosa.
- Campo novo numa linha de Cabos/Outros só vale se a aplicação também o aceitar em `CAMPOS_VINCULO` — senão a
  importação descarta. Ao mexer, atualizar o exemplo (`ATUALIZAR_EXEMPLOS=1 ./compilar.sh`) e o teste de contrato.
- A seção `cad` é do plugin: pode crescer (mantendo leitura do que já foi salvo — `VersaoCad`).

## Regra P5 — O desenho do usuário

- O plugin só **altera** o DWG em dois casos: gravar a obra no dicionário `ORCAMENTO_LEITOR` (ORC_SALVAR) e trocar a
  layer da anotação quando o usuário muda a operação **com a opção marcada** no pop-up. Nada mais (sem camada de marcas).
- Marcas na tela = gráficos transitórios.
- Excluir da obra não apaga nada do desenho.

## Regra P6 — A DLL versionada

`appload/nanoCAD26/OrcamentoNanoCAD.dll` é versionada (o usuário usa direto). Toda mudança em `src/` que for para o
repositório vai junto com a DLL recompilada (`./compilar.sh`, build determinístico). `appload/nanoCAD5/` é gerada no
Windows do usuário e **não** é versionada pelo agente.

## Regra P7 — Testes no CAD são do usuário

O ambiente de desenvolvimento não roda o nanoCAD. Tudo que só dá para conferir no CAD fica marcado 🟡 em
`REQUISITOS.md`/`STATE.md` até o usuário confirmar; erros que ele relatar (mensagens da linha de comando, logs do
`build_nanocad5.bat`) viram tarefa. Entregar em passos pequenos e testáveis.

## Regra P8 — Documentação do plugin

Ao concluir uma tarefa do plugin: `tasks/TASK-NNN-dd-mm-aaaa.md` (numeração própria do plugin), `CHANGELOG.md`,
`STATE.md` e, se mudar requisito, `REQUISITOS.md`. Se a tarefa tocar a aplicação, ela também ganha TASK no `.ai` principal.
