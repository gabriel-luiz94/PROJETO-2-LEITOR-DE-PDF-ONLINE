# Plugin nanoCAD — Orçamento (Leitor de Projetos)

Subproduto do **Leitor de Projetos Pro**: um plugin C# (.NET) que roda **dentro do nanoCAD**, lê o desenho da rede
(postes, vãos e anotações) e produz uma **obra totalmente parametrizada** (`<desenho>.obra.json`), que a aplicação
importa em **Carregar Obra** (`POST /api/obras/importar`) para gerar o orçamento. O cálculo **não** roda no CAD.

> Documentação para IA/desenvolvimento: [`.ai/`](.ai/) — requisitos, estado, arquitetura, regras e tarefas do plugin.
> Mockups do resultado esperado: [`referencias/mockups/`](referencias/mockups/).

## Usar no nanoCAD

| nanoCAD | Arquivo para o APPLOAD | Como obter |
|---|---|---|
| **26** (.NET 6+) | [`appload/nanoCAD26/OrcamentoNanoCAD.dll`](appload/nanoCAD26/) | já vem pronto nesta pasta |
| **5** (.NET Framework 4) | `appload/nanoCAD5/OrcamentoNanoCAD.dll` | rode `build_nanocad5.bat` no Windows com o nanoCAD 5 instalado (usa o compilador que já vem no Windows e as DLLs do próprio nanoCAD 5) |

1. No nanoCAD: `APPLOAD` (ou `NETLOAD`) → escolha a DLL. Para carregar sempre, adicione-a à lista de inicialização do APPLOAD.
2. Comando **`ORCAMENTO`** → *Construir do zero*, *Ler projeto* ou *Continuar obra salva*; escolha as regras (ex.: Pernambuco 081).
3. Confira no painel (filtros Ok / Conferir / Erro), ajuste no pop-up (duplo clique ou `ORC_EDITAR`) e **Salvar**.
4. Na aplicação: selecione o mesmo projeto (ex.: 081) → Carregar Obra → importar o `.obra.json` da pasta do desenho.

### Comandos

| Comando | O que faz |
|---|---|
| `ORCAMENTO` | tela inicial |
| `ORC_LER` / `ORC_RELER` / `ORC_ATUALIZAR` | lê o desenho pelas regras (layer + cor) · reler do zero (descarta ajustes) · reler mantendo os ajustes |
| `ORC_CONSTRUIR`, `ORC_POSTE`, `ORC_VAO`, `ORC_VINCULAR` | construir do zero: apontar poste (bloco + anotação) e vão (linha + anotação); refazer vínculos automáticos |
| `ORC_EDITAR` | pop-up do vão/poste: entrada (lida do desenho), saída (para a obra), ajustes (operação, metros, estrutura do vínculo; troca opcional da layer do texto) |
| `ORC_VINCULOS`, `ORC_ATIVOS` | tabela de vínculos · ativos com parâmetros (`N4(9400D, P, 3XA35P_7.94_-13.8KV, ?)`) |
| `ORC_SALVAR` | grava a obra **no DWG** e em **`<desenho>.obra.json`** na pasta do desenho |
| `ORC_CONTINUAR`, `ORC_PAINEL`, `ORC_ZOOM`, `ORC_MARCAS`, `ORC_REGRAS`, `ORC_AJUDA` | abrir a obra salva · reabrir o painel · localizar · ligar/desligar números e vínculos na tela · recarregar regras · ajuda |

### Regras

As regras do leitor e do vínculo são **os mesmos arquivos** de `../data/` do projeto principal, embutidos na DLL
(Pernambuco 081 e o padrão genérico). Para usar outras, crie ao lado da DLL a pasta `regras\<codigo>\` com
`processamento.json`, `classificacao.json`, `vinculacao.json` (exportados da aplicação) e, opcionalmente, `nome.txt`;
depois `ORC_REGRAS`.

## Pastas

```
plugin_nanocad/
├── .ai/                  requisitos, estado, arquitetura, regras, decisões e tarefas DO PLUGIN
├── appload/nanoCAD26/    DLL pronta para o APPLOAD (nanoCAD 26)
├── appload/nanoCAD5/     saída do build_nanocad5.bat
├── exemplos/             .obra.json gerado pelo núcleo (é o que o teste de contrato da aplicação importa)
├── referencias/          mockups e arquivos do resultado esperado (DWG/DXF, imagens…)
├── src/Nucleo/           lógica SEM CAD (leitura, regras, vínculo, .obra.json) — testável no Linux
├── src/Plugin/           camada do nanoCAD (comandos, telas, DWG, marcas na tela)
├── tests/Nucleo.Testes/  testes xUnit do núcleo (inclui paridade com o JS/Python da aplicação)
├── tests/Net40/          confere que o núcleo compila como o nanoCAD 5 exige (.NET 4 / C# 5)
├── tests/oraculo/        oráculos: rodam o código da APLICAÇÃO para gerar o esperado
├── build_nanocad5.bat    compila para o nanoCAD 5 no Windows
├── compilar_nanocad26.bat recompila a DLL do nanoCAD 26 (precisa do .NET SDK)
└── compilar.sh           desenvolvimento: testes + verificação .NET 4 + DLL do nanoCAD 26
```

## Desenvolvimento

```bash
cd plugin_nanocad
./compilar.sh                         # testes do núcleo + net40 + appload/nanoCAD26/OrcamentoNanoCAD.dll
ATUALIZAR_EXEMPLOS=1 ./compilar.sh    # também regrava exemplos/Projeto_Rua_das_Flores.obra.json
cd .. && python -m pytest tests/test_plugin_nanocad_contrato.py   # contrato com a aplicação
```

Precisa do .NET SDK 8. A API do nanoCAD 26 vem do pacote NuGet `ModPlus.nanoCAD.API.26` (só para compilar).
O ambiente de desenvolvimento não roda o nanoCAD: o teste dentro do CAD é feito pelo usuário, no Windows.
