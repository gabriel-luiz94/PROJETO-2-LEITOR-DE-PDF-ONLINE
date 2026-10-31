# CONTEXTO DO SUBPRODUTO — Plugin nanoCAD

> Contexto do plugin. O contexto geral do projeto (aplicação, regras de negócio RN-xx, glossário) está em `../../.ai/`
> — **não repita aqui**; referencie (`../../.ai/CONTEXT.md §7 RN-01`).
> Briefing original da ideação: [`BRIEFING.md`](BRIEFING.md). Requisitos numerados: [`REQUISITOS.md`](REQUISITOS.md).

## 1. O que é

Plugin **C# (.NET)** para o **nanoCAD 26** (e **nanoCAD 5**) que roda dentro do desenho e produz uma **obra totalmente
parametrizada** — o arquivo `<desenho>.obra.json` — para a aplicação online (Carregar Obra) ou o modo autônomo
calcularem o orçamento. O cálculo **não** roda no CAD (por ora).

Substitui, para quem desenha no nanoCAD, o caminho "exportar DXF → Leitor da aplicação → Resumo → vínculo no mapa":
o mesmo trabalho é feito no próprio desenho, com os números de cada poste/vão visíveis e a conferência ao lado.

## 2. Relação com a aplicação principal

| Assunto | Na aplicação | No plugin |
|---|---|---|
| Regras do leitor (layer + cor → operação/ativo/entidade) | `static/regras_leitor_engine.js` | `src/Nucleo/MotorLeitor.cs` — **porte**, mesmas regras (JSON de `data/`), paridade testada contra o JS |
| Anotação do cabo (LABEL) | `services/anotacao_cabos.py` ⇄ `static/anotacao_cabos.js` | `src/Nucleo/Anotacoes.cs` — porte parcial (código, comprimento, fases, tensão, descrição), paridade contra o Python |
| Regras de vínculo cabo ⇄ estrutura | `services/vinculacao_regras.py` ⇄ `static/vinculacao_regras.js` | `src/Nucleo/RegrasVinculo.cs` — porte (casa, info_cabo/estrutura, pode_vincular, parâmetros) |
| Geometria do DXF (redes, postes, guias) | `services/dxf_geometria.py` | `src/Nucleo/LeitorProjeto.cs` (mesmos critérios: cor decide, layer confirma; raio 4 × altura do texto) |
| Formato do arquivo de obra | `services/obras_arquivo.py` (`leitor-obra` v1) | `src/Nucleo/ObraArquivo.cs` — gera esse formato + seção `cad` |
| Cálculo do orçamento | `services/orcamento_calc.py` | — (fora do CAD) |

Decisão sobre onde roda a leitura: [ADR-001](decisions/ADR-001-leitura-no-cad-com-regras-como-dados.md).
Contrato do arquivo: [ADR-002](decisions/ADR-002-contrato-obra-json.md).

## 3. Fluxo

1. `ORCAMENTO` → tela inicial: **Construir do zero**, **Ler projeto**, **Continuar obra salva**; escolha das regras (ex.: Pernambuco 081) e do código do projeto na aplicação.
2. **Ler projeto**: o DWG vira `Desenho` (textos, blocos de poste, linhas) → motor do leitor → redes cortadas em vãos nos postes → anotações associadas → numeração P1…/V1… seguindo a rede → vínculo automático → conferência (Ok/Conferir/Erro).
3. **Construir do zero**: o usuário aponta poste (bloco + anotação) e vão (linha + anotação); o plugin sugere os postes das pontas.
4. **Conferência e ajustes** no painel e no pop-up (operação, metros, estrutura do vínculo; opcionalmente muda a layer do texto no desenho).
5. **Salvar**: a obra vai para o DWG (dicionário `ORCAMENTO_LEITOR`) e para `<desenho>.obra.json` na pasta do desenho. **Reler** = do zero; **Atualizar** = relê mantendo os ajustes.
6. Na aplicação: projeto selecionado → Carregar Obra → importar o `.obra.json` → orçamento.

## 4. Versões do nanoCAD

| Versão | Runtime | Entrega |
|---|---|---|
| nanoCAD 26 | .NET 6+ (assemblies `hostmgd`/`hostdbmgd` net6.0) | DLL pronta em `appload/nanoCAD26/` (compilada contra `ModPlus.nanoCAD.API.26`) |
| nanoCAD 5 | .NET Framework 4 | `build_nanocad5.bat` no Windows do usuário (csc do Windows, C# 5, DLLs do nanoCAD 5 instalado) |

O usuário confirmou que o nanoCAD 5 aceita plugin .NET. A API do 5 **não** foi verificada aqui (sem SDK/DLLs do 5); o
núcleo é conferido em .NET 4/C# 5 (`tests/Net40`), a camada do CAD só no Windows do usuário.

## 5. Ambiente

- Desenvolvimento em Linux/nuvem (sem nanoCAD): testes do núcleo (`dotnet test`) e contrato (`pytest`).
- Teste no CAD: com o usuário, no Windows — entregar passos pequenos e testáveis.
- Idioma: português em código de domínio, mensagens, documentação e commits (como o projeto principal).
