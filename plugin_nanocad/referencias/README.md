# Referências do resultado esperado

Material de referência do plugin (não é código). Coloque aqui o que ajuda a conferir o plugin:

| Pasta | O quê |
|---|---|
| `mockups/` | mockups conceituais aprovados na ideação (A–G). **Ilustrativos**: não são a interface final nem o nanoCAD real. |
| `desenhos/` | DWG/DXF reais de projetos (ex.: Pernambuco 081) para testar a leitura. |
| `esperado/` | o que o plugin deveria gerar para cada desenho: prints, planilhas, `.obra.json` conferidos à mão. |
| `capturas/` | prints do plugin rodando no nanoCAD (para comparar com os mockups e registrar problemas). |

Mockups atuais:

| Arquivo | Tela |
|---|---|
| `mockups/A_inicio.png` | Ao abrir o plugin: Construir do zero / Ler projeto / Continuar obra salva + regras do projeto |
| `mockups/B_construir.png` | Construir do zero: selecionar postes + anotações e vãos + anotações (P1…, V1…) |
| `mockups/C_ler.png` | Ler projeto: conferência com filtros Ok/Conferir/Erro, Reler/Atualizar/Salvar, vínculos em ciano |
| `mockups/D_vinculos.png` | Ver vínculos: vão, anotação, operação, poste A/estrutura, poste B/estrutura |
| `mockups/E_ativos.png` | Ver ativos com parâmetros: `N4(9400D, P, 3XA35P…, ?)` com pendências em destaque |
| `mockups/F_popup_layer.png` | Pop-up do vão: entrada / saída / ajustes; trocar a operação muda a layer do texto (opção marcada) |
| `mockups/G_persistencia.png` | Onde a obra fica salva (DWG + `.obra.json`) e quem calcula (aplicação / modo autônomo) |

> Desenhos de clientes podem ter dados sensíveis: só versione aqui o que puder ficar no repositório.
