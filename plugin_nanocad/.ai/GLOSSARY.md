# GLOSSÁRIO DO PLUGIN

> Termos do plugin. Os do domínio (ATIVO, COMPONENTE, operação I/R/M, Totalizadora…) estão em `../../.ai/GLOSSARY.md`.

| Termo | Significado |
|---|---|
| **Poste** (P1…) | bloco cujo nome contém "poste" + a anotação associada (`9400D, N1, N4`). Sem bloco por perto, a anotação vira um poste "só anotação" (Conferir). |
| **Vão** (V1…) | **linha da rede + anotação associada**. Uma rede (`REDE_MT`/`REDE_BT`) é cortada nos postes: cada pedaço entre dois postes é um vão. |
| **Trecho** | índice do vão dentro da mesma linha (polilinha que passa por vários postes vira vários vãos). |
| **Ponta A / B** | postes nas pontas do vão; A = o de número menor (por onde a numeração chegou). Ponta sem poste = Erro. |
| **Anotação** | o texto do desenho associado ao poste/vão (TEXT, parágrafo de MTEXT ou atributo). |
| **Avulso** (A1…) | texto lido que não ficou num poste/vão (segunda anotação, cabo sem vão, outras entidades). Vai para a obra como está. |
| **Estrutura / unidade** | cada item não-poste da anotação do poste (`N1`, `N4`); posição `k` na anotação. Id na obra = linha de Outros (+ k/100 a partir da 2ª). |
| **Vínculo** | ligação ponta do vão ⇄ estrutura do poste daquela ponta. Automático (regras) ou manual (✋, pop-up). Só I/M (e `*I`/`*M`). |
| **Slot** | posição do cabo entre os cabos de uma estrutura ("N1 (slot 2/2)"). |
| **Parâmetros** | o que a estrutura recebe dos cabos ligados: `tipocabo` (VALOR DE CABO, ex.: P), `cabo`/`cabo2` (descrição completa, ex.: `3XA35P_7.94_-13.8KV`). `?` = pendente. |
| **Guia** | linha curta fora das redes que liga a anotação à rede/poste (como na leitura de DXF). |
| **Altura típica** | mediana da altura dos textos; escala de todos os raios (associação 4×, poste na rede 1,5×, guia 40×) e das marcas. |
| **Marcas** | números e tracejados ciano desenhados como gráficos transitórios (não gravam no DWG). |
| **Regras do projeto** | processamento + classificação do leitor + regras de vínculo de um projeto (ex.: Pernambuco 081) — JSON da aplicação. |
| **Sessão** | estado do plugin num desenho aberto: obra, regras, marcas. |
| **Seção `cad`** | parte do `.obra.json` que só o plugin usa (números, handles, ajustes) — a aplicação ignora. |
| **Reler / Atualizar** | reler do zero (descarta ajustes) / reler mantendo os ajustes (casando por handle). |
| **Ajuste** | mudança feita no pop-up: operação, metros, estrutura do vínculo. "Voltar ao lido" desfaz. |
