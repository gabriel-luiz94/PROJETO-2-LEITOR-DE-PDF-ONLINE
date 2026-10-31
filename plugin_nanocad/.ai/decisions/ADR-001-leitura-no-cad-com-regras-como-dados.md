# ADR-001 — Leitura no CAD, em C#, com as regras como dados e paridade por oráculo

## Contexto

O briefing (§5, ponto em aberto 1) pedia decidir onde roda a lógica de leitura: (a) no servidor, (b) motor JS embutido
reaproveitando `static/*.js`, (c) reescrita em C#, preferindo "uma só fonte de regras". O usuário autorizou implementar.

## Problema

O leitor precisa ser **instantâneo e dentro do desenho** (numerar, mostrar vínculos, pop-up) — inclusive no nanoCAD 5
(.NET Framework 4) e sem depender da aplicação estar aberta.

## Alternativas consideradas

### (a) Servidor — o plugin manda a geometria e recebe a obra
- Prós: nenhuma cópia de lógica.
- Contras: exige a aplicação rodando e login (JWT) dentro do CAD; o motor do leitor é JS (ADR-003 do principal: motor no frontend) — não existe hoje um endpoint que faça leitor + vínculo; latência a cada ajuste.

### (b) Motor JS embutido
- Prós: roda o mesmo JS.
- Contras: .NET Framework 4 não tem motor JS moderno (JScript.NET é ES3; Chakra/JScript9 é ES5); os `static/*.js` usam ES6+. Pacotes como Jint exigem .NET mais novo/NuGet — inviável no nanoCAD 5 compilado pelo csc do Windows.

### (c) Porte em C# do MOTOR, com as REGRAS como dados
- Prós: roda em qualquer nanoCAD .NET, offline, instantâneo; as regras (o que muda no dia a dia) continuam numa só fonte (`data/*.json`, embutidos na DLL ou exportados da aplicação).
- Contras: terceira cópia do motor (JS, Python, C#).

## Decisão

(c), com **paridade automatizada**: oráculos que rodam o código da própria aplicação (Node para `regras_leitor_engine.js`,
Python para `anotacao_cabos.py`/`vinculacao_regras.py`) geram o esperado; o xUnit compara o porte; um teste avisa quando
o esperado versionado ficou velho; o pytest de contrato compara o vínculo do plugin com o da aplicação.

## Consequências

### Positivas
- Plugin independente da aplicação em tempo de uso; mesmo resultado do leitor da aplicação.

### Negativas / custos aceitos
- Mudança no motor/anotações/vínculo da aplicação exige repetir no C# (`RULES.md` P3).

### Implicações para implementações futuras
- Não copiar regras para o C#: sempre ler o JSON.
- Se um dia houver endpoint de leitura no servidor, o plugin pode usá-lo como alternativa (a camada `Nucleo` já isola a leitura).

## Evidência no código

`src/Nucleo/MotorLeitor.cs`, `Anotacoes.cs`, `RegrasVinculo.cs`; `tests/oraculo/`; `tests/Nucleo.Testes/ParidadeTests.cs`;
`../../tests/test_plugin_nanocad_contrato.py`.

## Data

2026-10-31

## Status

**Atual:** ACEITA
