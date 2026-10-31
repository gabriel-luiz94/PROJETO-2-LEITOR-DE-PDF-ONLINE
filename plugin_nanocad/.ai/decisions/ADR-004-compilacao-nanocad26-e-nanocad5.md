# ADR-004 — Compilação: DLL do nanoCAD 26 pronta no repositório; nanoCAD 5 compilado no Windows do usuário

## Contexto

O usuário quer "nessa pasta já obter o arquivo para o APPLOAD" e informou que o nanoCAD 5 suporta .NET. O ambiente de
desenvolvimento é Linux, sem nanoCAD. Verificado: o nanoCAD 26 usa `hostmgd.dll`/`hostdbmgd.dll` para **net6.0** (pacote
NuGet `ModPlus.nanoCAD.API.26`, que traz esses assemblies para compilação). Não há pacote equivalente para o nanoCAD 5
(.NET Framework 4).

## Decisão

- nanoCAD 26: compilar aqui (`net6.0-windows`, x64) contra `ModPlus.nanoCAD.API.26` (só compilação, sem copiar para a
  saída) e **versionar** `appload/nanoCAD26/OrcamentoNanoCAD.dll` (build determinístico).
- WinForms: assemblies de referência de `Microsoft.WindowsDesktop.App.Ref` (o SDK do Linux não tem o "Windows Desktop").
- nanoCAD 5: `build_nanocad5.bat` usa o `csc.exe` do .NET Framework 4 que já vem no Windows + as DLLs do nanoCAD 5
  instalado; código em **C# 5** (`LangVersion=5` garante); se a API do 5 não tiver `TransientManager`, recompila com `SEM_TRANSIENTES`.
- Não usar referências "falsas" (stubs) para gerar a DLL do 5 aqui: um nome/assinatura errado só apareceria ao carregar no CAD.

## Consequências

- A DLL do 26 compila contra a API real (erros de nome/assinatura aparecem na compilação), mas carregar e usar no CAD
  ainda depende do teste do usuário.
- A DLL do 5 é gerada na máquina do usuário; erros de API do 5 chegam como log do `.bat`.

## Evidência no código

`src/OrcamentoNanoCAD.csproj`, `build_nanocad5.bat`, `compilar.sh`, `tests/Net40/Net40.csproj`.

## Data

2026-10-31

## Status

**Atual:** ACEITA
