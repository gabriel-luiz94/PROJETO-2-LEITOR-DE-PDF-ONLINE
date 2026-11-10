#!/usr/bin/env bash
# Desenvolvimento (Linux/macOS/Windows com bash): testes do núcleo + verificação .NET 4/C# 5 (nanoCAD 5) + DLL do nanoCAD 26.
#   ./compilar.sh            testa e gera appload/nanoCAD26/OrcamentoNanoCAD.dll
#   ATUALIZAR_EXEMPLOS=1 ./compilar.sh   também regrava exemplos/*.obra.json
set -euo pipefail
cd "$(dirname "$0")"
dotnet test tests/Nucleo.Testes -c Release --nologo
dotnet build tests/Net40 -c Release --nologo -v q
rm -rf src/bin src/obj
dotnet build src/OrcamentoNanoCAD.csproj -c Release --nologo -v q
# manifesto da atualização automática (a aplicação serve a DLL + versão + SHA-256 em /api/plugin/atualizacao)
VERSAO=$(sed -n 's:.*<Version>\(.*\)</Version>.*:\1:p' src/OrcamentoNanoCAD.csproj)
SHA=$(sha256sum appload/nanoCAD26/OrcamentoNanoCAD.dll | cut -d' ' -f1)
printf '{\n  "versao": "%s",\n  "sha256": "%s",\n  "arquivo": "OrcamentoNanoCAD.dll"\n}\n' "$VERSAO" "$SHA" > appload/nanoCAD26/versao.json
ls -l appload/nanoCAD26/
cat appload/nanoCAD26/versao.json
