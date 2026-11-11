#!/usr/bin/env bash
# Desenvolvimento (Linux/macOS/Windows com bash): testes do núcleo + verificação .NET 4/C# 5 (nanoCAD 5) + DLL do nanoCAD 26.
#   ./compilar.sh            testa e gera appload/nanoCAD26/OrcamentoNanoCAD.dll
#   ATUALIZAR_EXEMPLOS=1 ./compilar.sh   também regrava exemplos/*.obra.json
set -euo pipefail
cd "$(dirname "$0")"
dotnet run --project ferramentas/GerarApiNanoCAD5 -c Release -v q -- .   # referências do nanoCAD 5 = o que o plugin usa da API
dotnet test tests/Nucleo.Testes -c Release --nologo
dotnet build tests/Net40 -c Release --nologo -v q
rm -rf src/bin src/obj
dotnet build src/OrcamentoNanoCAD.csproj -c Release --nologo -v q
# nanoCAD 5 (.NET 4): mesmo código, contra as referências geradas em src/ApiNanoCAD5 (ferramentas/GerarApiNanoCAD5)
rm -rf src/nanoCAD5/bin src/nanoCAD5/obj
dotnet build src/nanoCAD5/OrcamentoNanoCAD5.csproj -c Release --nologo -v q
# manifestos da atualização automática (a aplicação serve DLL + versão + SHA-256 em /api/plugin/atualizacao)
VERSAO=$(sed -n 's:.*<Version>\(.*\)</Version>.*:\1:p' src/Versao.props)
for ALVO in nanoCAD26 nanoCAD5; do
  SHA=$(sha256sum appload/$ALVO/OrcamentoNanoCAD.dll | cut -d' ' -f1)
  printf '{\n  "versao": "%s",\n  "sha256": "%s",\n  "arquivo": "OrcamentoNanoCAD.dll"\n}\n' "$VERSAO" "$SHA" > appload/$ALVO/versao.json
  ls -l appload/$ALVO/
done
