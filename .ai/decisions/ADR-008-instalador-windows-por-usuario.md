# ADR-008 — Instalação Windows por usuário

## Contexto

O executável portátil é distribuído pelo GitHub e se atualiza substituindo o binário em uso depois que o processo encerra. O Windows precisa de uma entrada de desinstalação para listar o aplicativo instalado.

## Decisão

Cada release compila o executável uma vez e empacota o mesmo arquivo com Inno Setup. O instalador usa `PrivilegesRequired=lowest`, instala em `%LOCALAPPDATA%\Programs\OrcaCore` e registra a desinstalação em HKCU. O ID do produto é `OrcaCore` e identifica a mesma instalação nas versões futuras.

O atualizador continua baixando o executável portátil com seu SHA-256. A versão em Aplicativos instalados é sincronizada na abertura da versão efetivamente executada, após conferir se o executável é o que está registrado para o usuário. O banco, as skills e a configuração locais seguem na pasta da instalação e são preservados nas atualizações e na desinstalação.

## Consequências

A instalação vale para a conta atual. Cada conta instala sua própria cópia. Dados de uma cópia portátil anterior continuam na pasta de origem e só são levados para a instalação por migração explícita.
