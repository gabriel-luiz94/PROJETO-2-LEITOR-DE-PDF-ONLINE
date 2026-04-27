' Script VBS — Inicia o Leitor PDF sem janela de console visível
' Duplo clique neste arquivo para abrir o leitor direto no browser

Dim WshShell, scriptDir, batPath

Set WshShell = CreateObject("WScript.Shell")
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
batPath = scriptDir & "\Iniciar_Leitor_PDF.bat"

' 0 = janela oculta | false = não aguardar
WshShell.Run Chr(34) & batPath & Chr(34), 0, False

Set WshShell = Nothing
