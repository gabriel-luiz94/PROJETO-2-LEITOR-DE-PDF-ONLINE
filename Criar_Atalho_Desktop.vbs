' Cria atalho do Leitor PDF na Area de Trabalho
Dim WshShell, oShortcut, desktopPath, vbsPath

Set WshShell = CreateObject("WScript.Shell")

' Caminho do VBS de inicializacao
vbsPath = "C:\Users\gabriel.sales\Desktop\PROJETOS VSCODE\PROJETO 2 LEITOR DE PDF ONLINE\Abrir_Leitor_PDF.vbs"

' Caminho da Area de Trabalho
desktopPath = WshShell.SpecialFolders("Desktop")

' Cria o atalho .lnk
Set oShortcut = WshShell.CreateShortcut(desktopPath & "\Leitor PDF Online.lnk")
oShortcut.TargetPath      = "wscript.exe"
oShortcut.Arguments       = Chr(34) & vbsPath & Chr(34)
oShortcut.WorkingDirectory = "C:\Users\gabriel.sales\Desktop\PROJETOS VSCODE\PROJETO 2 LEITOR DE PDF ONLINE"
oShortcut.Description     = "Abrir Leitor PDF Online localmente"
oShortcut.IconLocation    = "shell32.dll,13"
oShortcut.Save

WScript.Echo "Atalho 'Leitor PDF Online' criado na Area de Trabalho com sucesso!"
Set oShortcut = Nothing
Set WshShell = Nothing
