' RIOS — Startup Silencioso
' Inicia o servidor RIOS em background sem abrir janela de terminal.
' Para encerrar: abra o Gerenciador de Tarefas e finalize o processo "python.exe"

Dim WshShell, rios_dir, cmd

Set WshShell = CreateObject("WScript.Shell")

' Pasta onde este arquivo está salvo
rios_dir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Comando para iniciar o servidor
cmd = "python """ & rios_dir & "\rios_server.py"""

' Roda sem janela (0 = oculto, False = não espera terminar)
WshShell.Run cmd, 0, False

' Aguarda 2 segundos e abre o navegador
WScript.Sleep 2000
WshShell.Run "http://localhost:8765", 1, False
