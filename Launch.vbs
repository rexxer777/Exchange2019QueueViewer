Set objShell = CreateObject("WScript.Shell")
' Get current directory
strPath = objShell.CurrentDirectory & "\Queues.ps1"
' Run PowerShell hidden
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPath & """", 0, False
