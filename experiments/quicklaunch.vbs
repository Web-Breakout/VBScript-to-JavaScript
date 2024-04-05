Set objShell = CreateObject("Shell.Application")
strDownloads = CreateObject("WScript.Shell").SpecialFolders("Desktop")
objShell.ShellExecute strDownloads & "\My vbs app", "", "", "open", 1
