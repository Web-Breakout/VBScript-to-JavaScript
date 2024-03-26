Set objShell = CreateObject("Shell.Application")
strDownloads = CreateObject("WScript.Shell").SpecialFolders("Downloads")
objShell.ShellExecute strDownloads & "\img.html", "", "", "open", 1
