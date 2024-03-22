Function sendMessage(msgText, msgType, Title, EndOn, PreNext, PreCancel, PreExt, OpenApplication)
    Dim result
    result = MsgBox(msgText, msgType, Title)

    If result = vbOK Then
        ' Code to execute when the OK button is clicked
        If PreNext <> "" Then
            MsgBox PreNext
        End If
        If EndOn = "confirm" Then
            WScript.Quit
        End If
    ElseIf result = vbCancel Then
        ' Code to execute when the Cancel button is clicked
        If PreCancel <> "" Then
            MsgBox PreCancel
        End If
        If EndOn = "cancel" Then
            WScript.Quit
        End If
	ElseIf result = vbYes Then
        ' Code to execute when the Cancel button is clicked
        If PreYes <> "" Then
            MsgBox PreNext
        End If
        If EndOn = "confirm" Then
            WScript.Quit
        End If
	ElseIf result = vbNo Then
        ' Code to execute when the Cancel button is clicked
        If PreNo <> "" Then
            MsgBox PreCancel
        End If
        If EndOn = "cancel" Then
            WScript.Quit
        End If
	Else
		' Code to execute when other buttons are clicked
        If PreExt <> "" Then
            MsgBox PreExt
        End If
        If EndOn = "ext" Then
            WScript.Quit
        End If
    End If

    ' Run the specified application if the path is not blank
    If OpenApplication <> "" Then
        CreateObject("WScript.Shell").Run OpenApplication
    End If
End Function

' Usage
Dim UserInput
sendMessage "Welcome to [app name here]", vbOKCancel, "Pre-run", "cancel", "", "", "", ""
UserInput = InputBox("What would you like to be called?", "Name")
MsgBox "Ok, " & UserInput & "."
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "img.html", "", "", "open", 1

