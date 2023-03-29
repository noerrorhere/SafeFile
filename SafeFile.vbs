Option Explicit

Dim ready
ready = MsgBox("Are you ready?", vbYesNo, "Prompt")

If ready = vbNo Then
    MsgBox "Error", vbCritical, "Error"
    WScript.Quit
Else
    MsgBox "Hello", vbInformation, "Prompt"
    Dim i
    For i = 1 To 5
        MsgBox "AHAHAHA", vbInformation, "Prompt"
    Next
    OpenEdge "https://genius.com/Rick-astley-never-gonna-give-you-up-lyrics"
    Dim subscribed
    subscribed = MsgBox("Are you subscribed?", vbYesNo, "Prompt")
    If subscribed = vbYes Or subscribed = vbNo Then
        MsgBox "Access to webcam declined.", vbCritical, "Error"
        MsgBox "SafeFile.vbs could not enter your WiFi.", vbCritical, "Error"
    End If
End If

Dim objShell
Set objShell = CreateObject("WScript.Shell")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim strFilePath
strFilePath = "C:\HarmlessThing.what"
Dim objFile
Set objFile = objFSO.CreateTextFile(strFilePath)
objFile.WriteLine "I made a file on your PC."
objFile.WriteLine "Find it now."
objFile.Close

Dim answer
answer = InputBox("Do you wanna play a game?", "Prompt")

If answer = "yes" Then
    ' Restart the computer
    objShell.Run "shutdown.exe -r -t 0"
Else
    objFSO.DeleteFile(strFilePath)
End If

Sub OpenEdge(url)
    Dim objShell
    Set objShell = CreateObject("Shell.Application")
    objShell.ShellExecute "msedge.exe", url, "", "", 1
End Sub



