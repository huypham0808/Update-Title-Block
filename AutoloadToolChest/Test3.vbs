Set WshShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep 1000 ' Wait for the dialog box to appear
If WshShell.AppActivate ("Bluebeam Revu x64") Then
    For i = 1 To 10
        WshShell.SendKeys "{ENTER}"
        WScript.Sleep 100
    Next
ElseIf WshShell.AppActivate ("Overwrite Tool Set") Then
    For i = 1 To 10
        WshShell.SendKeys "{ENTER}"
        WScript.Sleep 100
    Next
End If
MsgBox "Load all Tool chest file completed.", vbInformation, "Auto load Tool Chest v.1.0 - CSS VN Team"
'Link chua file btx: C:\Users\ad\AppData\Roaming\Bluebeam Software\Revu\20