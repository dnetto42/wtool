Set WshShell = CreateObject("WScript.Shell")
keystr = ConvertToKey(WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))
dpath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
' MsgBox(keystr)

Title = "save in file ?"
Style = vbYesNo + vbInformation  
Msg = "window key : " + keystr


Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
    Set wshShell = CreateObject( "WScript.Shell" )
    strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

    dpath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutPutFile = FSO.OpenTextFile(dpath & "\Simple-Serial-" & strComputerName & ".txt" ,8 , True)
    OutPutFile.WriteLine "     ------------------------------------------"
    OutPutFile.WriteLine "                   System Details"
    OutPutFile.WriteLine "     ------------------------------------------"
    OutPutFile.WriteLine ""
    OutPutFile.WriteLine "     Computer Name:         "& strComputerName
    OutPutFile.WriteLine ""
    OutPutFile.WriteLine ""
    OutPutFile.WriteLine "     Serial Number:         "& keystr
    OutPutFile.WriteLine ""
    OutPutFile.WriteLine "     ------------------------------------------"

End If
' MsgBox MyString

Function ConvertToKey(Key)
Const KeyOffset = 52
i = 28
Chars = "BCDFGHJKMPQRTVWXY2346789"
Do
Cur = 0
x = 14
Do
Cur = Cur * 256
Cur = Key(x + KeyOffset) + Cur
Key(x + KeyOffset) = (Cur \ 24) And 255
Cur = Cur Mod 24
x = x -1
Loop While x >= 0
i = i -1
KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
If (((29 - i) Mod 6) = 0) And (i <> -1) Then
i = i -1
KeyOutput = "-" & KeyOutput
End If
Loop While i >= 0
ConvertToKey = KeyOutput
End Function