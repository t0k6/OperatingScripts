Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WSCRIPT.EXEなどのCSCRIPT.EXE以外の使用して実行された場合に
' CSCRIPT.EXEを使用して実行しなおす。
' このコードを利用する場合には、出来るだけ早い段階で呼び出すことで、
' それ以降のコードが実行されること無く、自動的に実行しなおすことが出来る。

If Right((LCase(WScript.FullName)),11) <> "cscript.exe" then
    Dim strCmd, i
    strCmd = "cscript """ & Wscript.scriptfullname & """"
    
    For i = 0 to WScript.Arguments.Count - 1
        strCmd = strCmd & " " & WScript.Arguments.Item(i)
    Next
    
    ' WScript.Echo "このスクリプトはCSCRIPT.EXEを使用して実行して下さい。" & vbCrlf & "例： " & strCmd
    
    Dim objShell
    Set objShell = WScript.CreateObject("WScript.Shell")
    objShell.run "cmd /K " & strCmd ' CSCRIPTで実行しなおし
    WScript.Quit(0)
End if