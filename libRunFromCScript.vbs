Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WSCRIPT.EXE�Ȃǂ�CSCRIPT.EXE�ȊO�̎g�p���Ď��s���ꂽ�ꍇ��
' CSCRIPT.EXE���g�p���Ď��s���Ȃ����B
' ���̃R�[�h�𗘗p����ꍇ�ɂ́A�o���邾�������i�K�ŌĂяo�����ƂŁA
' ����ȍ~�̃R�[�h�����s����邱�Ɩ����A�����I�Ɏ��s���Ȃ������Ƃ��o����B

If Right((LCase(WScript.FullName)),11) <> "cscript.exe" then
    Dim strCmd, i
    strCmd = "cscript """ & Wscript.scriptfullname & """"
    
    For i = 0 to WScript.Arguments.Count - 1
        strCmd = strCmd & " " & WScript.Arguments.Item(i)
    Next
    
    ' WScript.Echo "���̃X�N���v�g��CSCRIPT.EXE���g�p���Ď��s���ĉ������B" & vbCrlf & "��F " & strCmd
    
    Dim objShell
    Set objShell = WScript.CreateObject("WScript.Shell")
    objShell.run "cmd /K " & strCmd ' CSCRIPT�Ŏ��s���Ȃ���
    WScript.Quit(0)
End if