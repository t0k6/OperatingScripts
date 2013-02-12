Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ���O�t���������擾����B
' Inputs:
'   strName     �擾��������̖��O�iKey�j
'   strDefault  �擾����l�̋K��l
Private Function GetNamedArguments(ByVal strName, ByVal strDefault)
    If WScript.Arguments.Named.Exists(strName) Then
        GetNamedArguments = WScript.Arguments.Named.Item(strName)
    Else
        GetNamedArguments = strDefault
    End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' �������擾����B
Private Function GetEnvironments()
    Dim objEnv
    Set objEnv = CreateObject("Scripting.Dictionary")
    
    Dim objShell, objNetwork
    Set objShell = WScript.CreateObject("WScript.Shell")
    Set objNetwork = WScript.CreateObject("WScript.Network")
    
    objEnv.Add "Environment",       objShell.Environment("Process")
    
    objEnv.Add "ComputerName",      objNetwork.ComputerName
    objEnv.Add "UserDomain",        objNetwork.UserDomain
    objEnv.Add "UserName",          objNetwork.UserName
    
    objEnv.Add "Arguments",         WScript.Arguments
    objEnv.Add "FullName",          WScript.FullName
    objEnv.Add "Name",              WScript.Name
    objEnv.Add "Path",              WScript.Path
    objEnv.Add "ScriptFullName",    WScript.ScriptFullName
    objEnv.Add "ScriptName",        WScript.ScriptName
    objEnv.Add "Version",           WScript.Version
    
    Set GetEnvironments = objEnv    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' �z��ɗv�f��ǉ�����B
' Inputs:
'   arr     �Ώۂ̔z��
'   elm     �ǉ�����v�f
Sub PushArray(byRef arr, byRef elm)
    Dim i, tmp
    i = 0
    If IsArray(arr) Then
        ' �uDim hoge()�v�Œ�`���ꂽ�z���Ubound()�ő��G���[�̂���
        ' �v�f�𑖍����đ��݂���Ηv�f�����P���₷�d�l
        For Each tmp In arr
            i = 1
            Exit For
        Next
        If i=1 Then
            Redim Preserve arr(Ubound(arr)+1)
        Else
            Redim arr(0)    ' �v�f��������Ηv�f���P�̔z��ɒ�`���Ȃ���
        End If
    Else
        arr = Array(0)
    End If
    
    If IsObject(elm) Then
        Set arr(Ubound(arr)) = elm
    Else
        arr(Ubound(arr)) = elm
    End If
End Sub