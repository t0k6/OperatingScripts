Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 名前付き引数を取得する。
' Inputs:
'   strName     取得する引数の名前（Key）
'   strDefault  取得する値の規定値
Private Function GetNamedArguments(ByVal strName, ByVal strDefault)
    If WScript.Arguments.Named.Exists(strName) Then
        GetNamedArguments = WScript.Arguments.Named.Item(strName)
    Else
        GetNamedArguments = strDefault
    End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 環境情報を取得する。
Private Function GetEnvironments()
    Dim objEnv
    Set objEnv = CreateObject("Scripting.Dictionary")
    
    Dim objShell, objNetwork, obj
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