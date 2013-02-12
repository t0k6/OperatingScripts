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
' 配列に要素を追加する。
' Inputs:
'   arr     対象の配列
'   elm     追加する要素
Sub PushArray(byRef arr, byRef elm)
    Dim i, tmp
    i = 0
    If IsArray(arr) Then
        ' 「Dim hoge()」で定義された配列はUbound()で即エラーのため
        ' 要素を走査して存在すれば要素数を１つ増やす仕様
        For Each tmp In arr
            i = 1
            Exit For
        Next
        If i=1 Then
            Redim Preserve arr(Ubound(arr)+1)
        Else
            Redim arr(0)    ' 要素が無ければ要素数１の配列に定義しなおす
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