Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 実行時コマンドライン引数を取得する為の関数群


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
