Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ���s���R�}���h���C���������擾����ׂ̊֐��Q


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
