Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ManageShortcut クラス
' ショートカットファイルを扱う。
' filepathプロパティもしくはSetPathメソッドでファイルパスを指定することで、
' 存在するショートカットファイルであればその情報が取得される。
' 存在の有無に限らず、設定されたプロパティ値によるショートカットファイルが
' Makeメソッドで作成される。また存在すればDeleteメソッドで削除される。

Class ManageShortcut
    Private	objFSO
    Private objShortcut
    
    'ウィンドウをアクティブにして表示します。ウィンドウが最小化または最大化されている場合は、元のサイズと位置に戻ります。
    Public Property Get conWindowsStyleActive()
        conWindowsStyleActive = 1
    End Property
    
    'ウィンドウをアクティブにし、最大化ウィンドウとして表示します。
    Public Property Get conWindowsStyleMax()
        conWindowsStyleMax = 3
    End Property
    
    'ウィンドウを最小化し、次に上位となるウィンドウをアクティブにします。
    Public Property Get conWindowsStyleMin()
        conWindowsStyleMin = 7
    End Property
    
    Public Sub Class_Initialize()
        Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    End Sub
    
    Public Sub Class_Terminate()
        Set objFSO = Nothing
        Set objShortcut = Nothing
    End Sub
    
    ''''''''''''''''''''''''''''''
    ' ショートカットファイルのパス
    Public Property Let FilePath(path)
        Set objShortcut = Nothing
        Set objShortcut = objShell.CreateShortcut(path)
    End Property
    Public Property Get FilePath()
        FilePath = objShortcut.FullName
    End Property
    Public Function SetPath(path)
        Me.FilePath = path
        Set SetPath = objShortcut
    End Function
    
    ''''''''''''''''''''''''''''''
    ' それ以外のショートカットオブジェクトプロパティ
    Public Property Let TargetPath(str)
        objShortcut.TargetPath = str
    End Property
    Public Property Get TargetPath()
        TargetPath = objShortcut.TargetPath
    End Property
    Public Property Let WorkDir(str)
        objShortcut.WorkingDirectory = str
    End Property
    Public Property Get WorkDir()
        WorkDir = objShortcut.WorkingDirectory
    End Property
    Public Property Let Hotkey(str)
        objShortcut.Hotkey = str
    End Property
    Public Property Get Hotkey()
        Hotkey = objShortcut.Hotkey
    End Property
    Public Property Let WindowStyle(str)
        objShortcut.WindowStyle = str
    End Property
    Public Property Get WindowStyle()
        WindowStyle = objShortcut.WindowStyle
    End Property
    Public Property Let Description(str)
        objShortcut.Description = str
    End Property
    Public Property Get Description()
        Description = objShortcut.Description
    End Property
    Public Property Let Icon(str)
        objShortcut.IconLocation = str
    End Property
    Public Property Get Icon()
        Icon = objShortcut.IconLocation
    End Property
    Public Property Let Args(str)
        objShortcut.Arguments = str
    End Property
    Public Property Get Args()
        Args = objShortcut.Arguments
    End Property
    
    ''''''''''''''''''''''''''''''
    ' ショートカットファイルの存在
    Public Property Get Exists()
        Exists = objFSO.FileExists(strFilePath)
    End Property
    
    ''''''''''''''''''''''''''''''
    ' ショートカットファイルの作成
    Public Function Make()
        Me.Delete() ' 存在するなら削除してから
        objShortcut.Save
    End Function
    
    ''''''''''''''''''''''''''''''
    ' ショートカットファイルの削除
    Public Function Delete()
        If Me.exists Then Me.Delete()
    End Function
End Class
