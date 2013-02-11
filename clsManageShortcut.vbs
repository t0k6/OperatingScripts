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
    Public Property Let filepath(path)
        Set objShortcut = Nothing
        Set objShortcut = objShell.CreateShortcut(path)
    End Property
    Public Property Get filepath()
        filepath = objShortcut.FullName
    End Property
    Public Function SetPath(path)
        Me.filepath = path
        Set SetPath = objShortcut
    End Function
    
    ''''''''''''''''''''''''''''''
    ' それ以外のショートカットオブジェクトプロパティ
    Public Property Let targetpath(arg)
        objShortcut.TargetPath = arg
    End Property
    Public Property Get targetpath()
        targetpath = objShortcut.TargetPath
    End Property
    Public Property Let workdir(arg)
        objShortcut.WorkingDirectory = arg
    End Property
    Public Property Get workdir()
        targetpath = objShortcut.WorkingDirectory
    End Property
    Public Property Let hotkey(arg)
        objShortcut.Hotkey = arg
    End Property
    Public Property Get hotkey()
        targetpath = objShortcut.Hotkey
    End Property
    Public Property Let windowstyle(arg)
        objShortcut.WindowStyle = arg
    End Property
    Public Property Get windowstyle()
        targetpath = objShortcut.WindowStyle
    End Property
    Public Property Let description(arg)
        objShortcut.Description = arg
    End Property
    Public Property Get description()
        targetpath = objShortcut.Description
    End Property
    Public Property Let icon(arg)
        objShortcut.IconLocation = arg
    End Property
    Public Property Get icon()
        targetpath = objShortcut.IconLocation
    End Property
    Public Property Let args(arg)
        objShortcut.Arguments = arg
    End Property
    Public Property Get args()
        targetpath = objShortcut.Arguments
    End Property
    
    ''''''''''''''''''''''''''''''
    ' ショートカットファイルの存在
    Public Property Get exists()
        exists = objFSO.FileExists(strFilePath)
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
