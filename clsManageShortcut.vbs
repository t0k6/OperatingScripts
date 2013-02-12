Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ManageShortcut �N���X
' �V���[�g�J�b�g�t�@�C���������B
' filepath�v���p�e�B��������SetPath���\�b�h�Ńt�@�C���p�X���w�肷�邱�ƂŁA
' ���݂���V���[�g�J�b�g�t�@�C���ł���΂��̏�񂪎擾�����B
' ���݂̗L���Ɍ��炸�A�ݒ肳�ꂽ�v���p�e�B�l�ɂ��V���[�g�J�b�g�t�@�C����
' Make���\�b�h�ō쐬�����B�܂����݂����Delete���\�b�h�ō폜�����B

Class ManageShortcut
    Private	objFSO
    Private objShortcut
    
    '�E�B���h�E���A�N�e�B�u�ɂ��ĕ\�����܂��B�E�B���h�E���ŏ����܂��͍ő剻����Ă���ꍇ�́A���̃T�C�Y�ƈʒu�ɖ߂�܂��B
    Public Property Get conWindowsStyleActive()
        conWindowsStyleActive = 1
    End Property
    
    '�E�B���h�E���A�N�e�B�u�ɂ��A�ő剻�E�B���h�E�Ƃ��ĕ\�����܂��B
    Public Property Get conWindowsStyleMax()
        conWindowsStyleMax = 3
    End Property
    
    '�E�B���h�E���ŏ������A���ɏ�ʂƂȂ�E�B���h�E���A�N�e�B�u�ɂ��܂��B
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
    ' �V���[�g�J�b�g�t�@�C���̃p�X
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
    ' ����ȊO�̃V���[�g�J�b�g�I�u�W�F�N�g�v���p�e�B
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
    ' �V���[�g�J�b�g�t�@�C���̑���
    Public Property Get Exists()
        Exists = objFSO.FileExists(strFilePath)
    End Property
    
    ''''''''''''''''''''''''''''''
    ' �V���[�g�J�b�g�t�@�C���̍쐬
    Public Function Make()
        Me.Delete() ' ���݂���Ȃ�폜���Ă���
        objShortcut.Save
    End Function
    
    ''''''''''''''''''''''''''''''
    ' �V���[�g�J�b�g�t�@�C���̍폜
    Public Function Delete()
        If Me.exists Then Me.Delete()
    End Function
End Class
