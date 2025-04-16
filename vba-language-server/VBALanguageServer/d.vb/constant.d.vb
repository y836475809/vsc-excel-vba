Public Module ExcelVBAConstant
    Public Const vbCrLf = "\r\n"

    Public Const vbUpperCase = 1
    Public Const vbLowerCase = 2
    Public Const vbProperCase = 3
    Public Const vbWide = 4
    Public Const vbNarrow = 8
    Public Const vbKatakana = 16
    Public Const vbHiragana = 32
    Public Const vbUnicode = 64
    Public Const vbFromUnicode = 128

    ' File, Folder Attributes
    Public Const Normal = 0
    Public Const ReadOnly = 1
    Public Const Hidden = 2
    Public Const System = 4
    Public Const Volume = 8
    Public Const Directory = 16
    Public Const Archive = 32
    Public Const Alias = 1024
    Public Const Compressed = 2048

    ' OpenAsTextStream iomode
    Public Const ForReading = 1
    Public Const ForWriting = 2
    Public Const ForAppending = 8
    
    ' OpenAsTextStream format
    Public Const TristateUseDefault = -2
    Public Const TristateTrue = -1
    Public Const TristateFalse = 0

    ' GetSpecialFolder FolderSpec
    Public Const WindowsFolder = 0
    Public Const SystemFolder = 1
    Public Const TemporaryFolder = 2
End Module
