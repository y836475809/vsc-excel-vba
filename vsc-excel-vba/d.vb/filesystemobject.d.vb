Public Class TextStream
    Public ReadOnly Property AtEndOfLine As Boolean
    Public ReadOnly Property AtEndOfStream As Boolean
    Public ReadOnly Property Column As Long
    Public ReadOnly Property Line As Long

    Public Sub Close()
    End Sub

    Public Function Read(Characters As Long) As String
    End Function

    Public Function ReadAll() As String
    End Function

    Public Function ReadLine() As String
    End Function

    Public Sub Skip(Characters As Long)
    End Sub

    Public Sub SkipLine()
    End Sub

    Public Sub Write(Str As String)
    End Sub

    Public Sub WriteBlankLines(Lines As Long)
    End Sub

    Public Sub WriteLine(Optional Str As String = "")
    End Sub 
End Class

Public Class Drive
    Public ReadOnly Property Drives As Collection
    Public ReadOnly Property AvailableSpace As Long
    Public ReadOnly Property DriveLetter As String
    ' 0: Unknown
    ' 1: Removable
    ' 2: Fixed
    ' 3: Network
    ' 4: CD-ROM
    ' 5: RAM Disk
    Public ReadOnly Property DriveType As Long
    Public ReadOnly Property FileSystem As String
    Public ReadOnly Property FreeSpace As Long
    Public ReadOnly Property IsReady As Boolean
    Public ReadOnly Property Path As String
    Public ReadOnly Property RootFolder As Folder
    Public ReadOnly Property SerialNumber As Long
    Public ReadOnly Property ShareName As String
    Public ReadOnly Property TotalSize As Long
    Public Property VolumeName As String
End Class

Public Class File
    Public Property Attributes As Long
    Public ReadOnly Property DateCreated As String
    Public ReadOnly Property DateLastAccessed As String
    Public ReadOnly Property DateLastModified As String
    Public ReadOnly Property Drive As String
    Public Property Name As String
    Public ReadOnly Property ParentFolder As Folder
    Public ReadOnly Property Path As String
    Public ReadOnly Property ShortName As String
    Public ReadOnly Property ShortPath As String
    Public ReadOnly Property Size As Long
    Public ReadOnly Property Type As String
    
    Public Sub Copy(Source As String, Destination As String, Optional Overwrite As Boolean = True)
    End Sub

    Public Sub Delete(Optional force As Boolean = False)
    End Sub

    Public Sub Move(Destination As String)
    End Sub

    Public Function OpenAsTextStream(Optional IOmode As Long = ForReading, Optional Format As Long = TristateFalse) As TextStream
    End Function
End Class

Public Class Folder
    Public Property Attributes As Long
    Public ReadOnly Property DateCreated As String
    Public ReadOnly Property DateLastAccessed As String
    Public ReadOnly Property DateLastModified As String
    Public ReadOnly Property Drive As String
    Public Property Name As String
    ' File Object Collection
    Public ReadOnly Property Files As Collection
    ' bFolder Object Collection
    Public ReadOnly Property SubFolders As Collection
    Public ReadOnly Property IsRootFolder As Boolean
    Public ReadOnly Property ParentFolder As Folder
    Public ReadOnly Property Path As String
    Public ReadOnly Property ShortName As String
    Public ReadOnly Property ShortPath As String
    Public ReadOnly Property Size As Long
    Public ReadOnly Property Type As String
    
    Public Sub Copy(Source As String, Destination As String, Optional Overwrite As Boolean = True)
    End Sub

    Public Sub Delete(Optional force As Boolean = False)
    End Sub

    Public Sub Move(Destination As String)
    End Sub
End Class

Namespace Scripting

Public Class FileSystemObject
    Public ReadOnly Property Drives As Collection

    Public Function BuildPath(Path As String, Name As String) As String
    End Function
 
    Public Sub CopyFile(Source As String, Destination As String, Optional Overwrite As Boolean = True)
    End Sub
    
    Public Sub CopyFolder(Source As String, Destination As String, Optional Overwrite As Boolean = True)
    End Sub
    
    Public Function CreateFolder(FolderPath As String) As Folder
    End Function

    Public Function CreateTextFile(FileName As String, Optional Overwrite As Boolean = True, Optional Unicode As Boolean = False) As TextStream
    End Function

    Public Sub DeleteFile(FilePath As String, Optional force As Boolean = False)
    End Sub

    Public Sub DeleteFolder(FolderPath As String, Optional force As Boolean = False)
    End Sub

    Public Function DriveExists(DrivePath As String) As Boolean
    End Function

    Public Function FileExists(FilePath As String) As Boolean
    End Function

    Public Function FolderExists(FolderPath As String) As Boolean
    End Function

    Public Function GetAbsolutePathName(Path As String) As String
    End Function

    Public Function GetBaseName(Path As String) As String
    End Function

    Public Function GetDrive(DriveSpec As String) As Drive
    End Function

    Public Function GetDriveName(Path As String) As String
    End Function

    Public Function GetExtensionName(Path As String) As String
    End Function

    Public Function GetFile(Path As String) As File
    End Function

    Public Function GetFileName(Path As String) As String
    End Function

    Public Function GetFolder(Path As String) As Folder
    End Function

    Public Function GetParentFolderName(Path As String) As String
    End Function

    Public Function GetSpecialFolder(FolderSpec As Long) As Folder
    End Function

    Public Function GetTempName() As String
    End Function

    Public Sub MoveFile(Source As String, Destination As String)
    End Sub

    Public Sub MoveFolder(Source As String, Destination As String)
    End Sub

    Public Function OpenTextFile(FileName As String, Optional IOmode As Long = ForReading, Optional Create As Boolean = False, Optional Format As Long = TristateFalse) As TextStream
    End Function
End Class

End Namespace