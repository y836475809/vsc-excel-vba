Imports System.Collections

Public Class Workbooks : Implements IEnumerable
    Default Public Property Item(name As String) As Workbook
        Get : End Get
        Set(value As Workbook) : End Set
    End Property

    Default Public Property Item(index As Long) As Workbook
        Get : End Get
        Set(value As Workbook) : End Set
    End Property

    Public Property Count As Long

    Public Function Add(Optional Template As Object) As Workbook
    End Function

    ' UpdateLinks 0 ,3
    ' Format 1:tab 2:, 3: (space) 4:; 5:none 6: custom
    ' Origin XlPlatform
    ' CorruptLoad XlCorruptLoad
    Public Function Open(FileName As String, 
        Optional UpdateLinks As Long = 0, 
        Optional [ReadOnly] As Boolean = False, 
        Optional Format As Long = 2, 
        Optional Password As String = "", 
        Optional WriteResPassword As String = "", 
        Optional IgnoreReadOnlyRecommended As Boolean = False, 
        Optional Origin As Long = xlWindows,
        Optional Delimiter As String = "", 
        Optional Editabl As Boolean = False, 
        Optional Notify As Boolean = False, 
        Optional Converter As Object = Nothing, 
        Optional AddToMru As Boolean = False, 
        Optional Local As Boolean = False, 
        Optional CorruptLoad As Long = xlNormalLoad
        ) As Workbook
    End Function
End Class