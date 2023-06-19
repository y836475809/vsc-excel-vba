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
End Class