Imports System.Collections

Public Class Workbooks : Implements IEnumerable
    Public Property Bk(name As String) As Workbook
        Get : End Get
        Set(value As Workbook) : End Set
    End Property

    Public Property Bk(index As Long) As Workbook
        Get : End Get
        Set(value As Workbook) : End Set
    End Property

    Public Property Count As Long

    Public Function Add() As Workbook
    End Function
End Class