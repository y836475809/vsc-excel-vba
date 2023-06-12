Imports System.Collections

Public Class Worksheets : Implements IEnumerable
    Default Public Property Item(name As String) As Worksheet
        Get : End Get
        Set(value As Worksheet) : End Set
    End Property

    Default Public Property Item(index As Long) As Worksheet
        Get : End Get
        Set(value As Worksheet) : End Set
    End Property

    Public Property Count As Long

    Public Function Add() As Worksheet
    End Function
End Class