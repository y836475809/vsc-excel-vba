Imports System.Collections

Public Class Areas : Implements IEnumerable
    Default Public Property Item(index As Long) As Range
        Get : End Get
        Set(value As Range) : End Set
    End Property

    Public Property Count As Long
End Class