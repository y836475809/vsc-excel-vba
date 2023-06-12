Imports System.Collections

Public Class Dictionary : Implements IEnumerable
    Public Count As Long

    Public Sub Add(Key As String, Item As Object)
    End Sub

    Public Sub Add(Key As Long, Item As Object)
    End Sub

    Default Public Property Item(Index As Long) As Object
        Get : End Get
        Set(Value As Object) : End Set
    End Property

    Default Public Property Item(Key As String) As Object
        Get : End Get
        Set(Value As Object) : End Set
    End Property

    Public Function Items() As Object()
    End Function

    Public Function Keys() As String()
    End Function

    Public Function Keys() As Long()
    End Function

    Public Function Exists(Key As String) As Boolean
    End Function

    Public Function Exists(Key As Long) As Boolean
    End Function

    Public Sub Remove(Key As String)
    End Sub
    
    Public Sub Remove(Key As Long)
    End Sub

    Public Sub RemoveAll()
    End Sub
End Class