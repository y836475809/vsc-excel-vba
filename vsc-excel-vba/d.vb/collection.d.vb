Imports System.Collections 

Public Class Collection : Implements IEnumerable
    Public Count As Long

    Public Sub Add(Item As Object)
    End Sub
    
    Public Sub Add(Item As Object, Key As String)
    End Sub

    Public Sub Remove(Index As Long)
    End Sub

    Public Sub Remove(Key As String)
    End Sub

    Default Public Property Item(Index As Long) As Object
        Get : End Get
        Set(Value As Object) : End Set
    End Property

    Default Public Property Item(Key As String) As Object
        Get : End Get
        Set(Value As Object) : End Set
    End Property
End Class