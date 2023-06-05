Imports System.Collections

Public Class Dictionary : Implements IEnumerable
    Public Count As Long

    Public Sub Add(key As String, item As Object)
    End Sub

    Public Sub Add(key As Long, item As Object)
    End Sub

    Public Function Item(index As Long) As Object
    End Function

    Public Function Item(key As String) As Object
    End Function

    Public Function Items() As Object()
    End Function

    Public Function Keys() As String()
    End Function

    Public Function Keys() As Long()
    End Function

    Public Function Exists(key As String) As Boolean
    End Function

    Public Function Exists(key As Long) As Boolean
    End Function

    Public Sub Remove(key As String)
    End Sub
    
    Public Sub Remove(key As Long)
    End Sub

    Public Sub RemoveAll()
    End Sub
End Class