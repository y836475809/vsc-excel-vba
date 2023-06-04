Imports System.Collections 

Public Class Collection : Implements IEnumerable
    Public Count As Long

    Public Sub Add(item As Object)
    End Sub
    
    Public Sub Add(item As Object, key As String)
    End Sub

    Public Sub Remove(index As Long)
    End Sub

    Public Sub Remove(key As String)
    End Sub

    Public Function Item(index As Integer) As Object
    End Function

    Public Function Item(key As String) As Object
    End Function
End Class