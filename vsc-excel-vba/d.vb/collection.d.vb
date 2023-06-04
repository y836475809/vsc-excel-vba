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

  Public Property Item(index As Integer) As Object
  End Property
  Public Property Item(key As String) As Object
  End Property
End Class