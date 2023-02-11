Public Class Dictionary
  Public Count As Long

  Public Sub Add(key As String, item As Object)
  End Sub
  Public Sub Add(key As Integer, item As Object)
  End Sub

  Public Function Items() As Object()
  End Function

  Public Function Keys() As String()
  End Function
  Public Function Keys() As Integer()
  End Function

  Public Function Exists(key As String) As Boolean
  End Function
  Public Function Exists(key As Integer) As Boolean
  End Function

  Public Sub Remove(key As String)
  End Sub
  Public Sub Remove(key As Integer)
  End Sub

  Public Sub RemoveAll()
  End Sub
End Class