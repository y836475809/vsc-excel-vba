Public Class Collection
    Public Count As Long

    Public Sub Add(object item, string? Key)
    End Sub

    Public Sub Remove(Index as Long)
    End Sub
    Public Sub Remove(Key as string)
    End Sub

    Public Property Item(ByVal index As Integer) As Object
    End Property
    Public Property Item(ByVal Key As string) As Object
    End Property
End Class