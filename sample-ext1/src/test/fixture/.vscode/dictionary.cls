Public Class Dictionary
    Public Count As Long
    
    Public Sub Add(string Key, Object item)
    End Sub
    Public Sub Add(Integer Key, Object item)
    End Sub

    Public Function Items() as Object()
    End Function

    Public Function keys() as string()
    End Function
    Public Function keys() as Integer()
    End Function

    Public Function Exists(string Key) as Boolean
    End Function
    Public Function Exists(Integer Key) as Boolean
    End Function

    Public Sub Remove(string Key)
    End Sub
    Public Sub Remove(Integer Key)
    End Sub

    Public Sub RemoveAll()
    End Sub
End Class