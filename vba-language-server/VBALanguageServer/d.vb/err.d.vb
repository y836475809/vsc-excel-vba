Public Class Err
    Public Shared Property Number As Long
    Public Shared Property Source As String
    Public Shared Property Description As String
    Public Shared Property Helpfile As String
    Public Shared Property Helpcontext As String

    Public Shared Sub Raise(Number As Long, 
        Optional Source As String = "", 
        Optional Description As String = "",
        Optional Helpfile As String = "", 
        Optional Helpcontext As Long = 0)
    End Sub

    Public Shared Sub Clear()
    End Sub
End Class