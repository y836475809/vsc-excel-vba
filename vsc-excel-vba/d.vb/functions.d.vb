
Public Module ExcelVBAFunctions
    Public Function MsgBox(
        Prompt As String,
        Buttons As Long = 0,Title As String = "",
        Helpfile As String = "",Context As Long = 0) As Long
    End Function

    Public Function Replace(Expression As String, Find As String, Replace As String,
        Start As Long = 1, Count As Long = -1, Compare As Long = 0) As String
    End Function

    Public Function Split(
        Expression As String, Delimiter As String, 
        Optional limit As Long = 0, Optional compare As Long = 0) As String()
    End Function

    Public Function InStr(Optional Start As Long = 0, 
        String1 As String, String2 As String, 
        Optional Compare As Long = -1) As Long
    End Function

    Public Function CStr(Expression As Object) As String
    End Function

    Public Function CDbl(Expression As String) As Double
    End Function
End Module