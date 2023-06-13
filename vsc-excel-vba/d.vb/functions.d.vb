
Public Module ExcelVBAFunctions
    Public Function MsgBox(
        Prompt As String,
        Optional Buttons As Long = 0, Optional Title As String = "",
        Optional Helpfile As String = "", Optional Context As Long = 0) As Long
    End Function

    Public Function Replace(Expression As String, Find As String, Replace As String,
        Optional Start As Long = 1, Optional Count As Long = -1, Optional Compare As Long = 0) As String
    End Function

    Public Function Split(
        Expression As String, Delimiter As String, 
        Optional Optional limit As Long = 0, Optional compare As Long = 0) As String()
    End Function

    Public Function InStr(Start As Long, 
        String1 As String, String2 As String, 
        Optional Compare As Long = -1) As Long
    End Function
    Public Function InStr(Optional Start As Long, 
        String1 As String, String2 As String, 
        Optional Compare As Long = -1) As Long
    End Function

    Public Function CStr(Expression As Object) As String
    End Function

    Public Function CDbl(Expression As String) As Double
    End Function

    Public Function Join(Sourcearray() As String, Optional delimiter As String = " ") As String
    End Function

    Public Function LTrim(Str As String) As String
    End Function

    Public Function RTrim(Str As String) As String
    End Function
    
    Public Function Trim(Str As String) As String
    End Function

    Public Function Len(Str As String) As Long
    End Function

    Public Function UCase(Str As String) As String
    End Function

    Public Function UCase(Expression As String, Optional Format As String = "") As String
    End Function
End Module