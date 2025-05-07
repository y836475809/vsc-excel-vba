
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
    Public Function InStr(String1 As String, String2 As String, 
        Optional Compare As Long = -1) As Long
    End Function

    Public Function CStr(Expression As Object) As String
    End Function

    Public Function CDbl(Expression As String) As Double
    End Function

    Public Function CByte(Expression As String) As Byte
    End Function

    Public Function CInt(Expression As String) As Integer
    End Function

    Public Function CLng(Expression As String) As Long
    End Function

    Public Function CBool(Expression As String) As Boolean
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
    Public Function LenB(Str As String) As Long
    End Function

    Public Function LCase(Str As String) As String
    End Function

    Public Function UCase(Str As String) As String
    End Function

    Public Function UCase(Expression As String, Optional Format As String = "") As String
    End Function

    Public Function Mid(Str As String, Start As Long, Optional Length As Long = -1) As String
    End Function

    Public Function StrConv(Str As String, Conversion As Integer)
    End Function

    Public Function Now() As Date
    End Function

    Public Function Left(Str As String, length As Long) As String
    End Function

    Public Function Right(Str As String, length As Long) As String
    End Function

    Public Function LBound(Ary As Object, Optional dimension As Long = 1) As Long
    End Function

    Public Function UBound(Ary As Object, Optional dimension As Long = 1) As Long
    End Function

    Public Function Val(Str As String) As Double
    End Function
    
    Public Function Chr(CharCode As Long) As String
    End Function

    Public Function Str(Number As Long) As String
    End Function
    Public Function Str(Number As Double) As String
    End Function

    Public Function FreeFile(Optional Rangenumber As Object = 0) As Integer
    End Function

    Public Sub Print(Filenumber As Object, Expression As String)
    End Sub

    Public Sub Write(Filenumber As Object, Expression As String)
    End Sub

    Public Sub Close(Filenumber As Object)
    End Sub

    ' Scripting.Dictionary
    ' Scripting.FileSystemObject
    Public Function CreateObject(Name As String) As Object
    End Function

    Public Function IIf(Expr As Boolean, TruePart As Object, FalsePart As Object) As Object
    End Function

    Public Function IsNumeric(Expression As String) As Boolean
    End Function

    Public Function Array(ParamArray args As Object()) As Object
    End Function
End Module