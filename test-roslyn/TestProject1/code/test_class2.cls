Option Explicit On

Public Class Person
    ' メンバ変数
    Public Name As String
    Private Age As Long
    Public Mother As Person

    ''' <summary>
    '''  テストメッセージ
    ''' </summary>
    ''' <param name='val1'></param>
    ''' <param name='val2'></param>
    ''' <returns></returns>
    Public Sub SayHello2(val1, val2)
        MsgBox "Hello!"
    End Sub
End Class