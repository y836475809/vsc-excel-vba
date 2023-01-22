Option Explicit On

''' <summary>
'''  個人
''' </summary>
Public Class Person
    ''' <summary>
    '''  メンバ変数
    ''' </summary> 
    Public Name As String

    Private Age As Long
    Public Mother As Person

    ''' <summary>
    '''  テストメッセージ
    ''' </summary>
    ''' <param name='val1'></param>
    ''' <param name='val2'></param>
    ''' <returns></returns>
    Public Sub SayHello(val1, val2)
        MsgBox "Hello!"
    End Sub
End Class