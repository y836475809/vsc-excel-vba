VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit On

''' <summary>
'''  個人
''' </summary>

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