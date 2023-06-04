VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test個人"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
'''  個人
''' </summary>

''' <summary>
'''  メンバ変数
''' </summary> 
Public Name As String

Private Age As Long
Public Mother As Person

Public Property Get Name2() As String
  Name2 = LCase(Name)
End Property

Public Property Let Name2(argName As String)
  Me.Name = argName
End Property


Property Get Name3() As String
  Name3 = LCase(Name)
End Property

Property Let Name3(argName As String)
  Me.Name = argName
End Property

''' <summary>
'''  テストメッセージ
''' </summary>
''' <param name='val1'></param>
''' <param name='val2'></param>
''' <returns></returns>
 Public Sub SaHello(val1, val2)
    MsgBox "Hello!"
End Sub