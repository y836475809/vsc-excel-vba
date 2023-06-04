VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' <summary>
'''  メンバ変数
''' </summary>
Public Name As String

Public Age As Long

Private Id As Long

Public Property Prop1(index As Long) As Long
End Property

''' <summary>
'''  メッセージ表示
''' </summary>
''' <param name='val1'></param>
''' <returns></returns>
Public Sub Hello(val1 As String)
    MsgBox val1
End Sub

''' <summary>
'''  メッセージ取得
''' </summary>
''' <param name='val1'></param>
''' <returns>メッセージ</returns>
Public Function GetHello(val As String) As String
    Hello = "Hello!"
End Sub