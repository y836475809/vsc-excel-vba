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
'''  �����o�ϐ�
''' </summary>
Public Name As String

Public Age As Long

Private Id As Long

Public Property Prop1(index As Long) As Long
End Property

''' <summary>
'''  ���b�Z�[�W�\��
''' </summary>
''' <param name='val1'></param>
''' <returns></returns>
Public Sub Hello(val1 As String)
    MsgBox val1
End Sub

''' <summary>
'''  ���b�Z�[�W�擾
''' </summary>
''' <param name='val1'></param>
''' <returns>���b�Z�[�W</returns>
Public Function GetHello(val As String) As String
    Hello = "Hello!"
End Sub