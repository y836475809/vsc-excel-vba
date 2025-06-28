Dim a As Long

Dim MyArray1(10) As Long
Dim MyArray2(10, 5) As String
Dim sngMulti1(7 To 13) As String 
Dim sngMulti2(1 To 5, 1 To 10) As Single 
Dim Member() As String

Set A = Obj

Sub Sub1()
  Set AB = Obj
End Sub

Private Type typeSya
  num    As Long
  name As String
  utc_name(0 To 31) As Integer
End Type

Private Type typeSya2
  num    As Long
  name As String
End Type

Property Get Name2() As String
  Name1 = 1
  Name2 = Name1
  Set Name2 = Name1
End Property

Sub Sub2()
  Set AB = Obj
End Sub

Property Let Name2(ByVal argName As String)
  Name = argName
End Property