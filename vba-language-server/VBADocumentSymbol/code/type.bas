Dim a as Long
Dim b()
Dim c() as Long
Const d = 100
Const e = "100"

Public Type type1
  num As Long
  name As String
End Type

Public Type type2
  num As Long
  name As String
  obj As New Object
  utc(0 To 31) As Integer
End Type

Private Type type_private
  num As Long
  name As String
End Type

Type type_no_visible
  num As Long
  name As String
End Type

Public Type type_comment ' comment
  ' comment
  num As Long    ' comment
  name As String ' comment
  ' comment
End Type ' comment
