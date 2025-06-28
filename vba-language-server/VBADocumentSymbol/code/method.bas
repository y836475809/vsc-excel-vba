Dim m_a As long
Dim m_b() As long

Sub sub1()
  Dim sub_a As long
  sub_a = 1
End Sub

Sub sub2(ByRef a() As string, b As Long)
  Dim sub_a As long
  sub_a = 1
End Sub

Function func1() As Long
  func_a = 1
End Function

Function func2(ByRef a() As string, b As Long) As Long
  func_a = 1
End Function

