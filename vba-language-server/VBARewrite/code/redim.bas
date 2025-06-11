Dim m_a As long
Dim m_b() As long


Sub sub1()
  Dim sub_a As long
  Dim sub_b() As long
  Redim sub_b(1, 3)
  Redim sub_c(1 to 3, 1 to 3) As Long

  sub_a = 1
  Let sub_b = 1
  Set sub_c = 1 
End Sub

Function func1() As Long
  Dim func_a As long
  Dim func_b() As long
  Redim func_b(1, 3)
  Redim func_c(1,3) As Long

  func_a = 1
  Let func_b = 1
  Set func_c = 1 
End Function

