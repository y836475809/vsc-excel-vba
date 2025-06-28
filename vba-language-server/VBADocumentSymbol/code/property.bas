Dim a as Long
Dim b()
Dim c() as Long
Const d = 100
Const e = "100"

Property Get prop_get1() As String
  Set prop_get1 = name1
End Property

Property Get prop_get2()
  Set prop_get2 = name1
End Property

Property Let prop_let1(ByVal arg As String)
  name = arg
End Property

Property Set prop_set1(ByVal arg As String)
  name = arg
End Property
