
Property Get prop_get1() As String
  name1 = 1
  prop_get1 = name1
  Set prop_get1 = name1
End Property

Property Get prop_get2()
  name1 = 1
  prop_get1 = name1
  Set prop_get1 = name1
End Property

Property Let prop_let1(ByVal arg As String)
  name = arg
End Property

Property Set prop_set1(ByVal arg As String)
  name = arg
End Property
