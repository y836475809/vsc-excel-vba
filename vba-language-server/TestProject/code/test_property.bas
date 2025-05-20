
Public Property Get Name1() As String
Name1 = LCase(Name)
End Property

Public Property Set Name1(argName As String)
Dim a As String
a = argName
End Property

Property Set Name3(argName As String)
Me.Name = argName
End Property

Property Get Name2() As String
Name2 = LCase(Name)
End Property
