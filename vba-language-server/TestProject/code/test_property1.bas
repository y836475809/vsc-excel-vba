
Public Property Get Name1() As String
	Name1 = LCase(Name)
End Property

Public Property Set Name1(argName As String)
	Dim a As String
	a = argName
End Property

Public Property Set Name2(argName As String)
	Me.Name = argName
End Property

Public Property Get Name3() As String
	Name3 = LCase(Name)
End Property
