Property Get Name1() As String
	Name1 = LCase(Name)
End Property

Property Set Name1(argName As String)
	Dim a As String
	a = argName
End Property

Property Set Name2(argName As String)
	Me.Name = argName
End Property

Property Get Name3() As String
	Name3 = LCase(Name)
End Property