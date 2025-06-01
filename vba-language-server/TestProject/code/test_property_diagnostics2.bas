Attribute VB_Name = "test"

Public Property Ge Name1() As String 'Get -> Ge
	Name1 = "Name"
End Property

Public Property Set Name1(argName As String)
	Dim a As String
	a = argName
End Property

Public Property Get Name2() As String
	Name2 = "Name"
End Property

Public Property Set Name3(argName As Stri) 'String -> Stri
	Dim a As String
	a = "Name"
End Property

Public Property Set Name4()
	Name3 = "Name"
End Property