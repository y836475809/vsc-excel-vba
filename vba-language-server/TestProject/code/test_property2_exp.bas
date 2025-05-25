Property  Name1() As String
Set : End Set
Get
	Name1 = LCase(Name)
End Get : End Property

Private Sub set_p_Name1(argName As String)
	Dim a As String
	a = argName
End Sub

Private Sub set_Name2(argName As String)
	Me.Name = argName
End Sub
Public Property Name2 As String

Property ReadOnly Name3() As String
Get
	Name3 = LCase(Name)
End Get : End Property