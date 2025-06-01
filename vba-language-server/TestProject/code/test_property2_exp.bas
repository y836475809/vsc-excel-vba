Property Name1() As String : Set : End Set : Get
	Name1 = LCase(Name)
End Get : End Property

Private Sub R__Name1(argName As String)
	Dim a As String
	a = argName
End Sub

Private Sub R__Name2(argName As String)
	Me.Name = argName
End Sub

ReadOnly Property Name3() As String : Get
	Name3 = LCase(Name)
End Get : End Property
WriteOnly Property Name2 As String