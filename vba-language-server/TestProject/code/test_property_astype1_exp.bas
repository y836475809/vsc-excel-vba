Public Property Name1() As {0}{1} : Set : End Set : Get
	Name1 = LCase(Name)
End Get : End Property

Private Sub R__Name1(argName1{1} As {0})
	Dim a As String
	a = argName
End Sub

Public ReadOnly Property Name2() As {0}{1} : Get
	Name2 = LCase(Name)
End Get : End Property

Private Sub R__Name3(argName1{1} As {0})
	Dim a As String
	a = argName
End Sub
Public WriteOnly Property Name3 As {0}{1}