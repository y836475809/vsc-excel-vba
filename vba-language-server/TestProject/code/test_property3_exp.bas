Public Property Name1() As String : Set : End Set : Get
	Name1 = LCase(Name)
End Get : End Property

Private Sub set_p_Name1(argName1 As String, argName2 As Object )
	Dim a As String
	a = argName
End Sub

Public WriteOnly Property Name2() As Object  : Set(  argName1 As Object , argName2 As Object )
	Dim a As String
End Set : End Property

Public WriteOnly Property Name3() : Set()
	Dim a As String
End Set : End Property