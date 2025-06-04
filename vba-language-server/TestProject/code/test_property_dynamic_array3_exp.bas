Dim m
Dim m1(,)
Dim m2() As Long
Dim m3() As Long

Private Sub R__Name1(argName As String)
	ReDim m1(1, 2)

	ReDim m22(1, 3)
End Sub

Dim m21()
Dim m22(,)
Dim m23(,) As Long

Private Sub R__Name2(argName As String)
	Dim m2(,) As Long
	ReDim m2(1, 3)

	ReDim m23(1, 3)
End Sub
WriteOnly Property Name1 As String
WriteOnly Property Name2 As String