Dim m_a As long
Dim m_b() As Long
Dim m_c()

Sub sub1()
	ReDim m_b(1, 2)
	ReDim m_c(1, 2)

	Dim sub_a As Long
	Dim sub_b() As Long
	ReDim sub_b(1, 3)
	ReDim sub_c(1 To 3, 1 To 3) As Long 'redim
End Sub

Function func1() As Long
	ReDim m_b(1, 2)
	ReDim m_c(1, 2)

	Dim func_a As Long
	Dim func_b() As Long
	ReDim func_b(1, 3)
	ReDim func_c(1, 3) As Long 'redim
End Function