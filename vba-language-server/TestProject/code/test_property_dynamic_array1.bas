Dim m
Dim m1()
Dim m2() As Long
Dim m3() As Long

Public Property Get Name1(argName As String)
	ReDim m1(1, 2)

	Dim ary1 As Long

	Dim ary2() As Long
	ReDim ary2(1, 3)

	ReDim ary3(1, 3) As Long 'redim
	ReDim ary3(1, 3)
	ReDim ary3(0 TO 1, 0 TO 3)
End Property

Property Let Name2(argName As String)
	ReDim m2(1, 2)

	Dim ary1 As Long

	Dim ary2() As Long
	ReDim ary2(1, 3)

	ReDim ary3(1, 3) As Long 'redim
	ReDim ary3(1, 3)
	ReDim ary3(0 TO 1, 0 TO 3)
End Property

Property Set Name3(argName As String)
	ReDim m2(1, 3)

	Dim ary1 As Long

	Dim ary2() As Long
	ReDim ary2(1, 3)

	ReDim ary3(1, 3) As Long 'redim
	ReDim ary3(1, 3)
	ReDim ary3(0 TO 1, 0 TO 3)
End Property