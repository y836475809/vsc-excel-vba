Public Property Get Name1(argName As String)
	Dim ary1() As Long
	ReDim ary1(2)

	ReDim ary2(2) As Long 'redim
	ReDim ary2(0 TO 2) As Long
End Property

Property Let Name2(argName As String)
	Dim ary1() As Long
	ReDim ary1(1, 2)

	ReDim ary2(1, 2) As Long 'redim
	ReDim ary2(0 TO 1, 0 TO 2) As Long
End Property

Property Set Name3(argName As String)
	Dim ary1() As Long
	ReDim ary1(1, 2, 3)

	ReDim ary2(1, 2, 3) As Long 'redim
	ReDim ary2(0 TO 1, 0 TO 2, 0 TO 3) As Long
End Property