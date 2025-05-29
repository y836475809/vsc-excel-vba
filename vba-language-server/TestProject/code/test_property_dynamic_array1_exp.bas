Dim m
Dim m1(,)
Dim m2(,) As Long
Dim m3() As Long

Public ReadOnly Property Name1(argName As String) : Get
	ReDim m1(1, 2)

	Dim ary1 As Long

	Dim ary2(,) As Long
	ReDim ary2(1, 3)

	Dim ary3(,) As Long:ReDim ary3(1, 3)         'redim
	ReDim ary3(1, 3)
	ReDim ary3(0 TO 1, 0 TO 3)
End Get : End Property

WriteOnly Property Name2() As String : Set(argName As String)
	ReDim m2(1, 2)

	Dim ary1 As Long

	Dim ary2(,) As Long
	ReDim ary2(1, 3)

	Dim ary3(,) As Long:ReDim ary3(1, 3)         'redim
	ReDim ary3(1, 3)
	ReDim ary3(0 TO 1, 0 TO 3)
End Set : End Property

WriteOnly Property Name3() As String : Set(argName As String)
	ReDim m2(1, 3)

	Dim ary1 As Long

	Dim ary2(,) As Long
	ReDim ary2(1, 3)

	Dim ary3(,) As Long:ReDim ary3(1, 3)         'redim
	ReDim ary3(1, 3)
	ReDim ary3(0 TO 1, 0 TO 3)
End Set : End Property