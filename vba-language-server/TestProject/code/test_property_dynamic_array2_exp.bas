Public Property ReadOnly Name1(argName As String)
Get
	Dim ary1() As Long
	ReDim ary1(2)

	Dim ary2() As Long:ReDim ary2(2)         'redim
	ReDim ary2(0 TO 2) As Long
End Get : End Property

Private Sub set_Name2(argName As String)
	Dim ary1(,) As Long
	ReDim ary1(1, 2)

	Dim ary2(,) As Long:ReDim ary2(1, 2)         'redim
	ReDim ary2(0 TO 1, 0 TO 2) As Long
End Sub
Public Property Name2 As String

Private Sub set_Name3(argName As String)
	Dim ary1(,,) As Long
	ReDim ary1(1, 2, 3)

	Dim ary2(,,) As Long:ReDim ary2(1, 2, 3)         'redim
	ReDim ary2(0 TO 1, 0 TO 2, 0 TO 3) As Long
End Sub
Public Property Name3 As String