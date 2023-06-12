Public Class Application
    Public Shared Property ThisWorkbook As Workbook

	Public Shared Property ActiveSheet As Worksheet
	Public Shared Property ActiveWorkbook As Workbook

	Public Shared Function Workbooks() As Workbooks
	End Function
	Public Shared Function Workbooks(Name As String) As Workbook
	End Function
	Public Shared Function Workbooks(Index As Long) As Workbook
	End Function

	Public Shared Function Worksheets() As Worksheets
	End Function
	Public Shared Function Worksheets(Name As String) As Worksheet
	End Function
	Public Shared Function Worksheets(Index As Long) As Worksheet
	End Function
End Class