Public Class Application
	Public Shared Property UserName As String
	Public Shared Property DisplayAlerts As Boolean
	Public Shared Property ScreenUpdating As Boolean
	Public Shared Property EnableEvents As Boolean
	Public Shared Property Calculation As Long
	Public Shared Property Interactive As Boolean
	Public Shared Property Cursor As Long

    Public Shared Property ThisWorkbook As Workbook

	Public Shared Property ActiveSheet As Worksheet
	Public Shared Property ActiveWorkbook As Workbook
	
	Public Shared Property Sheets As Worksheets
	End Function

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

	Public Shared Function GetOpenFilename(
		Optional FileFilter As String, Optional FilterIndex As Long, 
		Optional Title As String, Optional ButtonText As String, Optional MultiSelect As Boolean = False) As Object
	End Function

	Public Shared Sub Quit()
	End Sub
End Class