Public Class Application
    Public Shared Property ThisWorkbook As Workbook
    Public Shared Property Workbooks As Workbooks

    Public Shared Function Workbooks() As Workbooks
	End Function

	Public Shared Function Workbooks(name As String) As Workbook
	End Function
    
	Public Shared Function Workbooks(index As Long) As Workbook
	End Function
End Class