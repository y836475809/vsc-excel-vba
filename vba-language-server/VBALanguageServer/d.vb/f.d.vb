
Public Module f
    ' Range
    Public Function Range(Cell1 As String, Optional Cell2 As String = "") As Range
    End Function
    Public Function Range(Cell1 As Range, Optional Cell2 As Range = Nothing) As Range
    End Function

    Public Function Rows(Cell1 As String, Optional Cell2 As String = "") As Range
    End Function
    Public Function Rows(Cell1 As Range, Optional Cell2 As Range = Nothing) As Range
    End Function

    Public Function Selection() As Range
    End Function

    ' Cells
    Public Function Cells(Optional RowIndex As Long, Optional ColumIndex As Long) As Range
    End Function

    ' Workbooks
    Public Function Workbooks() As Workbooks
    End Function
    Public Function Workbooks(name As String) As Workbook
    End Function
    Public Function Workbooks(index As Long) As Workbook
    End Function

    ' Worksheets
    Public Function Worksheets() As Worksheets
    End Function
    Public Function Worksheets(name As String) As Worksheet
    End Function
    Public Function Worksheets(index As Long) As Worksheet
    End Function

    Public Function Rows() As Range
    End Function

    Public Function ThisWorkbook() As Workbook
    End Function
End Module