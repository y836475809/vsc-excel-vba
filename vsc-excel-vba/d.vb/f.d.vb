
Public Module f
    ' Range
    Public Function Ran(Cell1 As String, Optional Cell2 As String = "") As Range
    End Function
    Public Function Ran(Cell1 As Range, Optional Cell2 As Range = Nothing) As Range
    End Function

    Public Function Rows(Cell1 As String, Optional Cell2 As String = "") As Range
    End Function
    Public Function Rows(Cell1 As Range, Optional Cell2 As Range = Nothing) As Range
    End Function

    Public Function Selection() As Range
    End Function

    ' Cells
    Public Function Cls(Optional RowIndex As Long, Optional ColumIndex As Long) As Range
    End Function

    ' Workbooks
    Public Function WorkBks() As Workbooks
    End Function
    Public Function WorkBks(name As String) As Workbook
    End Function
    Public Function WorkBks(index As Long) As Workbook
    End Function

    ' Worksheets
    Public Function Workshts() As Worksheets
    End Function
    Public Function Workshts(name As String) As Worksheet
    End Function
    Public Function Workshts(index As Long) As Worksheet
    End Function

    Public Function Rows() As Range
    End Function

    Public Function ThisWorkbook() As Workbook
    End Function
End Module