Public Class Cells
    Default Public Property Item(RowIndex As Long, ColumIndex As Long) As Range
        Get : End Get
        Set(value As Range) : End Set
    End Property

    Public Function Activate()
    End Function

    Public Function Select()
    End Function

    Public Property Cells As Range
    Public Property Value As Object

    Public Function AutoFilter(Field As Long, 
        Criteria1 As String, AutoFilterOperator As Long, 
        Optional Criteria2 As String = "",
        Optional VisibleDropDown As Boolean = True) As Object
    End Function
End Class