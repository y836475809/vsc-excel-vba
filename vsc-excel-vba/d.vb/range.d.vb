Public Class Range
    Default Public Property Item(Cell1 As String, Optional Cell2 As String = "") As Range
        Get : End Get
        Set(Value As Range) : End Set
    End Property

    Default Public Property Item(Cell1 As Range, Optional Cell2 As Range = Nothing) As Range
        Get : End Get
        Set(Value As Range) : End Set
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