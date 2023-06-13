Public Class Range
    Default Public Property Item(Cell1 As String, Optional Cell2 As String = "") As Range
        Get : End Get
        Set(Value As Range) : End Set
    End Property

    Default Public Property Item(Cell1 As Range, Optional Cell2 As Range = Nothing) As Range
        Get : End Get
        Set(Value As Range) : End Set
    End Property

    Default Public Property Item(RowIndex As Long, Optional ColumIndex As Long = 0) As Range
        Get : End Get
        Set(Value As Range) : End Set
    End Property

    Public Sub Activate()
    End Sub

    Public Sub ClearOutline()
    End Sub

    Public Sub Clear()
    End Sub

    Public Function Select()
    End Function

    Public ReadOnly Property Areas As Areas
    Public ReadOnly Property Count As Long
    Public Property Cells As Range
    Public Property Value As Object
    Public Property Rows As Range

    Public Function AutoFilter(Field As Long, 
        Criteria1 As String, AutoFilterOperator As Long, 
        Optional Criteria2 As String = "",
        Optional VisibleDropDown As Boolean = True) As Object
    End Function
End Class