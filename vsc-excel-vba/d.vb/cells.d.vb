Public Class Cells
    Inherits Range
    Default Public Overloads Property Item(RowIndex As Long, Optional ColumIndex As Long = 0) As Range
        Get : End Get
        Set(value As Long) : End Set
        Set(value As String) : End Set
    End Property
End Class