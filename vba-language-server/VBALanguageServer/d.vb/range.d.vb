Imports System.Collections

Public Class Interior
    Public Property Color As Object
    Public Property ColorIndex As Long
End Class

Public Class Range : Implements IEnumerable
    Default Public Property Item(Cell1 As String, Optional Cell2 As String = "") As Range
        Get : End Get
        Set(Value As Range) : End Set
    End Property

    Default Public Property Item(Cell1 As Range, Optional Cell2 As Range = Nothing) As Range
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
    Public ReadOnly Property Row As Long
    Public ReadOnly Property Column As Long
    Public ReadOnly Property Columns As Range
    Public ReadOnly Property CurrentRegion As Range
    Public Property Cells As Range
    Public Property Value As Object
    Public Property Rows As Range
    Public Property Range As Range
    Public Property Interior As Interior

    Public Function AutoFilter(
        Optional Field As Long, 
        Optional Criteria1 As String, 
        Optional AutoFilterOperator As Long, 
        Optional Criteria2 As String = "",
        Optional VisibleDropDown As Boolean = True) As Range
    End Function
    
    ' Type XlAutoFillType
    Public Sub AutoFill(
        Destination As Range, 
        Optional Type As Long)
    End Sub

    Public Sub Copy(Optional Destination As Range = Nothing)
    End Sub
    
    Public Sub Cut(Optional Destination As Range = Nothing)
    End Sub
    
    ' Shift XlDeleteShiftDirection
    Public Sub Delete(Optional Shift As Long = 0)
    End Sub

    Public Sub Parse(Optional ParseLine As String = "", Optional Destination As Range = Nothing)
    End Sub  

    Public Function Find(What As Object, 
        Optional After As Object, 
        Optional LookIn As XlFindLookIn = xlValues, 
        Optional LookAt As XlLookAt = xlWhole) As Range
    End Function

    Public Function End(Direction As XlDirection) As Range
    End Function

    Public Function SpecialCells(Type As XlCellType, Optional Value As Object) As Range
    End Function   
End Class