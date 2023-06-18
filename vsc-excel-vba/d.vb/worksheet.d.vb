Imports System.Collections

Public Class AutoFilter
    Public Property Range As Range
    Public Sub ApplyFilter()
    End Sub
    Public Sub ShowAllData()
    End Sub
End Class

Public Class Worksheet
    Public Property Name As String
    Public Property Range As Range
    Public Property Cells As Cells
    Public Property Columns As Range
    Public Property AutoFilter As AutoFilter


    Public Sub Activate()
    End Sub
    Public Sub ShowAllData()
    End Sub
End Class