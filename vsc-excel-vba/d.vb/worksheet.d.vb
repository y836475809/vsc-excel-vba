Imports System.Collections

Public Class AutoFilter
    Public Sub ApplyFilter()
    End Sub
    Public Sub ShowAllData()
    End Sub
End Class

Public Class Worksheet
    Public Property Name As String
    Public Property Range As Range
    Public Property Cells As Cells
    Public Property AutoFilter As AutoFilter

    Public Sub Activate()
    End Sub
    Public Sub ShowAllData()
    End Sub
End Class