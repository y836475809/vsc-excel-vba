Imports System.Collections

Public Class Workbook
    Public Property Path As String
    Public Property Name As String
    Public Property Fullname As String
    Public Property Worksheets As Worksheets
    Public Property ActiveSheet As Worksheet
    Public Property Sheets As Sheets
    Public Property Saved As Boolean

    Public Sub Activate()
    End Sub

    Public Sub Close(
        Optional SaveChanges As Boolean = True, 
        Optional FileName As String = "", 
        Optional RouteWorkbook As Boolean = False)
    End Sub
End Class