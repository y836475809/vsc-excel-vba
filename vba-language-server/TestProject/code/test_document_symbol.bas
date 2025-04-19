Attribute VB_Name = "test_document_symbol"
Option ExplicitOn

Private fieldvar1 As String, fieldvar2 As Long
Private fieldvar3 As String

Public Enum testEnum
  e1
  e2
  e3
End Enum

Public Type type1
  num1 As Long
  name1 As String
  utc1() As Integer
End Type

Private Type type2
  num2 As Long
  name2 As String
End Type

Type type3
  num3 As Long
  name3 As String
End Type

Property Get prop_get1() As String
  dim name1 as string
  name1 = 1
  Set prop_get1 = name1
End Property

Property Let prop_let1(ByVal arg As String)
  name = arg
End Property

Property Set prop_set1(ByVal arg As String)
  name = arg
End Property

Sub Main()
  Dim num1 as Long
  Dim name1 As string

  Debug.Print "ƒeƒXƒg"

  Dim End_Row As Long, End_Col As Long
  End_Row = Cells(Rows.Count, 2).End(xlUp).Row
End Sub
