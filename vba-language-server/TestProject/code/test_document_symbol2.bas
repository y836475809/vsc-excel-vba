Attribute VB_Name = "test_document_symbol"
Option ExplicitOn

Private field_var1 As String
Private field_var2 As String, field_var3 As Long
Dim field_dim_var4 as Long
Const field_const_var5 = 100

Public Enum testEnum
  e1 = 1
  e2
End Enum

Public Type type1
  num1 As Long
  utc1() As Integer
End Type

Property Get prop_get_set1() As String
  dim name1 as string
  name1 = 1
  Set prop_get_set1 = name1
End Property

Property Set prop_get_set1(name as string)
  dim name1 as string
  Set name1 = name
End Property

Function func1() As Long
  Dim num1 as Long
  Dim name1 As string
End Function

Sub sub1()
  Dim num1 as Long
  Dim name1 As string
End Sub