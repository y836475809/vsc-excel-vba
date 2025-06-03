Attribute VB_Name = "test_document_symbol"
Option ExplicitOn

Private field_var1 As String
Private field_var2 As String, field_var3 As Long
Dim field_dim_var4 as Long
Const field_const_var5 = 100

Public Enum testEnum
  e1 = 1
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

Property Get prop_get_set1() As String
  dim name1 as string
  name1 = 1
  Set prop_get_set1 = name1
End Property

Property Set prop_get_set1(name as string)
  dim name1 as string
  Set name1 = name
End Property

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

Function func1() As Long
  Dim num1 as Long
  Dim name1 As string
End Function

Sub sub1()
  Dim num1 as Long
  Dim name1 As string
End Sub
