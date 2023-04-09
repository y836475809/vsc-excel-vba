Attribute VB_Name = "Module1"

''' <summary>
'''  モジュールbuf
''' </summary>
Private buf As String

''' <summary>
'''  モジュール関数
''' </summary>
Public Sub Sample1()
End Sub

Private Function Sample2() As String
End Sub

Private Sub Sample3()
    buf = "ss"

    Dim numL as Long
    numL = 10

    Dim numI as Integer
    numI = 10

    Dim numB as Byte
    numB = 10

    Dim numD as Double
    numD = 10

    Dim numDate as Date
    numDate = "2000/12/31"

    Dim numBool as Boolean
    numBool = True

    Dim numObj as Object
    numObj = "Object"

    Dim numVa as Variant
    numVa = "Variant"
End Sub

Sub call1()
    Dim c As Class1
    Dim c1 As New Class1
    c1.Hello
    c1.Name

    Sample1() ' call
    Module2Sample1() ' call
    ' completion position

End Sub