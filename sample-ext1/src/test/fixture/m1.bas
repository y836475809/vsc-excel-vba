Module Module1
''' <summary>
'''  モジュールbuf
''' </summary>
Private buf As String

''' <summary>
'''  テストメッセージ
''' </summary>
''' <returns></returns>
Sub Sample1()
    Range("A1") = "tanaka"
End Sub

Sub Sample2()
    Dim testList As New Collection
    ' Set testList = New Collection
    testList.Item
End Sub

Sub Sample3()
    Dim testDict As New Dictionary
    testDict.
End Sub

Sub Sample4()
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
    Dim p2 As New Person
    ' Set p2 = New Person
    buf = "ss"
    p2.SayHello
    ' completion position

End Sub

End Module