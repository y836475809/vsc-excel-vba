Attribute VB_Name = "m1"
Option Explicit


''' <summary>
'''  ���W���[��buf
''' </summary>
Private buf As String

' Function Range()

' End Function
''' <summary>
'''  �e�X�g���b�Z�[�W
''' </summary>
''' <returns></returns>
Sub Sample1()
    'Range("A1").
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
    Dim x As MSXML2.DOMDocument60
    x.getElementsByTagName()
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
    numDate.

    Dim numBool as Boolean
    numBool = True
    Dim numObj as Object
    numObj = "Object"

    Dim numVa as Variant
    numVa = "Variant"
    Dim numVa2 as Variant
    numVa2 = "Variant"
End Sub

Function callFunc() As Long
    Dim p2 As New Person
    p2.SayHello(1,1)

    Dim datFile As String
    Dim Output As String
    Open datFile For Output As #1

End Function

Sub call1()
    Dim p2 As New Person
    Set p2 = New Person
    buf = "ss"
    nondim = 10
    
    'buf.
    'thisModule.Range("").
    'Range("A3:E5").
    Call Sample2()
    p2.SayHello(1,1)
    p2.QSayHell
    ' completion position

End Sub