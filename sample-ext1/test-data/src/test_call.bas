Attribute VB_Name = "test_call"
Option Explicit

Sub Main()
    testSub 	    '����
    Call testSub 	'����
    testSubArg 123 	'����
    testSubArg(123) 	'����
    Call testSubArg 123 	'�G���[
    Call testSubArg(123) 	'����

    testFunc 	    '����
    Call testFunc 	'����
    testFuncArg 123 	'����
    testFuncArg(123) 	'����
    Call testFuncArg 123 	'�G���[
    Call testFuncArg(123) 	'����

    Dim ret As Long
    ret = testFunc 	        '����
    ret = Call testFunc 	'�G���[
    ret = testFuncArg 123 	'�G���[
    ret = testFuncArg(123) '����
    ret = Call testFuncArg 123     '�G���[
    ret = Call testFuncArg(123) 	'�G���[
End Sub

Sub MainArgs()
    testSubArg2 1, 2       '����
    testSubArg2(1, 2)      '�G���[
    Call testSubArg2 1, 2  '�G���[
    Call testSubArg2(1, 2) '����

    testFuncArg2 1, 2        '����
    testFuncArg2 (1,2)       '�G���[
    Call testFuncArg2 1,2    '�G���[
    Call testFuncArg2(1, 2)  '����

    Dim ret As Long
    ret = testFuncArg2 1,2       '�G���[
    ret = testFuncArg2(1, 2)     '����
    ret = Call testFuncArg2 1,2  '�G���[
    ret = Call testFuncArg2(1,2) '�G���[
End Sub

Sub testSub()
End Sub

Sub testSubArg(a As Long)
End Sub

Function testFunc() As Long
    testFunc = 10
End Function

Function testFuncArg(a As Long) As Long
    testFuncArg = 10
End Function

Sub testSubArg2(a As Long, b As Long)
End Sub

Function testFuncArg2(a As Long, b As Long) As Long
    testFuncArg2 = 10
End Function