Attribute VB_Name = "test_call"
Option Explicit

Sub Main()
    testSub 	    '����
    Call testSub 	'����
    testSubArgs 123 	'����
    testSubArgs(123) 	'����
    Call testSubArgs 123 	'�G���[
    Call testSubArgs(123) 	'����

    testFunc 	    '����
    Call testFunc 	'����
    testFuncArgs 123 	'����
    testFuncArgs(123) 	'����
    Call testFuncArgs 123 	'�G���[
    Call testFuncArgs(123) 	'����

    Dim ret As Long
    ret = testFunc 	        '����
    ret = Call testFunc 	'�G���[
    ret = testFuncArgs 123 	'�G���[
    ret = testFuncArgs(123) '����
    ret = Call testFuncArgs 123     '�G���[
    ret = Call testFuncArgs(123) 	'�G���[
End Sub

Sub testSub()
End Sub

Sub testSubArgs(a As Long)
End Sub

Function testFunc() As Long
    testFunc = 10
End Function

Function testFuncArgs(a As Long) As Long
    testFuncArgs = 10
End Function