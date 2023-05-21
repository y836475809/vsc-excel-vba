Attribute VB_Name = "test_call"
Option Explicit

Sub Main()
    testSub 	    '正常
    Call testSub 	'正常
    testSubArgs 123 	'正常
    testSubArgs(123) 	'正常
    Call testSubArgs 123 	'エラー
    Call testSubArgs(123) 	'正常

    testFunc 	    '正常
    Call testFunc 	'正常
    testFuncArgs 123 	'正常
    testFuncArgs(123) 	'正常
    Call testFuncArgs 123 	'エラー
    Call testFuncArgs(123) 	'正常

    Dim ret As Long
    ret = testFunc 	        '正常
    ret = Call testFunc 	'エラー
    ret = testFuncArgs 123 	'エラー
    ret = testFuncArgs(123) '正常
    ret = Call testFuncArgs 123     'エラー
    ret = Call testFuncArgs(123) 	'エラー
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