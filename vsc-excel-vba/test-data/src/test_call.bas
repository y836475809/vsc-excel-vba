Attribute VB_Name = "test_call"
Option Explicit

Sub Main()
    testSub 	    '正常
    Call testSub 	'正常
    testSubArg 123 	'正常
    testSubArg(123) 	'正常
    Call testSubArg 123 	'エラー
    Call testSubArg(123) 	'正常

    testFunc 	    '正常
    Call testFunc 	'正常
    testFuncArg 123 	'正常
    testFuncArg(123) 	'正常
    Call testFuncArg 123 	'エラー
    Call testFuncArg(123) 	'正常

    Dim ret As Long
    ret = testFunc 	        '正常
    ret = Call testFunc 	'エラー
    ret = testFuncArg 123 	'エラー
    ret = testFuncArg(123) '正常
    ret = Call testFuncArg 123     'エラー
    ret = Call testFuncArg(123) 	'エラー
End Sub

Sub MainArgs()
    testSubArg2 1, 2       '正常
    testSubArg2(1, 2)      'エラー
    Call testSubArg2 1, 2  'エラー
    Call testSubArg2(1, 2) '正常

    testFuncArg2 1, 2        '正常
    testFuncArg2 (1,2)       'エラー
    Call testFuncArg2 1,2    'エラー
    Call testFuncArg2(1, 2)  '正常

    Dim ret As Long
    ret = testFuncArg2 1,2       'エラー
    ret = testFuncArg2(1, 2)     '正常
    ret = Call testFuncArg2 1,2  'エラー
    ret = Call testFuncArg2(1,2) 'エラー
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