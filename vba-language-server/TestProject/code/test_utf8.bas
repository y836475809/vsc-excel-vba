Attribute VB_Name = "test_utf8"
Option Explicit On

'接続
Sub DataBase_Sample()
    '接続
    Set obj = CreateObject("")
    '接続
    With obj
        .Prop1 = "1" '接続
        .Prop2 = "1" '接続
        .Prop3 = "1" '接続
        .Prop4 = "1" '接続
        .Prop5 = "1"
    End With

    '接続
    text = "text"
    Set obj = Nothing
End Sub
