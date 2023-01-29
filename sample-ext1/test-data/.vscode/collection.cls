VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Count As Long

Public Sub Add(item As Object)
End Sub
Public Sub Add(item As Object, key As String)
End Sub

Public Sub Remove(index As Long)
End Sub
Public Sub Remove(key As String)
End Sub

Public Property Item(index As Integer) As Object
End Property
Public Property Item(key As String) As Object
End Property