
namespace TestProject1 {
	class PropCode {
		public static string getSrc() {
			return @$"Public Class C1
Private Name As String
Property Get Name1() As String
    Dim pp As Long
    pp = 200
    Name1 = LCase(Me.Name)
    Name1 = LCase(Me.Name)
End Property

  Public Property Let Name1(arg1 As String)
    Me.Name = arg1
End Property
'ppp
  Property Get Name2() As String
    Dim pp As Long
    pp = 200
    Name2 = LCase(Me.Name)
    Name2 = LCase(Me.Name)
End Property

Public Property Let Name2(arg1 As String)
    Me.Name = arg1
End Property
End Class";
		}
		public static string getPre() {
                return @$"Public Class C1
Private Name As String
Private Function  getName1() As String
    Dim pp As Long
    pp = 200
    getName1 = LCase(Me.Name)
    getName1 = LCase(Me.Name)
End Function

  Private Sub       letName1(arg1 As String)
    Me.Name = arg1
End Sub     
'ppp
Private   Function  getName2() As String
    Dim pp As Long
    pp = 200
    getName2 = LCase(Me.Name)
    getName2 = LCase(Me.Name)
End Function

Private Sub       letName2(arg1 As String)
    Me.Name = arg1
End Sub     
Public Property Name1 As String
Public Property Name2 As String
End Class";
        }

        public static string GetSrcNotAs() {
            return @$"Public Class C1
Private Name As String
Property Get Name1()
End Property
Public Property Let Name1(arg1 As String)
End Property
End Class";
        }
        public static string GetPreNotAs() {
            return @$"Public Class C1
Private Name As String
Private Function  getName1()
End Function
Private Sub       letName1(arg1 As String)
End Sub     
Public Property Name1
End Class";
        }
    }
}
