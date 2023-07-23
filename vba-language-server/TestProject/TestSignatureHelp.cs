using VBACodeAnalysis;
using System;
using System.Collections.Generic;
using Xunit;

namespace TestProject {
    public class TestSignatureHelp {
        [Fact]
        public async void TestMethod()
        {
            var class1Name = "test_class1.cls";
            var class1Code = @"Public Class SigTest
    Public Sub Add(Key As String, Item As Object)
    End Sub
    Public Sub Add(Key As Long, Item As Object)
    End Sub
    Public Function Add(Key As Long) As Object
    End Function
End Class";
            var mod1Name = "test_module1.bas";
            var mod1Code = @"Public Module
Public Sub Main()
Dim t As New SigTest
t.Add(
End Sub
End Module";
            var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
            vbaca.setSetting(new RewriteSetting());
            vbaca.AddDocument(class1Name, class1Code);
            vbaca.AddDocument(mod1Name, mod1Code);
			var (procLine, procChara, argPosition) = vbaca.GetSignaturePosition(mod1Name, 3, "t.Add(".Length);
			var act = vbaca.GetSignatureHelp(mod1Name, procLine, procChara).Result;
			foreach (var item in act) {
				item.ActiveParameter = argPosition;
			}
			var pre = new List<SignatureHelpItem>() {
				new SignatureHelpItem {
					ActiveParameter = 0,
					DisplayText = "Public Sub Add(Key As String, Item As Variant)",
					Description = "",
					ReturnType = "Void",
					Args = new List<ArgumentItem> {
						new ArgumentItem("Key", "String"),
						new ArgumentItem("Item", "Variant"),
					}
				},
				new SignatureHelpItem {
					ActiveParameter = 0,
					DisplayText = "Public Sub Add(Key As Long, Item As Variant)",
					Description = "",
					ReturnType = "Void",
					Args = new List<ArgumentItem> {
						new ArgumentItem("Key", "Long"),
						new ArgumentItem("Item", "Variant"),
					}
				},
				new SignatureHelpItem {
					ActiveParameter = 0,
					DisplayText = "Public Function Add(Key As Long) As Variant",
					Description = "",
					ReturnType = "Variant",
					Args = new List<ArgumentItem> {
						new ArgumentItem("Key", "Long")
					}
				}
			};
			Helper.AssertSignatureHelp(pre, act);
		}
    }
}
