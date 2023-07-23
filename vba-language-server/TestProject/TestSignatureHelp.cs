using VBACodeAnalysis;
using System;
using System.Collections.Generic;
using Xunit;

namespace TestProject {
    public class TestSignatureHelp {
		private string GetClassCode() {
			return @"Public Class SigTest
    Default Public Property Item(Index As Long) As Object
        Get : End Get
        Set(Value As Object) : End Set
    End Property
    Default Public Property Item(Key As String) As Object
        Get : End Get
        Set(Value As Object) : End Set
    End Property

    Public Sub Add(Key As String, Item As Object)
    End Sub
    Public Sub Add(Key As Long, Item As Object)
    End Sub
    Public Function Add(Key As Long) As Object
    End Function
End Class";
		}
		private string GetTestCode(string code) {
			return $@"Public Module
Private fv As New SigTest
Public Sub Main()
Private lv As New SigTest
Dim t As New SigTest
{code}
End Sub
End Module";
		}
		
		[Fact]
		public void TestMethod() {
			var class1Name = "test_class1.cls";
			var class1Code = GetClassCode();
			var mod1Name = "test_module1.bas";
			var code = "t.Add(";
			var mod1Code = GetTestCode(code);
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			vbaca.setSetting(new RewriteSetting());
			vbaca.AddDocument(class1Name, class1Code);
			vbaca.AddDocument(mod1Name, mod1Code);
			var (procLine, procChara, argPosition) = vbaca.GetSignaturePosition(mod1Name, 5, code.Length);
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

		[Theory]
		[InlineData("fv(")]
		[InlineData("lv(")]
		public void TestFiledLocal(string code) {
			var class1Name = "test_class1.cls";
			var class1Code = GetClassCode();
			var mod1Name = "test_module1.bas";
			//var code = "fv(";
			var mod1Code = GetTestCode(code);
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			vbaca.setSetting(new RewriteSetting());
			vbaca.AddDocument(class1Name, class1Code);
			vbaca.AddDocument(mod1Name, mod1Code);
			var (procLine, procChara, argPosition) = vbaca.GetSignaturePosition(mod1Name, 5, code.Length);
			var act = vbaca.GetSignatureHelp(mod1Name, procLine, procChara).Result;
			foreach (var item in act) {
				item.ActiveParameter = argPosition;
			}
			var pre = new List<SignatureHelpItem>() {
				new SignatureHelpItem {
					ActiveParameter = 0,
					DisplayText = "SigTest(Index As Long) As Variant",
					Description = "",
					ReturnType = "Variant",
					Args = new List<ArgumentItem> {
						new ArgumentItem("Index", "Long")
					}
				},
				new SignatureHelpItem {
					ActiveParameter = 0,
					DisplayText = "SigTest(Key As String) As Variant",
					Description = "",
					ReturnType = "Variant",
					Args = new List<ArgumentItem> {
						new ArgumentItem("Key", "String"),
					}
				}
			};
			Helper.AssertSignatureHelp(pre, act);
		}
	}
}
