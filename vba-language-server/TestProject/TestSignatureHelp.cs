using VBACodeAnalysis;
using System;
using System.Collections.Generic;
using Xunit;

namespace TestProject {
    public class TestSignatureHelp {
		[Theory]
		[InlineData(" sub0(", 2, -1, -1, -1)]
		[InlineData(" sub0()", 7, -1, -1, -1)]
		[InlineData(" sub4(1, ", 2, -1, -1, -1)]
		[InlineData(" sub4(1, 2, 3, 4)", 17, -1, -1, -1)]

		[InlineData(" sub0(", 6, 2, 1, 0)]
		[InlineData(" sub0()", 6, 2, 1, 0)]
		[InlineData(" sub4(", 6, 2, 1, 0)]

		[InlineData(" sub4(1, ", 7, 2, 1, 0)]
		[InlineData(" sub4(1, 2, ", 8, 2, 1, 1)]
		[InlineData(" sub4(1, 2, ", 9, 2, 1, 1)]
		[InlineData(" sub4(1, 2, ", 10, 2, 1, 1)]
		[InlineData(" sub4(1, 2, 3, ", 11, 2, 1, 2)]
		[InlineData(" sub4(1, 2, 3, 4", 14, 2, 1, 3)]

		[InlineData(" sub4(1, )", 7, 2, 1, 0)]
		[InlineData(" sub4(1, 2, )", 8, 2, 1, 1)]
		[InlineData(" sub4(1, 2, )", 9, 2, 1, 1)]
		[InlineData(" sub4(1, 2, )", 10, 2, 1, 1)]
		[InlineData(" sub4(1, 2, 3, )", 11, 2, 1, 2)]
		[InlineData(" sub4(1, 2, 3, 4)", 14, 2, 1, 3)]
		public void TestGetSignaturePosition(string code, int codePos, int procLine, int procChara, int argPos) {
			var mod1 = @"Attribute VB_Name = ""module1""
Sub sub0()
End Sub
Sub sub4(a0 As Long, a1 As Long, a3 As Long, a4 As Long)
End Sub
";
			var modName = "mod2";
			var mod2 = $@"Attribute VB_Name = ""module2""
Sub Main()
{code}
End Sub
";
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			vbaca.setSetting(new RewriteSetting());
			vbaca.AddDocument("mod1", mod1);
			vbaca.AddDocument(modName, mod2);
			var actRet = 
				vbaca.GetSignaturePosition(modName, 2, codePos);
			Assert.Equal((procLine, procChara, argPos), actRet);
		}

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

		[Fact]
		public void TestProp() {
			var class1Name = "test_class1.cls";
			var class1Code = GetClassCode();
			var class2Name = "test_class2.cls";
			var class2Code = @"Public Class UseSigTest
Public ReadOnly Property SigTest As SigTest
End Class";
			var mod1Name = "test_module1.bas";
			var mod1Code = @"Public Module
Public Sub Main()
Dim test As New UseSigTest
test.SigTest(
End Sub
End Module";
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			vbaca.setSetting(new RewriteSetting());
			vbaca.AddDocument(class1Name, class1Code);
			vbaca.AddDocument(class2Name, class2Code);
			vbaca.AddDocument(mod1Name, mod1Code);
			var code = "test.SigTest(";
			var (procLine, procChara, argPosition) = vbaca.GetSignaturePosition(mod1Name, 3, code.Length);
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
