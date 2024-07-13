using System;
using System.Collections.Generic;
using System.Linq;
using TestProject;
using VBACodeAnalysis;
using Xunit;

namespace TestPreprocVBA {
	using static System.Runtime.CompilerServices.RuntimeHelpers;
	using ColumnShiftDict = Dictionary<int, List<ColumnShift>>;
	using DiagoItemList = List<DiagnosticItem>;
	using LineReMapDict = Dictionary<int, int>;

	class TestPreprocVBA : PreprocVBA {
		public Dictionary<string, ColumnShiftDict> ColDict {
			get { return _fileColShiftDict; }
		}
		public Dictionary<string, LineReMapDict> LineDict {
			get { return _fileLineReMapDict; }
		}
	}

	public class TestRewriteVBA {
		[Fact]
		public void TestType() {
			var codeFmt = @"
Public {0} type_pub
    {1}num As Long
End {0}
Private {0} type_pri
    {1}num As Long
End {0}
{0} type_non
    {1}num As Long
End {0}
";
			var code = string.Format(codeFmt, "Type", "");
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var v = "Public ";
			var preCode = string.Format(codeFmt, "Structure", v);
			Helper.AssertCode(preCode, actCode);

			var preColDict = new ColumnShiftDict {
				{1, new List<ColumnShift>{ new(1, 12, 5) } },
				{2, new List<ColumnShift>{ new(2, 4, v.Length) } },

				{4, new List<ColumnShift>{ new(4, 13, 5) } },
				{5, new List<ColumnShift>{ new(5, 4, v.Length) } },

				{7, new List<ColumnShift>{ new(7, 5, 5) } },
				{8, new List<ColumnShift>{ new(8, 4, v.Length) } },
			};
			var actColDict = pp.ColDict["test"];
			Helper.AssertColumnShiftDict(preColDict, actColDict);
		}

		[Fact]
		public void TestTypeMember() {
			var codeFmt = @"
Public {0} type_pub
    {1}num1 As Long
    {1}num2(10) As Long
    {1}num3(1 to 5) As Long
    Public num4 As Long
    Private num5 As Long
End {0}
";
			var code = string.Format(codeFmt, "Type", "");
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var v = "Public ";
			var preCode = string.Format(codeFmt, "Structure", v);
			Helper.AssertCode(preCode, actCode);

			var preColDict = new ColumnShiftDict {
				{1, new List<ColumnShift>{ new(1, 12, 5) } },
				{2, new List<ColumnShift>{ new(2, 4, v.Length) } },
				{3, new List<ColumnShift>{ new(3, 4, v.Length) } },
				{4, new List<ColumnShift>{ new(4, 4, v.Length) } },
			};
			var actColDict = pp.ColDict["test"];
			Helper.AssertColumnShiftDict(preColDict, actColDict);
		}

		[Fact]
		public void TestProperty() {
			var code = @"Module TestType
Property Get Name1() As String
    name1 = 1
    Name1 = name1
    Set Name1 = name1
End Property
Property Let Name1(ByVal arg As String)
    name = arg
End Property

Property Get Name2()
    name1 = 1
    Name2 = name1
    Set Name2 = name1
End Property
Property Set Name2(ByVal arg As String)
    name = arg
End Property

Property Let Name3(ByVal arg As String)
    name = arg
End Property

End Module";
			var preCode = @"Module TestType
Private Function GetName1() As String
    name1 = 1
    GetName1 = name1
        GetName1 = name1
End Function
Private Sub SetName1(ByVal arg As String)
    name = arg
End Sub

Private Function GetName2()
    name1 = 1
    GetName2 = name1
        GetName2 = name1
End Function
Private Sub SetName2(ByVal arg As String)
    name = arg
End Sub

Private Sub SetName3(ByVal arg As String)
    name = arg
End Sub

Public Property Name1 As String
Public Property Name2 As String
Public Property Name3 As String
End Module";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			Helper.AssertCode(preCode, actCode);

			var preColDict = new ColumnShiftDict {
				{1, new List<ColumnShift>{ new(1, 13, 7) } },
				{3, new List<ColumnShift>{ new(3, 4, 3) } },
				{4, new List<ColumnShift>{ new(4, 8, 3) } },
				{6, new List<ColumnShift>{ new(6, 13, 2) } },
				{10, new List<ColumnShift>{ new(10, 13, 7) } },
				{12, new List<ColumnShift>{ new(12, 4, 3) } },
				{13, new List<ColumnShift>{ new(13, 8, 3) } },
				{15, new List<ColumnShift>{ new(15, 13, 2) } },
				{19, new List<ColumnShift>{ new(19, 13, 2) } },
			};
			var actColDict = pp.ColDict["test"];
			Helper.AssertColumnShiftDict(preColDict, actColDict);

			var preLineDict = new LineReMapDict {
				{23,  1},
				{24, 10 },
				{25, 19 },
			};
			var actLineDict = pp.LineDict["test"];
			Helper.AssertDict(preLineDict, actLineDict);
		}

		[Fact]
		public void TestLetSet() {
			var code = @"
Let a = 10
Set b = 10
Let a = 10:Set b = 10

Function func1() As int
    Let f1 = 10
    Set f1 = 10
End Function
Sub sub1() As int
    Let s1 = 10
    Set s2 = 10
End Sub

Property Get Name1() As String
    Let g1 = 10
    Set g2 = 10
End Property
Property Let Name1(n As String)
    Let l1 = 10
    Set l2 = 10
End Property
Property Set Name2(n As String)
    Let s1 = 10
    Set s2 = 10
End Property
";
			var preCode = @"
    a = 10
    b = 10
    a = 10:    b = 10

Function func1() As int
        f1 = 10
        f1 = 10
End Function
Sub sub1() As int
        s1 = 10
        s2 = 10
End Sub

Private Function GetName1() As String
        g1 = 10
        g2 = 10
End Function
Private Sub SetName1(n As String)
        l1 = 10
        l2 = 10
End Sub
Private Sub SetName2(n As String)
        s1 = 10
        s2 = 10
End Sub
Public Property Name1 As String
Public Property Name2 As String
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			Helper.AssertCode(preCode, actCode);

			var preColDict = new ColumnShiftDict {
				{14, new List<ColumnShift>{ new(14, 13, 7) } },
				{18, new List<ColumnShift>{ new(18, 13, 2) } },
				{22, new List<ColumnShift>{ new(22, 13, 2) } },
			};
			var actColDict = pp.ColDict["test"];
			Helper.AssertColumnShiftDict(preColDict, actColDict);

			var preLineDict = new LineReMapDict {
				{26, 14},
				{27, 22 },
			};
			var actLineDict = pp.LineDict["test"];
			Helper.AssertDict(preLineDict, actLineDict);
		}

		[Fact]
		public void TestFileNumber() {
			var codeFmt = @"
date = #1/2/3#

Dim d#
d = 1.0
d = 2.0#/10
d = 3.0#/ 10
d = 4.0# / 10
d = 5.0 # / 10

a = {0}1  
a = {0}10

n = 1
b = {0}n
a = {0}1:b = {0}10
a = {0}2 : b={0}10
a = {0}3:b={0}10:c=2
a = {0}4:b={0}n

Function func1() As int
    f1 = {0}10
    f1 = {0}n
End Function
Sub sub1() As int
    s1 = {0}10
    s1 = {0}n
End Sub
";
			var code = string.Format(codeFmt, "#");
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = string.Format(codeFmt, " ");
			Helper.AssertCode(preCode, actCode);

			var actColDict = pp.ColDict["test"];
			Assert.Empty(actColDict);
		}

		[Fact]
		public void TestFileNumberInProperty() {
			var code = @"
Property Get Name1() As String
    g1 = #10
    g1 = #n
    Name1 = #10
    Name1 = #n
End Property
Property Let Name1(n As String)
    l1 = #10
    l1 = #n
End Property
Property Set Name2(n As String)
    s1 = #10
    s1 = #n
End Property
";
			var preCode = @"
Private Function GetName1() As String
    g1 =  10
    g1 =  n
    GetName1 =  10
    GetName1 =  n
End Function
Private Sub SetName1(n As String)
    l1 =  10
    l1 =  n
End Sub
Private Sub SetName2(n As String)
    s1 =  10
    s1 =  n
End Sub
Public Property Name1 As String
Public Property Name2 As String
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			Helper.AssertCode(preCode, actCode);
		}

		[Fact]
		public void TestVariant() {
			var codeFmt = @"
Dim a As {0}
Function func1(arg As {0}) As {0}
    Dim f1 As {0}
End Function
Sub sub1() As {0}
    Dim s1 As {0}
End Sub
";
			var code = string.Format(codeFmt, "Variant");
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = string.Format(codeFmt, "Object ");
			Helper.AssertCode(preCode, actCode);

			var actColDict = pp.ColDict["test"];
			Assert.Empty(actColDict);
		}

		[Fact]
		public void TestVariantInType() {
			var codeFmt = @"
Public {0} type_pub
    {2}num As {1}
End {0}
Private {0} type_pri
    {2}num As {1}
End {0}
{0} type_non
    {2}num As {1}
End {0}
";
			var code = string.Format(codeFmt, "Type", "Variant", "");
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var v = "Public ";
			var preCode = string.Format(codeFmt, "Structure", "Object ", v);
			Helper.AssertCode(preCode, actCode);

			var preColDict = new ColumnShiftDict {
				{1, new List<ColumnShift>{ new(1, 12, 5) } },
				{2, new List<ColumnShift>{ new(2, 4, v.Length) } },

				{4, new List<ColumnShift>{ new(4, 13, 5) } },
				{5, new List<ColumnShift>{ new(5, 4, v.Length) } },

				{7, new List<ColumnShift>{ new(7, 5, 5) } },
				{8, new List<ColumnShift>{ new(8, 4, v.Length) } },
			};
			var actColDict = pp.ColDict["test"];
			Helper.AssertColumnShiftDict(preColDict, actColDict);
		}

		[Fact]
		public void TestVariantInProperty() {
			var code = @"
Property Get Name1() As Variant
    Dim g1 As Variant
End Property
Property Let Name1(n As Variant)
    Dim l1 As Variant
End Property
Property Set Name2(n As Variant)
    Dim s1 As Variant
End Property
";
			var preCode = @"
Private Function GetName1() As Object 
    Dim g1 As Object 
End Function
Private Sub SetName1(n As Object )
    Dim l1 As Object 
End Sub
Private Sub SetName2(n As Object )
    Dim s1 As Object 
End Sub
Public Property Name1 As Object
Public Property Name2 As Object
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			Helper.AssertCode(preCode, actCode);

			var preColDict = new ColumnShiftDict {
				{1, new List<ColumnShift>{ new(1, 13, 7) } },
				{4, new List<ColumnShift>{ new(4, 13, 2) } },
				{7, new List<ColumnShift>{ new(7, 13, 2) } },
			};
			var actColDict = pp.ColDict["test"];
			Helper.AssertColumnShiftDict(preColDict, actColDict);

			var preLineDict = new LineReMapDict {
				{10, 1},
				{11, 7 },
			};
			var actLineDict = pp.LineDict["test"];
			Helper.AssertDict(preLineDict, actLineDict);
		}

		[Fact]
		public void TestVBAFunction() {
			var codeFmt = @"
a = {0}Range
a = {0}Range(0, 1)
a = {0}Range(0, 1).Cells

For Each v In {0}Range
Next
For Each v In {0}Range(0, 1).Cells
Next

ActiveSheet.Range(2,2) = 100
ActiveSheet. _ 
Range(""B2"")  = 100

ActiveSheet.Cells(2,2) = 100
{0}Cells(1,3) = 100
{0}Range({0}Cells(2, 1), {0}Cells(lastRow, 1)) = ""

Dim a As Range
Dim a As New Range
a = New Range
";
			var code = string.Format(codeFmt, "");
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = string.Format(codeFmt, "f.");
			Helper.AssertCode(preCode, actCode);

			var preColDict = new ColumnShiftDict {
				{1, new List<ColumnShift>{ new(1, 4, 2) } },
				{2, new List<ColumnShift>{ new(2, 4, 2) } },
				{3, new List<ColumnShift>{ new(3, 4, 2) } },
				{5, new List<ColumnShift>{ new(5, 14, 2) } },
				{7, new List<ColumnShift>{ new(7, 14, 2) } },
				{15, new List<ColumnShift>{ new(15, 0, 2) } },
				{16, new List<ColumnShift>{
					new(16, 0, 2),
					new(16, 6, 2),
					new(16, 19, 2),
				} },
			};
			var actColDict = pp.ColDict["test"];
			Helper.AssertColumnShiftDict(preColDict, actColDict);
		}

		[Fact]
		public void TestPredefined() {
			var codeFmt = @"
Dim a As String: a = {0}CStr(10)
a = {0}CStr(10)
a = {0}CStr
a = func(0, 1)
Call func(0, 1)
Call sub(0, 1)
";
			var code = string.Format(codeFmt, "");
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preText = "ExcelVBAFunctions.";
			var preCode = string.Format(codeFmt, preText);
			Helper.AssertCode(preCode, actCode);

			var colshift = preText.Length;
			var preColDict = new ColumnShiftDict {
				{1, new List<ColumnShift>{ new(1, 21, colshift) } },
				{2, new List<ColumnShift>{ new(2, 4, colshift) } },
				{3, new List<ColumnShift>{ new(3, 4, colshift) } },
			};
			var actColDict = pp.ColDict["test"];
			Helper.AssertColumnShiftDict(preColDict, actColDict);
		}

		[Fact]
		public void TestModuleHeaderClass() {
			var codeFmt = @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""Person""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
{0}

Dim a As long";
			{
				var code = string.Format(codeFmt, "Option Explicit");
				var pp = new PreprocVBA();
				var actCode = pp.Rewrite("test", code);
				var preCode = $@"Option Explicit On
Public Class Person
{string.Concat(Enumerable.Repeat("\r\n", 6))}


Dim a As long
End Class";
				Helper.AssertCode(preCode, actCode);
			}
			{
				var code = string.Format(codeFmt, "");
				var pp = new PreprocVBA();
				var actCode = pp.Rewrite("test", code);
				var preCode = $@"Public Class Person
{string.Concat(Enumerable.Repeat("\r\n", 7))}


Dim a As long
End Class";
				Helper.AssertCode(preCode, actCode);
			}
		}

		[Fact]
		public void TestModuleHeaderBas() {
			var codeFmt = @"Attribute VB_Name = ""テスト""
{0}
Dim a As long";
			{
				var code = string.Format(codeFmt, "Option Explicit");
				var pp = new PreprocVBA();
				var actCode = pp.Rewrite("test", code);
				var preCode = $@"Option Explicit On
Public Module テスト
Dim a As long
End Module";
				Helper.AssertCode(preCode, actCode);
			}
			{
				var code = string.Format(codeFmt, "");
				var pp = new PreprocVBA();
				var actCode = pp.Rewrite("test", code);
				var preCode = $@"Public Module テスト

Dim a As long
End Module";
				Helper.AssertCode(preCode, actCode);
			}
		}

		[Theory]
		[InlineData("Open fp1 For Output As{0}1")]
		[InlineData("Open fp2 For Output Access Read As{0}1")]
		[InlineData("Open fp3 For Output Access Read Shared As{0}1")]
		[InlineData("Open fp4 For Output Access Read Write Lock Read As{0}1")]
		[InlineData("Open fp5 For Output Access Read Lock Read As{0}1")]
		[InlineData("Open fp6 For Output Access Read Write Shared As{0}1")]
		public void TestOpen1(string code) {
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", string.Format(code, "  #"));
			var preCode = string.Format(code, "  ,");
			Helper.AssertCode(preCode, actCode);
		}

		[Theory]
		[InlineData("Open fp1 For Output As{0}fn")]
		[InlineData("Open fp2 For Output Access Read As{0}fn")]
		[InlineData("Open fp3 For Output Access Read Shared As{0}fn")]
		[InlineData("Open fp4 For Output Access Read Write Lock Read As{0}fn")]
		[InlineData("Open fp5 For Output Access Read Lock Read As{0}fn")]
		[InlineData("Open fp6 For Output Access Read Write Shared As{0}fn")]
		public void TestOpen2(string code) {
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", string.Format(code, "  "));
			var preCode = string.Format(code, " ,");
			Helper.AssertCode(preCode, actCode);
		}
	}
}
