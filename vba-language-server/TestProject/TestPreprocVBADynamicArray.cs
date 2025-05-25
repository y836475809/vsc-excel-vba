using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using TestProject;
using VBACodeAnalysis;
using Xunit;

namespace TestPreprocVBADynamicArray {
	using static Antlr4.Runtime.Atn.SemanticContext;
	using static System.Runtime.InteropServices.JavaScript.JSType;
	using ColumnShiftDict = Dictionary<int, List<ColumnShift>>;
	using LineReMapDict = Dictionary<int, int>;

	class TestPreprocVBA : PreprocVBA {
		public Dictionary<string, ColumnShiftDict> ColDict {
			get { return _fileColShiftDict; }
		}
		public Dictionary<string, LineReMapDict> LineDict {
			get { return _fileLineReMapDict; }
		}
	}

	public class TestRewriteDynamicArray {
		private string GetCode(string[] lines) {
			var code = string.Join("\r\n", lines);
			return $"\r\n{code}\r\n";
		}

		[Fact]
		public void TestReDimNotDim() {
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", GetCode([
				"Sub sub1()",
				"ReDim ary(2)",
				"End Sub"
			]));
			var preCode = GetCode([
				"Sub sub1()", 
				"Dim ary():ReDim ary(2)", 
				"End Sub"
			]);
			Helper.AssertCode(preCode, actCode);

			var cs = pp.GetColShift("test", 2, 0);
			Assert.Equal("Dim ary():".Length, cs);
		}

		[Fact]
		public void TestReDimNotDim2() {
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", GetCode([
				"Sub sub1()",
				"ReDim ary(2, 2)",
				"End Sub"
			]));
			var preCode = GetCode([
				"Sub sub1()",
				"Dim ary(,):ReDim ary(2, 2)",
				"End Sub"
			]);
			Helper.AssertCode(preCode, actCode);

			var cs = pp.GetColShift("test", 2, 0);
			Assert.Equal("Dim ary(,):".Length, cs);
		}

		[Fact]
		public void TestReDimAsNotDim() {
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", GetCode([
				"Sub sub1()",
				"ReDim ary(2) As Long",
				"End Sub"
			]));
			var preCode = GetCode([
				"Sub sub1()",
				$"Dim ary() As Long:ReDim ary(2) {new string(' ', "As Long".Length)}",
				"End Sub"
			]);
			Helper.AssertCode(preCode, actCode);

			var cs = pp.GetColShift("test", 2, 0);
			Assert.Equal("Dim ary() As Long:".Length, cs);
		}

		[Fact]
		public void TestReDimAsNotDim2() {
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", GetCode([
				"Sub sub1()",
				"ReDim ary(2, 3) As Long",
				"End Sub"
			]));
			var preCode = GetCode([
				"Sub sub1()",
				$"Dim ary(,) As Long:ReDim ary(2, 3) {new string(' ', "As Long".Length)}",
				"End Sub"
			]);
			Helper.AssertCode(preCode, actCode);

			var cs = pp.GetColShift("test", 2, 0);
			Assert.Equal("Dim ary(,) As Long:".Length, cs);
		}

		[Fact]
		public void TestReDimDimension() {
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", GetCode([
				"Sub sub1()",
				"Dim ary() As Long",
				"ReDim ary(1 To 2, 1 To 3)",
				"End Sub"
			]));
			var preCode = GetCode([
				"Sub sub1()",
				"Dim ary(,) As Long",
				"ReDim ary(1 To 2, 1 To 3)",
				"End Sub"
			]);
			Helper.AssertCode(preCode, actCode);

			var cs = pp.GetColShift("test", 2, "Dim ary(,".Length);
			Assert.Equal(1, cs);
		}

		[Fact]
		public void TestFieldReDimDimension() {
			var preprocVBA = new TestPreprocVBA();
			var actCode = preprocVBA.Rewrite("test", Helper.getCode("test_dynamic_array1.bas"));
			var expCode = Helper.getCode($"test_dynamic_array1_exp.bas");
			Helper.AssertCode(expCode, actCode);

			var expColDict = new ColumnShiftDict {
					{1, new (){ new(1, 8, 1) } },
					{2, new (){ new(2, 8, 1) } },

					{9, new (){ new(9, 11, 1) } },
					{11, new (){ new(11, 1, 21) } },

					{19, new (){ new(19, 12, 1) } },
					{21, new (){ new(21, 1, 22) } },
				};
			var actColDict = preprocVBA.ColDict["test"];
			Helper.AssertColumnShiftDict(expColDict, actColDict);
		}
	}
}
