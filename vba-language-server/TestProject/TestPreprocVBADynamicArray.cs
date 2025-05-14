using System;
using System.Collections.Generic;
using System.Linq;
using TestProject;
using VBACodeAnalysis;
using Xunit;

namespace TestPreprocVBADynamicArray {
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
		[Fact]
		public void TestReDimNotDim() {
			var code = @"
Sub sub1()
ReDim ary(2)
End Sub
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = string.Join("\r\n", [
				"",
				"Sub sub1()", 
				"Dim ary():ReDim ary(2)", 
				"End Sub",
				"",
			]);
			Helper.AssertCode(preCode, actCode);

			var cs = pp.GetColShift("test", 2, 0);
			Assert.Equal("Dim ary():".Length, cs);
		}

		[Fact]
		public void TestReDimNotDim2() {
			var code = @"
Sub sub1()
ReDim ary(2, 2)
End Sub
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = string.Join("\r\n", [
				"",
				"Sub sub1()",
				"Dim ary(,):ReDim ary(2, 2)",
				"End Sub",
				"",
			]);
			Helper.AssertCode(preCode, actCode);

			var cs = pp.GetColShift("test", 2, 0);
			Assert.Equal("Dim ary(,):".Length, cs);
		}

		[Fact]
		public void TestReDimAsNotDim() {
			var code = @"
Sub sub1()
ReDim ary(2) As Long
End Sub
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = string.Join("\r\n", [
				"",
				"Sub sub1()",
				$"Dim ary() As Long:ReDim ary(2) {new string(' ', "As Long".Length)}",
				"End Sub",
				"",
			]);
			Helper.AssertCode(preCode, actCode);

			var cs = pp.GetColShift("test", 2, 0);
			Assert.Equal("Dim ary() As Long:".Length, cs);
		}

		[Fact]
		public void TestReDimAsNotDim2() {
			var code = @"
Sub sub1()
ReDim ary(2, 3) As Long
End Sub
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = string.Join("\r\n", [
				"",
				"Sub sub1()",
				$"Dim ary(,) As Long:ReDim ary(2, 3) {new string(' ', "As Long".Length)}",
				"End Sub",
				"",
			]);
			Helper.AssertCode(preCode, actCode);

			var cs = pp.GetColShift("test", 2, 0);
			Assert.Equal("Dim ary(,) As Long:".Length, cs);
		}

		[Fact]
		public void TestReDimDimension() {
			var code = @"
Sub sub1()
Dim ary() As Long
ReDim ary(1 To 2, 1 To 3)
End Sub
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = string.Join("\r\n", [
				"",
				"Sub sub1()",
				"Dim ary(,) As Long",
				"ReDim ary(1 To 2, 1 To 3)",
				"End Sub",
				"",
			]);
			Helper.AssertCode(preCode, actCode);

			var cs = pp.GetColShift("test", 2, "Dim ary(,".Length);
			Assert.Equal(1, cs);
		}
	}
}
