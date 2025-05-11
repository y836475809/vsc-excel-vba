using System;
using System.Collections.Generic;
using System.Linq;
using TestProject;
using VBACodeAnalysis;
using Xunit;

namespace TestPreprocVBAReDim {
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
	public class TestRewriteVBA {
		[Fact]
		public void TestReDimNotDim() {
			var code = @"
ReDim ary(2)
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = "\r\nDim ary() : ReDim ary(2)\r\n";
			Helper.AssertCode(preCode, actCode);
		}

		[Fact]
		public void TestReDimNotDim2() {
			var code = @"
ReDim ary(2, 2)
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = "\r\nDim ary(,) : ReDim ary(2, 2)\r\n";
			Helper.AssertCode(preCode, actCode);
		}

		[Fact]
		public void TestReDimAsNotDim() {
			var code = @"
ReDim ary(2) As Long
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = "\r\nDim ary() As Long: ReDim ary(2)\r\n";
			Helper.AssertCode(preCode, actCode);
		}

		[Fact]
		public void TestReDimAsNotDim2() {
			var code = @"
ReDim ary(2, 3) As Long
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = "\r\nDim ary(,) As Long: ReDim ary(2, 3)\r\n";
			Helper.AssertCode(preCode, actCode);
		}

		[Fact]
		public void TestReDimDimension() {
			var code = @"
Dim ary() As Long
ReDim ary(1 To 2, 1 To 3)
";
			var pp = new TestPreprocVBA();
			var actCode = pp.Rewrite("test", code);
			var preCode = "\r\nDim ary(,)\r\nReDim ary(0 To 2, 0 To 3)\r\n";
			Helper.AssertCode(preCode, actCode);
		}
	}
}
