using System;
using VBACodeAnalysis;
using Xunit;

namespace TestProject {
	public class TestRewriteVBALetSet() {
		[Fact]
		public void TestRewrite() {
			var code = Helper.getCode("test_letset1.bas");
			var rewriter = new TestVBARewriter();
			var actCode = rewriter.Rewrite("test", code);
			var expCode = Helper.getCode($"test_letset1_exp.bas");
			Helper.AssertCode(expCode, actCode);

			var expColDict = new ColumnShiftDict {
					{14, new (){ new(14, 13, -4) } },
					{19, new (){ new(19, 13, 2) } },
					{24, new (){ new(24, 13, 2) } },
				};
			var actColDict = rewriter.ColDict("test");
			Helper.AssertColumnShiftDict(expColDict, actColDict);

			var expLineMapDict = new LineMapDict {
					{ 28, 24 }
				};
			var actLineDict = rewriter.LineMapDict("test");
			Helper.AssertDict(actLineDict, expLineMapDict);
		}
	}
}
