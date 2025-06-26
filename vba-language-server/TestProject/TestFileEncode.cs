using System;
using VBARewrite;
using Xunit;

namespace TestProject {
	public class TestFileEncode() {
		[Fact]
		public void TestRewriteUTF8VBACode() {
			var code = Helper.getCode("test_utf8.bas");
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			var rewriter = new VBARewriter();
			Assert.Throws<ArgumentOutOfRangeException>(() => { rewriter.Rewrite("test", code); });
		}
	}
}
