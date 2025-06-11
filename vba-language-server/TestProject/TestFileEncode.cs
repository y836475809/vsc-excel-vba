using System;
using VBACodeAnalysis;
using Xunit;

namespace TestProject {
	public class TestFileEncode() {
		[Fact]
		public void TestRewriteVBACode() {
			var code = Helper.getCode("test_utf8.bas");
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			var preprocVBA = new PreprocVBA();
			Assert.Throws<ArgumentOutOfRangeException>(() => { preprocVBA.Rewrite("test", code); });
		}
	}
}
