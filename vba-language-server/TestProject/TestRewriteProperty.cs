using VBACodeAnalysis;
using Xunit;

namespace TestProject {
	public class TestRewriteProperty() {
		[Theory]
		[InlineData("1")]
		[InlineData("2")]
		public void TestRewrite(string codeId) {
			var code = Helper.getCode($"test_property{codeId}.bas");
			var preprocVBA = new PreprocVBA();
			var actCode = preprocVBA.Rewrite("test", code);
			var expCode = Helper.getCode($"test_property{codeId}_exp.bas");
			Helper.AssertCode(expCode, actCode);
		}
	}
}
