using System.Collections.Generic;
using System.Threading.Tasks;
using VBACodeAnalysis;
using Xunit;


namespace TestProject {
	public class TestPropertyDiagnostics() {
		[Fact]
		public async Task TestNonDiagnostics() {
			var name = "test";
			var vbacode = Helper.getCode($"test_property_diagnostics1.bas");
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			vbaca.setSetting(new RewriteSetting());
			var vbCode = vbaca.Rewrite(name, vbacode);
			vbaca.AddDocument(name, vbCode);
			var diagnostics = await vbaca.GetDiagnostics(name);
			Assert.Empty( diagnostics );
		}
	}
}
