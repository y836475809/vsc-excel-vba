using System.Collections.Generic;
using System.Linq;
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
			var vbCode = vbaca.Rewrite(name, vbacode);
			vbaca.AddDocument(name, vbCode);
			var diagnostics = await vbaca.GetDiagnostics(name);
			Assert.Empty( diagnostics );
		}

		[Fact]
		public async Task TestDiagnostics() {
			var name = "test";
			var vbacode = Helper.getCode($"test_property_diagnostics2.bas");
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			var vbCode = vbaca.Rewrite(name, vbacode);
			vbaca.AddDocument(name, vbCode);
			var diagnostics = await vbaca.GetDiagnostics(name);
			
			Assert.Equal(4, diagnostics.Count);
			var actIdLineList = diagnostics.Select(x => (x.ID, x.Start.Item1)).ToList();
			Assert.True(("BC30205", 2).Equals(actIdLineList[0]));
			Assert.True(("BC30188", 3).Equals(actIdLineList[1]));
			Assert.True(("BC30431", 4).Equals(actIdLineList[2]));
			Assert.True(("BC30002", 15).Equals(actIdLineList[3]));
		}
	}
}
