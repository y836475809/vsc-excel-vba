using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using System.Collections.Generic;
using System.Linq;
using VBACodeAnalysis;
using Xunit;

namespace TestProject {
	public class TestRewriteVBAClassToFunction {
		AdhocWorkspace workspace;
		Project project;
		Rewrite rewrite;

		private void Setup() {
			var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
			workspace = new AdhocWorkspace(host);
			var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), "MyProject", "MyProject", LanguageNames.VisualBasic).
			WithMetadataReferences(new[] {
				MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
			});
			project = workspace.AddProject(projectInfo);

			var setting = new RewriteSetting();
			var settingVBA = setting.VBAClassToFunction;
			settingVBA.ModuleName = "f";
			settingVBA.VBAClasses.Add("cells");
			settingVBA.VBAClasses.Add("range");
			rewrite = new Rewrite(setting);
		}

		Document AddDoc(string name, string code) {
			var doc = workspace.AddDocument(
				project.Id, name, SourceText.From(code));
			workspace.TryApplyChanges(
				workspace.CurrentSolution.WithDocumentName(doc.Id, name));
			return doc;
		}

		[Fact]
		public void TestClassForEahc() {
			Setup();
			var cellsClass = $@"Public Class Cells
End Class";
			var rangeClass = $@"Public Class Range
Public Sub Test()
End Sub
End Class";
			var srcCode = $@"Module Module1
Sub Main()
Dim c As Cells
c.Test
Dim r As Range
r.Test
Dim v As Variant
For Each v In Range
Next
For Each v In Range(1)
Next
For Each v In Range.Cells
Next
For Each v In Range(1).Cells
Next
Cells(1,2)
Cells.Test
Range(""A1"")
Range.Test
' Cells(1,2)
' Cells.Test
' Range(""A1"")
' Range.Test
End Sub
End Module";
			var preCode = $@"Module Module1
Sub Main()
Dim c As Cells
c.Test
Dim r As Range
r.Test
Dim v As Variant
For Each v In f.Range
Next
For Each v In f.Range(1)
Next
For Each v In f.Range.Cells
Next
For Each v In f.Range(1).Cells
Next
f.Cells(1,2)
f.Cells.Test
f.Range(""A1"")
f.Range.Test
' Cells(1,2)
' Cells.Test
' Range(""A1"")
' Range.Test
End Sub
End Module";
			AddDoc("cells", cellsClass);
			AddDoc("range", rangeClass);
			var doc = AddDoc("srcCode", srcCode);
			var docRoot = doc.GetSyntaxRootAsync().Result;
			var (root, dict) = rewrite.VBAClassToFunction(docRoot);
			var actCode = root.GetText().ToString();
			Helper.AssertCode(preCode, actCode);

			var predict = new Dictionary<int, List<LocationDiff>> {
				{7, new List<LocationDiff>{ new LocationDiff(7, 14, 2)} },
				{9, new List<LocationDiff>{ new LocationDiff(9, 14, 2) } },
				{11, new List<LocationDiff>{ new LocationDiff(11, 14, 2) } },
				{13, new List<LocationDiff>{ new LocationDiff(13, 14, 2) } },
				{15, new List<LocationDiff>{ new LocationDiff(15, 0, 2)} },
				{16, new List<LocationDiff>{ new LocationDiff(16, 0, 2)} },
				{17, new List<LocationDiff>{ new LocationDiff(17, 0, 2)} },
				{18, new List<LocationDiff>{ new LocationDiff(18, 0, 2)} },
			};
			Helper.AssertLocationDiffDict(predict, dict);
		}
	}
}
