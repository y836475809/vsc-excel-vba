using VBACodeAnalysis;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using System.Collections.Generic;
using System.Linq;
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
            setting.NameSpace = "f";
            setting.Mapping.Add(new List<string> { "Cells", "Cls" });
            setting.Mapping.Add(new List<string> { "Range", "Ran" });
            setting.convert();
            rewrite = new Rewrite(setting);
        }

        Document AddDoc(string name, string code) {
            var doc = workspace.AddDocument(
                project.Id, name, SourceText.From(code));
            workspace.TryApplyChanges(
                workspace.CurrentSolution.WithDocumentName(doc.Id, name));
            return doc;
        }

        string GetText(SyntaxNode root, List<TextChange> changes) {
            var changed = root.SyntaxTree
                .WithChangedText(root.GetText().WithChanges(changes))
                .GetRootAsync().Result;
            return changed.GetText().ToString();
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
For Each v In f.Ran
Next
f.Cls(1,2)
f.Cls.Test
f.Ran(""A1"")
f.Ran.Test
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
            var changes = rewrite.VBAClassToFunction(doc, docRoot.DescendantNodes());
            var actCode = GetText(docRoot, changes);
            Helper.AssertCode(preCode, actCode);
        }
    }
}
