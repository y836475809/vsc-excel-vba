using VBACodeAnalysis;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace TestProject {
    public class TestRewriteProperty {
        [Fact]
        public void Test1() {
            var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
            var workspace = new AdhocWorkspace(host);
            var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), "MyProject", "MyProject", LanguageNames.VisualBasic).
            WithMetadataReferences(new[] {
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
            });
            var project = workspace.AddProject(projectInfo);

            var srcData = PropCode.getSrc();
            var doc = workspace.AddDocument(
                project.Id, "c1", SourceText.From(srcData));

            var rewriteProp = new RewriteProperty();
            var result = rewriteProp.Rewrite(doc.GetSyntaxRootAsync().Result);
            var rewriteText = result.root.GetText().ToString();

			var preData = PropCode.getPre();
			Helper.AssertCode(preData, rewriteText);

            var predict = new Dictionary<int, List<LocationDiff>> {
                {2, new List<LocationDiff>{ new LocationDiff(2, 7, 7) } },
                {5, new List<LocationDiff>{ new LocationDiff(5, 12, 3)} },
                {6, new List<LocationDiff>{ new LocationDiff(6, 12,  3)} },
                {9, new List<LocationDiff>{ new LocationDiff(9, 7, 1)} },
                {13, new List<LocationDiff>{ new LocationDiff(13, 7, 7)} },
                {16, new List<LocationDiff>{ new LocationDiff(16, 12, 3)} },
                {17, new List<LocationDiff>{ new LocationDiff(17, 12, 3)} },
                {20, new List<LocationDiff>{ new LocationDiff(20, 7, 1)} },
            };
            Helper.AssertLocationDiffDict(predict, result.dict);

            var preLinedict = new Dictionary<int,  int> {
                { 23, 2 },
                { 24, 13 }
            };
            preLinedict.All(x => rewriteProp.lineMappingDict.Contains(x));
            foreach (var item in preLinedict) {
                Assert.Equal(item.Value, rewriteProp.lineMappingDict[item.Key]);
            }
        }

        [Fact]
        public void TestNotAs() {
            var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
            var workspace = new AdhocWorkspace(host);
            var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), "MyProject", "MyProject", LanguageNames.VisualBasic).
            WithMetadataReferences(new[] {
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
            });
            var project = workspace.AddProject(projectInfo);

            var srcData = PropCode.GetSrcNotAs();
            var doc = workspace.AddDocument(
                project.Id, "c1", SourceText.From(srcData));

            var rewriteProp = new RewriteProperty();
            var result = rewriteProp.Rewrite(doc.GetSyntaxRootAsync().Result);
            var rewriteText = result.root.GetText().ToString();
            var preData = PropCode.GetPreNotAs();
            var prelines = preData.Split("\r\n");
            var actlines = rewriteText.Split("\r\n");
            for (int i = 0; i < prelines.Length; i++) {
                Assert.True(prelines[i] == actlines[i], $"{i}");
            }
        }

        [Fact]
        public void TestRename() {
            var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
            var workspace = new AdhocWorkspace(host);
            var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), "MyProject", "MyProject", LanguageNames.VisualBasic).
            WithMetadataReferences(new[] {
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
            });
            var project = workspace.AddProject(projectInfo);

            var srcData = PropCode.GetSrcAsignSamePart();
            var doc = workspace.AddDocument(
                project.Id, "c1", SourceText.From(srcData));

            var rewriteProp = new RewriteProperty();
            var result = rewriteProp.Rewrite(doc.GetSyntaxRootAsync().Result);
            var rewriteText = result.root.GetText().ToString();
            var preData = PropCode.GetPreAsignSamePart();
            var prelines = preData.Split("\r\n");
            var actlines = rewriteText.Split("\r\n");
            for (int i = 0; i < prelines.Length; i++) {
                Assert.True(prelines[i] == actlines[i], $"{i}");
            }
        }
    }
}
