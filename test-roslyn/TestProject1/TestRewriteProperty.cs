using ConsoleApp1;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace TestProject1 {
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
            var docRoot = rewriteProp.Rewrite(doc.GetSyntaxRootAsync().Result);
            var rewriteText = docRoot.GetText().ToString();
            var preData = PropCode.getPre();
            var prelines = preData.Split("\r\n");
            var actlines = rewriteText.Split("\r\n");
            for (int i = 0; i < prelines.Length; i++) {
                Assert.True(prelines[i] == actlines[i], $"{i}");
            }
            var predict = new Dictionary<int, (int, int)> {
                {2, (7, 7)},
                {5, (12, -3)},
                {6, (12, -3)},
                {9, (7, -1)},
                {13, (7, 7)},
                {16, (12, -3)},
                {17, (12, -3)},
                {20, (7, -1)},
            };
            predict.All(x => rewriteProp.charaOffsetDict.Contains(x));
            foreach (var item in predict) {
                Assert.Equal(item.Value, rewriteProp.charaOffsetDict[item.Key]);
            }

            var preLinedict = new Dictionary<int,  int> {
                { 23, 2 },
                { 24, 13 }
            };
            preLinedict.All(x => rewriteProp.lineMappingDict.Contains(x));
            foreach (var item in preLinedict) {
                Assert.Equal(item.Value, rewriteProp.lineMappingDict[item.Key]);
            }
        }
    }
}
