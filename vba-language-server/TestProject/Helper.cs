using System;
using System.Collections.Generic;
using System.IO;
//using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using VBACodeAnalysis;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using Xunit;
using System.Linq;

namespace TestProject {
    using locDiffDiict = Dictionary<int, List<LocationDiff>>;

    class Helper {
        public static string getPath(string fileName, [CallerFilePath] string filePath = "") {
            return Path.Combine(Path.GetDirectoryName(filePath), "code", fileName);
        }
        public static int getPosition(string code, string target) {
            return code.LastIndexOf(target) + target.Length;
        }
        public static string getCode(string fileName) {
            var filePath = Helper.getPath(fileName);
            using (var sr = new StreamReader(filePath)) {
                return sr.ReadToEnd();
            }
        }
        public static Document MakeDoc(string code, string name = "testcode") {
            var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
            var workspace = new AdhocWorkspace(host);
            workspace.ClearSolution();
            var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), "MyProject", "MyProject", LanguageNames.VisualBasic).
            WithMetadataReferences(new[] {
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
            });
            var project = workspace.AddProject(projectInfo);
            return workspace.AddDocument(
                project.Id, name, SourceText.From(code));
        }

        public static List<DiagnosticItem> GetDiagnostics(string code, List<string> errorTypes) {
            var vbaDiagnostic = new VBADiagnostic();
            var doc = MakeDoc(code);
            var items = vbaDiagnostic.GetDiagnostics(doc).Result;
            items.Sort((a, b) => {
                if (a.StartLine != b.StartLine) {
                    return a.StartLine - b.StartLine;
                }
                return a.StartChara - b.EndChara;
            });
            return items.FindAll(x => errorTypes.Contains(x.Severity.ToLower()));
        }

        public static void AssertCode(string pre, string act) {
            var preLines = pre.Split("\r\n");
            var actLines = act.Split("\r\n");
            Assert.Equal(preLines.Length, actLines.Length);
            for (int i = 0; i < preLines.Length; i++) {
                Assert.True(preLines[i] == actLines[i], $"Fault {i}");
            }
        }

        public static void AssertLocationDiffDict(locDiffDiict pre, locDiffDiict act) {
            pre.All(x => act.Contains(x));
            foreach (var item in pre) {
                var preList = item.Value;
                var actList = act[item.Key];
                foreach (var (First, Second) in preList.Zip(actList)) {
                    Assert.Equal(First.Line, Second.Line);
                    Assert.Equal(First.Chara, Second.Chara);
                    Assert.Equal(First.Diff, Second.Diff);
                }
            }
        }
    }
}
