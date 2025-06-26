global using ColumnShift = VBARewrite.ColumnShift;
global using ColumnShiftDict = System.Collections.Generic.Dictionary<int, System.Collections.Generic.List<VBARewrite.ColumnShift>>;
global using LineMapDict = System.Collections.Generic.Dictionary<int, int>;
global using VBADiagnostic = VBARewrite.VBADiagnostic;

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using VBACodeAnalysis;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using Xunit;
using System.Linq;
using System.Text;
using VBARewrite;

namespace TestProject {
	class TestVBARewriter : VBARewriter {
		public ColumnShiftDict ColDict(string name) {
			return vbCodeDict[name].ColShiftDict;
		}

		public LineMapDict LineMapDict(string name) {
			return vbCodeDict[name].LineMapDict;
		}
	}

	class Helper {
        public static string getPath(string fileName, [CallerFilePath] string filePath = "") {
            return Path.Combine(Path.GetDirectoryName(filePath), "code", fileName);
        }
        public static int getPosition(string code, string target) {
            return code.LastIndexOf(target) + target.Length;
        }
        public static string getCode(string fileName) {
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
			var filePath = Helper.getPath(fileName);
            return VBALanguageServer.Util.GetCode(filePath);
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

        public static List<VBADiagnostic> GetDiagnostics(string code, List<string> errorTypes) {
            var provider = new VBADiagnosticProvider();
            var doc = MakeDoc(code);
            var items = provider.GetDiagnostics(doc).Result;
            items.Sort((a, b) => {
                if (a.Start.Item1 != b.Start.Item1) {
                    return a.Start.Item1 - b.Start.Item1;
                }
                return a.Start.Item2 - b.End.Item2;
            });
            return items.FindAll(x => errorTypes.Contains(x.Severity.ToLower()));
        }

		public static void AssertDiagnostics(List<VBADiagnostic> pre, List<VBADiagnostic> act) {
            var comp = (VBADiagnostic a, VBADiagnostic b) => {
				if (a.Start.Item1 != b.Start.Item1) {
					return a.Start.Item1 - b.Start.Item1;
				}
				return a.Start.Item2 - b.End.Item2;
			};
            pre.Sort((a, b) => { return comp(a, b); });
			act.Sort((a, b) => { return comp(a, b); });
			
			foreach ((VBADiagnostic prItem, VBADiagnostic actItem) in pre.Zip(act)) {
				Assert.Equal(prItem.ID, actItem.ID);
				Assert.Equal(prItem.Code, actItem.Code);
				Assert.Equal(prItem.Severity, actItem.Severity);
				//Assert.Equal(prItem.Message, actItem.Message);
				Assert.Equal(prItem.Start.Item1, actItem.Start.Item1);
				Assert.Equal(prItem.Start.Item2, actItem.Start.Item2);
				Assert.Equal(prItem.End.Item1, actItem.End.Item1);
				Assert.Equal(prItem.End.Item2, actItem.End.Item2);
			}
		}

		public static void AssertCode(string pre, string act) {
            var preLines = pre.Split("\r\n");
            var actLines = act.Split("\r\n");
            Assert.Equal(preLines.Length, actLines.Length);
            for (int i = 0; i < preLines.Length; i++) {
                Assert.True(preLines[i] == actLines[i], $"Fault {i}");
            }
        }

        public static void AssertSignatureHelp(List<VBASignatureInfo> pre, List<VBASignatureInfo> act) {
            Assert.Equal(pre.Count, act.Count);
            foreach (var (First, Second) in pre.Zip(act)) {
                Assert.Equal(First.Label, Second.Label);
                Assert.Equal(First.Doc, Second.Doc);
                Assert.Equal(First.ParameterInfos.Count, Second.ParameterInfos.Count);
                foreach (var (firstParam, secondParam) in First.ParameterInfos.Zip(Second.ParameterInfos)) { 
                    Assert.Equal(firstParam.Label, secondParam.Label);
                    Assert.Equal(firstParam.Doc, secondParam.Doc);
                }
            }
        }

        public static void AssertCompletionItem(List<VBACompletionItem> pre, List<VBACompletionItem> act) {
			Assert.Equal(pre.Count, act.Count);
			foreach (var (first, second) in pre.Zip(act)) {
                Assert.Equal(first.Label, second.Label);
				Assert.Equal(first.Display, second.Display);
				Assert.Equal(first.Doc, second.Doc);
				Assert.Equal(first.Kind, second.Kind);
			}
        }

		public static void AssertColumnShiftDict(ColumnShiftDict pre, ColumnShiftDict act) {
			Assert.True(pre.Keys.All(x => act.Keys.Contains(x)));
			foreach (var item in pre) {
				var preList = item.Value;
				var actList = act[item.Key];
				foreach (var (First, Second) in preList.Zip(actList)) {
				    Assert.Equal(First.LineIndex, Second.LineIndex);
					Assert.Equal(First.StartCol, Second.StartCol);
					Assert.Equal(First.ShiftCol, Second.ShiftCol);
				}
			}
		}

		public static void AssertLineShift(List<(int, int)> exp, List<(int, int)> act) {
			Assert.Equal(exp.Count, act.Count);
			var items = exp.Zip(act, (line, count) => (line, count));
			foreach (var (v1, v2) in items) {
				Assert.True(v1.Equals(v2));
			}
		}

		public static void AssertDict<T1, T2>(Dictionary<T1, T2> pre, Dictionary<T1, T2> act) {
			Assert.True(pre.Keys.All(x => act.Keys.Contains(x)));
			foreach (var item in pre) {
				var preValue = item.Value;
				var actValue = act[item.Key];
				Assert.Equal(preValue, actValue);
			}
		}
	}
}
