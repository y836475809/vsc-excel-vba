﻿using System;
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
	using ChangeDict = Dictionary<int, List<ChangeVBA>>;
	using ColumnShiftDict = Dictionary<int, List<ColumnShift>>;
	using LineReMapDict = Dictionary<int, int>;
	using DiagoItem = DiagnosticItem;

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

		public static void AssertDiagnoList(List<DiagoItem> pre, List<DiagoItem> act) {
            var comp = (DiagoItem a, DiagoItem b) => {
				if (a.StartLine != b.StartLine) {
					return a.StartLine - b.StartLine;
				}
				return a.StartChara - b.EndChara;
			};
            pre.Sort((a, b) => { return comp(a, b); });
			act.Sort((a, b) => { return comp(a, b); });
			
			foreach ((DiagoItem prItem, DiagoItem actItem) in pre.Zip(act)) {
				Assert.True(prItem.Eq(actItem));
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

        public static void AssertSignatureHelp(List<SignatureHelpItem> pre, List<SignatureHelpItem> act) {
            Assert.Equal(pre.Count, act.Count);
            foreach (var (First, Second) in pre.Zip(act)) {
                Assert.Equal(First.ActiveParameter, Second.ActiveParameter);
                Assert.Equal(First.DisplayText, Second.DisplayText);
                Assert.Equal(First.Description, Second.Description);
                Assert.Equal(First.ReturnType, Second.ReturnType);

                Assert.Equal(First.Args.Count, Second.Args.Count);
                foreach (var (FirstArg, SecondArg) in First.Args.Zip(Second.Args)) { 
                    Assert.Equal(FirstArg.Name, SecondArg.Name);
                    Assert.Equal(FirstArg.AsType, SecondArg.AsType);
                }
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
