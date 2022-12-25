using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1 {
    public class MyCodeAnalysis {
        private AdhocWorkspace workspace;
        private Project project;
        private Dictionary<string, DocumentId> doc_id_dict;

        public MyCodeAnalysis() {
            var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
            workspace = new AdhocWorkspace(host);

            var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), "MyProject", "MyProject", LanguageNames.VisualBasic).
            WithMetadataReferences(new[] {
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
            });
            project = workspace.AddProject(projectInfo);

            doc_id_dict = new Dictionary<string, DocumentId>();
        }

        public void AddDocumentstring(string name, string text) {
            var doc = workspace.AddDocument(project.Id, name, SourceText.From(text));
            doc_id_dict.Add(name, doc.Id);
        }

        public void ChangeDocumentstring(string name, string text) {
            var docId = doc_id_dict[name];
            workspace.TryApplyChanges(
                workspace.CurrentSolution.WithDocumentText(docId, SourceText.From(text)));
        }

        public void Completion() {

        }

        public int p1() {
            return 10;
        }
    }
}
