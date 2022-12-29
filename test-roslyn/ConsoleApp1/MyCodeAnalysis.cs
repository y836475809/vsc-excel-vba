using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Recommendations;
using Microsoft.CodeAnalysis.Text;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

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

        public void AddDocument(string name, string text) {
            if (doc_id_dict.ContainsKey(name)){
                ChangeDocument(name, text);
            } else {
                var doc = workspace.AddDocument(project.Id, name, SourceText.From(text));
                doc_id_dict.Add(name, doc.Id);
            }
        }

        public void DeleteDocument(string name)
        {
            var docId = doc_id_dict[name];
            workspace.TryApplyChanges(
               workspace.CurrentSolution.RemoveDocument(docId));
            doc_id_dict.Remove(name);
        }

        public void ChangeDocument(string name, string text) {
            var docId = doc_id_dict[name];
            workspace.TryApplyChanges(
                workspace.CurrentSolution.WithDocumentText(docId, SourceText.From(text)));
        }


        public async Task<IEnumerable<CompletionItem>> GetCompletions(string name, string text, int position) {
            ChangeDocument(name, text);
            var completions = new List<CompletionItem>();
            var docId = doc_id_dict[name];
            var doc = workspace.CurrentSolution.GetDocument(docId);
            var symbols = await Recommender.GetRecommendedSymbolsAtPositionAsync(doc, position);
            foreach (var symbol in symbols) {
                var completionItem = new CompletionItem();
                if(symbol.ContainingType?.Name == "Object") {
                    continue;
                }
                completionItem.DisplayText = symbol.ToDisplayString();
                completionItem.CompletionText = symbol.MetadataName;
                completionItem.Description = symbol.GetDocumentationCommentXml();
                completionItem.Kind = symbol.Kind.ToString();
                completionItem.ReturnType = "";
                if (symbol.Kind == SymbolKind.Method) {
                    var methodSymbol = symbol as IMethodSymbol;
                    completionItem.ReturnType = methodSymbol.ReturnType.ToDisplayString();
                }
                completions.Add(completionItem);
            }
            return completions;
        }

        public int p1() {
            return 10;
        }
    }
}
