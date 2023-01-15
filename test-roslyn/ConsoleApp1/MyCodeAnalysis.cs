using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Recommendations;
using Microsoft.CodeAnalysis.Text;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

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
            if (!doc_id_dict.ContainsKey(name)) {
                return;
            }
            var docId = doc_id_dict[name];
            workspace.TryApplyChanges(
               workspace.CurrentSolution.RemoveDocument(docId));
            doc_id_dict.Remove(name);
        }

        public void ChangeDocument(string name, string text) {
            if (!doc_id_dict.ContainsKey(name)) {
                return;
            }
            var docId = doc_id_dict[name];
            workspace.TryApplyChanges(
                workspace.CurrentSolution.WithDocumentText(docId, SourceText.From(text)));
        }


        public async Task<List<CompletionItem>> GetCompletions(string name, string text, int position) {
            ChangeDocument(name, text);
            var completions = new List<CompletionItem>();
            if (!doc_id_dict.ContainsKey(name)) {
                return completions;
            }
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

        public async Task<List<DefinitionItem>> GetDefinitions(string name, string text, int position) {
            var items = new List<DefinitionItem>();
            if (!doc_id_dict.ContainsKey(name)) {
                return items;
            }
            var docId = doc_id_dict[name];
            if (workspace.CurrentSolution.ContainsDocument(docId)) {
                var doc = workspace.CurrentSolution.GetDocument(docId);
                var model = await doc.GetSemanticModelAsync();
                var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);

                if (symbol == null) {
                    return items;
                }
                foreach (var loc in symbol.Locations) {
                    var span = loc?.SourceSpan;
                    var tree = loc?.SourceTree;
                    if (span != null && tree != null) {
                        var start = tree.GetLineSpan(span.Value).StartLinePosition;
                        var end = tree.GetLineSpan(span.Value).EndLinePosition;
                        items.Add(new DefinitionItem(
                            tree.FilePath,
                            new Location(span.Value.Start,  start.Line, start.Character), 
                            new Location(span.Value.End, end.Line, end.Character)));
                    }
                }
            }
            return items;
        }

        public async Task<CompletionItem> GetHover(string name, string text, int position) {
            //var items = new List<CompletionItem>();
            var completionItem = new CompletionItem();
            if (!doc_id_dict.ContainsKey(name)) {
                return completionItem;
            }
            var docId = doc_id_dict[name];
            if (workspace.CurrentSolution.ContainsDocument(docId)) {
                var doc = workspace.CurrentSolution.GetDocument(docId);
                var model = await doc.GetSemanticModelAsync();
                var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);
                if (symbol != null) {
                    //if (symbol.ContainingType?.Name == "Object") {
                    //    continue;
                    //}
                    completionItem.DisplayText = symbol.ToDisplayString();
                    completionItem.CompletionText = symbol.MetadataName;
                    completionItem.Description = symbol.GetDocumentationCommentXml();
                    completionItem.Kind = symbol.Kind.ToString();
                    completionItem.ReturnType = "";
                    if (symbol.Kind == SymbolKind.Method) {
                        var methodSymbol = symbol as IMethodSymbol;
                        completionItem.ReturnType = methodSymbol.ReturnType.ToDisplayString();
                    }
                    if (symbol.Kind == SymbolKind.Local) {
                        var localSymbol = symbol as ILocalSymbol;
                        var kind = localSymbol.Type.Name;
                        var kindLower = kind.ToLower();
                        if (kindLower == "int64") {
                            kind = "Long";
                        }
                        if (kindLower == "int32") {
                            kind = "Integer";
                        }
                        if (kindLower == "datetime") {
                            kind = "Date";
                        }
                        completionItem.DisplayText = kind;
                        completionItem.CompletionText = kind;
                        completionItem.Kind = kind;
                    }
                }
                //completions.Add(completionItem);
            }
            return completionItem;
        }

        public DescriptionItem ParseDescriptionXML(string xml) {
            var doc = XElement.Parse(xml);
            var summary_text = doc.Element("summary")?.Value??"";
            var ps = new List<DescriptionParam> ();
            var param_elms = doc.Elements("param");
            if (param_elms != null) {
                foreach (var param_elm in param_elms) {
                    var name = param_elm.Attribute("name").Value;
                    ps.Add(new DescriptionParam(name, param_elm.Value));
                }
            }
            var returns_text = doc.Element("returns")?.Value ?? "";
            return new DescriptionItem(summary_text, ps, returns_text);
        }
    }
}
