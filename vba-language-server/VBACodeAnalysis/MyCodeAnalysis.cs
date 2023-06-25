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
using System.Linq;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using Microsoft.CodeAnalysis.Completion;
using Microsoft.CodeAnalysis.VisualBasic;

namespace VBACodeAnalysis {
    using lineCharaOffDict = Dictionary<int, (int, int)>;

    public class MyCodeAnalysis {
        private AdhocWorkspace workspace;
        private Project project;
        private Dictionary<string, DocumentId> doc_id_dict;
        private RewriteSetting rewriteSetting;
        private Rewrite rewrite;
        private MyDiagnostic myDiagnostic;

        public Dictionary<string, lineCharaOffDict> charaOffsetDict;
        public Dictionary<string, Dictionary<int, int>> lineMappingDict;

        public MyCodeAnalysis() {
            var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
            workspace = new AdhocWorkspace(host);

            var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), "MyProject", "MyProject", LanguageNames.VisualBasic).
            WithMetadataReferences(new[] {
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
            });
            project = workspace.AddProject(projectInfo);

            doc_id_dict = new Dictionary<string, DocumentId>();
            charaOffsetDict = new Dictionary<string, lineCharaOffDict>();
            lineMappingDict = new Dictionary<string, Dictionary<int, int>>();
        }

        public void setSetting(RewriteSetting rewriteSetting) {
            this.rewriteSetting = rewriteSetting;
            rewrite = new Rewrite(rewriteSetting);
            myDiagnostic = new MyDiagnostic(rewriteSetting);
        }

        public void AddDocument(string name, string text, bool applyChanges= true) {
            if (doc_id_dict.ContainsKey(name)) {
                var docId = doc_id_dict[name];
                workspace.TryApplyChanges(
                    workspace.CurrentSolution.WithDocumentText(
                        docId, SourceText.From(text)));
            } else {
                var doc = workspace.AddDocument(project.Id, name, SourceText.From(text));
                workspace.TryApplyChanges(
                    workspace.CurrentSolution.WithDocumentFilePath(doc.Id, name));
                doc_id_dict.Add(name, doc.Id);
            }
			if (applyChanges) {
                ChangeDocument(name, text);
            }
        }
        public void ApplyChanges(List<string> names) {
            Solution solution = workspace.CurrentSolution;
            foreach (var name in names) {
                if (!doc_id_dict.ContainsKey(name)) {
                    continue;
                }
                if (name.EndsWith(".d.vb")) {
                    continue;
                }
                var docId = doc_id_dict[name];
                var doc = solution.GetDocument(docId);
                var reSourceText = rewrite.RewriteStatement(doc);
                charaOffsetDict[name] = rewrite.charaOffsetDict;
                lineMappingDict[name] = rewrite.lineMappingDict;

                solution = solution.WithDocumentText(docId, reSourceText);
            }
            workspace.TryApplyChanges(solution);
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
            if (name.EndsWith(".d.vb")) {
                return;
            }
            var docId = doc_id_dict[name];
            var doc = workspace.CurrentSolution.GetDocument(docId);
            doc = doc.WithText(SourceText.From(text));
            var reSourceText = rewrite.RewriteStatement(doc);
            charaOffsetDict[name] = rewrite.charaOffsetDict;
            lineMappingDict[name] = rewrite.lineMappingDict;
            workspace.TryApplyChanges(
                workspace.CurrentSolution.WithDocumentText(docId, reSourceText));
        }

        private bool IsCompletionItem(ISymbol symbol) {
            var names = new string[] {
                   "Int64", "Int32", "Double", "Byte",
                   "String", "Boolean", "Object"
            };
            var name = symbol.ContainingType?.Name;
            if(name != null) {
                //Console.WriteLine($"name={name}");
                if (names.Contains(name)) {
                    return false;
                }
            }
            return true;
        }

        private int getoffset(string name, int line, int chara) {
            if (!charaOffsetDict.ContainsKey(name)) {
                return 0;
            }
            var offdict = charaOffsetDict[name];
            if (offdict.ContainsKey(line)) {
                var (s, offset) = offdict[line];
                if(chara >= s) {
                    return offset;
                }
            }
            return 0;
        }

        public async Task<List<CompletionItem>> GetCompletions(string name, string text, int line, int chara) {
			ChangeDocument(name, text);
			var completions = new List<CompletionItem>();
            if (!doc_id_dict.ContainsKey(name)) {
                return completions;
            }
            var adjChara = getoffset(name, line, chara) + chara;
            var docId = doc_id_dict[name];
            var doc = workspace.CurrentSolution.GetDocument(docId);
            var position = doc.GetTextAsync().Result.Lines.GetPosition(new LinePosition(line, adjChara));

            var completionService = CompletionService.GetService(doc);
            var results = await completionService.GetCompletionsAsync(doc, position);
            completions.AddRange(results.ItemsList.Where(x => {
                return x.Tags.Contains("Keyword");
            }).Select(x => {
                var compItem = new CompletionItem();
                compItem.DisplayText = x.DisplayText;
                compItem.CompletionText = x.DisplayText;
                compItem.Description = x.Properties.Values.ToString();
                compItem.Kind = "Keyword";
                return compItem;
            }));

            var symbols = await Recommender.GetRecommendedSymbolsAtPositionAsync(doc, position);
            foreach (var symbol in symbols) {
                var completionItem = new CompletionItem();
                if (!IsCompletionItem(symbol)) {
                    break;
                }
                completionItem.DisplayText = symbol.MetadataName;
                completionItem.CompletionText = symbol.ToDisplayString();
                completionItem.Description = symbol.GetDocumentationCommentXml();
                completionItem.Kind = symbol.Kind.ToString();
                if (symbol is INamedTypeSymbol namedType) {
                    if (namedType.TypeKind == TypeKind.Class) {
                        completionItem.Kind = TypeKind.Class.ToString();
                    }
                }   
                completionItem.ReturnType = "";
                if (symbol.Kind == SymbolKind.Method) {
                    var methodSymbol = symbol as IMethodSymbol;
                    completionItem.ReturnType = methodSymbol.ReturnType.ToDisplayString();
                }
                completions.Add(completionItem);
            }
            return completions;
        }

        public async Task<List<DefinitionItem>> GetDefinitions(string name, string text, int line, int chara) {
            var items = new List<DefinitionItem>();
            if (!doc_id_dict.ContainsKey(name)) {
                return items;
            }
            var docId = doc_id_dict[name];
            if (workspace.CurrentSolution.ContainsDocument(docId)) {
                var adjChara = getoffset(name, line, chara) + chara;
                var doc = workspace.CurrentSolution.GetDocument(docId);
                var model = await doc.GetSemanticModelAsync();
                var lines = doc.GetTextAsync().Result.Lines;
                if(lines.Count <= line || line < 0) {
                    return items;
                }
                if (lines[line].End <= adjChara || adjChara < 0) {
                    return items;
                }

                var position = lines.GetPosition(new LinePosition(line, adjChara));
                var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);

                if (symbol == null) {
                    return items;
                }
                bool isClass = false;
                if (symbol is INamedTypeSymbol namedType) {
                    if (namedType.TypeKind == TypeKind.Class) {
                        isClass = true;
                    }
                }
                if(symbol is IMethodSymbol methodTypel) {
                    if (methodTypel.MethodKind == MethodKind.Constructor) {
                        isClass = true;
                    }
                }
                foreach (var loc in symbol.Locations) {
                    var span = loc?.SourceSpan;
                    var tree = loc?.SourceTree;
                    if (span != null && tree != null) {
                        var start = tree.GetLineSpan(span.Value).StartLinePosition;
                        var end = tree.GetLineSpan(span.Value).EndLinePosition;
                        if(HasLineMapping(tree.FilePath, start.Line)) {
                            var linemap = lineMappingDict[tree.FilePath];
                            var mapedline = linemap[start.Line];
                            items.Add(new DefinitionItem(
                                tree.FilePath,
                                new Location(span.Value.Start, mapedline, 0),
                                new Location(span.Value.End, mapedline, 0),
                                isClass));
						} else {
                            items.Add(new DefinitionItem(
                                tree.FilePath,
                                new Location(span.Value.Start, start.Line, start.Character),
                                new Location(span.Value.End, end.Line, end.Character),
                                isClass));
                        }
                    }
                }
            }
            return items;
        }

        private bool HasLineMapping(string filePath, int startLine) {
            if (!lineMappingDict.ContainsKey(filePath)) {
                return false;
            }
            var linemap = lineMappingDict[filePath];
            if (!linemap.ContainsKey(startLine)) {
                return false;
            }
            return true;
        }

        public (int, int, int) GetSignaturePosition(string name, string text, int line, int chara) {
            var procLine = -1;
            var procChara = -1;
            var argPosition = 0;

            if (!doc_id_dict.ContainsKey(name)) {
                return (procLine, procChara, - 1);
            }
            var docId = doc_id_dict[name];
            if (!workspace.CurrentSolution.ContainsDocument(docId)) {
                return (procLine, procChara, -1);
            }

			ChangeDocument(name, text);

            var doc = workspace.CurrentSolution.GetDocument(docId);
            var position = doc.GetTextAsync().Result.Lines.GetPosition(new LinePosition(line, chara));
            var rootNode = doc.GetSyntaxRootAsync().Result;
            var currentToken = rootNode.FindToken(position);
            var currentNode = rootNode.FindNode(currentToken.Span);

            var preToken = currentToken.GetPreviousToken();
            if (preToken.Text == "(") {
                var chNodes = currentNode.Parent.ChildNodes();
				if (chNodes.Any()) {
					var args = chNodes.Where(x => x.IsKind(SyntaxKind.ArgumentList));
					if (args.Any()) {
                        var procToken = args.First().ChildTokens().First().GetPreviousToken();
                        var lp = procToken.GetLocation().GetLineSpan();
                        procLine = lp.StartLinePosition.Line;
                        procChara = lp.StartLinePosition.Character;
                        return (procLine, procChara, argPosition);
                    }
				}
            } 

            if (currentNode.IsKind(SyntaxKind.ArgumentList)) {
                var procToken = currentNode.ChildTokens().First().GetPreviousToken().GetLocation().GetLineSpan();
                procLine = procToken.StartLinePosition.Line;
                procChara = procToken.StartLinePosition.Character;
                var commnaTokens = currentNode.ChildTokens().Where(x => x.IsKind(SyntaxKind.CommaToken));
                foreach (var item in commnaTokens) {
                    if (item.Span.End <= position) {
                        argPosition++;
                    }
                }
                return (procLine, procChara, argPosition);
            } 

            var node = rootNode.FindNode(currentToken.Span).Parent;
            const int seacrhMax = 5;
            for (int i = 0; i < seacrhMax; i++) {
                if (node.IsKind(SyntaxKind.ArgumentList)) {
                    var procToken = node.ChildTokens().First().GetPreviousToken();
                    var lp = procToken.GetLocation().GetLineSpan();
                    procLine = lp.StartLinePosition.Line;
                    procChara = lp.StartLinePosition.Character;
                    var commnaTokens = node.ChildTokens().Where(x => x.IsKind(SyntaxKind.CommaToken));
                    foreach (var item in commnaTokens) {
                        if (item.Span.End <= position) {
                            argPosition++;
                        }
                    }
                    break;
                }
                if (node.Parent == null) {
                    break;
                }
                node = node.Parent;
            }

            return (procLine, procChara, argPosition);
        }

        public async Task<SignatureHelpItem> GetSignatureHelp(string name, int position) {
            if (!doc_id_dict.ContainsKey(name)) {
                return null;
            }
            var docId = doc_id_dict[name];
            if (!workspace.CurrentSolution.ContainsDocument(docId)) {
                return null;
            }
            var doc = workspace.CurrentSolution.GetDocument(docId);
            var model = await doc.GetSemanticModelAsync();
            var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);
            if (symbol == null) {
                return null;
            }

            var item = new SignatureHelpItem();
            if (symbol is IMethodSymbol methodSymb) {
				foreach (var param in methodSymb.Parameters) {
                    item.Args.Add(new ArgumentItem(
                        param.Name, param.Type.ToDisplayString()));
                } 
            }
            
            if (symbol.ContainingType?.Name == rewriteSetting.NameSpace
                    && rewriteSetting.getRestoreDict().ContainsKey(symbol.Name)) {
                var rewName = $"{symbol.Name}(";
                var resName = $"{rewriteSetting.getRestoreDict()[symbol.Name]}(";
                item.DisplayText = symbol.ToDisplayString().Replace(
                    rewName, resName);
            } else {
                item.DisplayText = symbol.ToDisplayString();
            }

            item.Description = symbol.GetDocumentationCommentXml();
            item.Kind = symbol.Kind.ToString();
            item.ReturnType = "";
            if (symbol.Kind == SymbolKind.Method) {
                var methodSymbol = symbol as IMethodSymbol;
                item.ReturnType = methodSymbol.ReturnType.ToDisplayString();
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
                item.DisplayText = kind;
                item.Kind = kind;
            }
            return item;
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
                //var pp =  SymbolFinder.FindDeclarationsAsync(doc.Project, "Range", false).Result.ToList();
                var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);
                if (symbol != null) {
                    //if (symbol.ContainingType?.Name == "Object") {
                    //    continue;
                    //}
                    
                    //if (symbol.ContainingType?.Name == "f" && symbol.Name == "Ran") {
                    if (symbol.ContainingType?.Name == rewriteSetting.NameSpace
                            && rewriteSetting.getRestoreDict().ContainsKey(symbol.Name)) {
                        var rewName = $"{symbol.Name}(";
                        var resName = $"{rewriteSetting.getRestoreDict()[symbol.Name]}(";
                        //completionItem.DisplayText = symbol.ToDisplayString().Replace("Ran(", "Range(");
                        completionItem.DisplayText = symbol.ToDisplayString().Replace(
                            rewName, resName);
                    } else {
                        completionItem.DisplayText = symbol.ToDisplayString();
                    }
                    //completionItem.DisplayText = symbol.ToDisplayString();
                    completionItem.CompletionText = symbol.MetadataName;
                    completionItem.Description = symbol.GetDocumentationCommentXml();
                    completionItem.Kind = symbol.Kind.ToString();
                    completionItem.ReturnType = "";
                    if (symbol.Kind == SymbolKind.Method) {
                        var methodSymbol = symbol as IMethodSymbol;
                        completionItem.ReturnType = methodSymbol.ReturnType.ToDisplayString();
                    }
                    if (symbol is IFieldSymbol fieldSymbol) {
                        var typeName = ConvKind(fieldSymbol.Type.Name);
                        var accessibility = fieldSymbol.DeclaredAccessibility.ToString();
                        var dispText = $"{accessibility} {fieldSymbol.Name} As {typeName}";
                        if (fieldSymbol.IsConst) {
                            var constValue = fieldSymbol.ConstantValue;
                            if (typeName.ToLower() == "string") {
                                constValue = @$"""{constValue}""";
                            }
                            dispText = $"{accessibility} Const {fieldSymbol.Name} As {typeName} = {constValue}";
						}
                        completionItem.DisplayText = dispText;
                        completionItem.CompletionText = typeName;
                        completionItem.Kind = typeName;
                    }
                    if (symbol is ILocalSymbol localSymbol) {
                        var typeName = ConvKind(localSymbol.Type.Name);
                        var accessibility = "Local";
                        var dispText = $"{accessibility} {localSymbol.Name} As {typeName}";
                        if (localSymbol.IsConst) {
                            var constValue = localSymbol.ConstantValue;
                            if (typeName.ToLower() == "string") {
                                constValue = @$"""{constValue}""";
                            }
                            dispText = $"{accessibility} Const {localSymbol.Name} As {typeName} = {constValue}";
                        }
                        completionItem.DisplayText = dispText;
                        completionItem.CompletionText = typeName;
                        completionItem.Kind = typeName;
                    }
                }
            }
            return completionItem;
        }

        private string ConvKind(string TypeName) {
            var lower = TypeName.ToLower();
            var convTypeName = TypeName;
            if (lower == "int64") {
                convTypeName = "Long";
            }
            if (lower == "int32") {
                convTypeName = "Integer";
            }
            if (lower == "datetime") {
                convTypeName = "Date";
            }
            if (lower == "object") {
                convTypeName = "Variant";
            }
            return convTypeName;
        }

        public async Task<List<DiagnosticItem>> GetDiagnostics(string name) {
			var docId = doc_id_dict[name];
			var doc = workspace.CurrentSolution.GetDocument(docId);
			return await myDiagnostic.GetDiagnostics(doc);
         }

        public async Task<List<ReferenceItem>> GetReferences(string name, int line, int chara) {
            var items = new List<ReferenceItem>();
            if (!doc_id_dict.ContainsKey(name)) {
                return items;
            }
            var docId = doc_id_dict[name];
            if (!workspace.CurrentSolution.ContainsDocument(docId)) {
                return items;
            }
            var adjChara = getoffset(name, line, chara) + chara;
            var doc = workspace.CurrentSolution.GetDocument(docId);
            var model = await doc.GetSemanticModelAsync();
            var position = doc.GetTextAsync().Result.Lines.GetPosition(new LinePosition(line, adjChara));
            var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);
            if (symbol == null) {
                return items;
            }
			if (symbol.IsDefinition) {
                foreach (var loc in symbol.Locations) {
                    var filePath = loc.SourceTree.FilePath;
                    var start = loc.GetLineSpan().StartLinePosition;
                    var end = loc.GetLineSpan().EndLinePosition;
                    var startLoc = new Location(0, start.Line, start.Character);
                    var endLoc = new Location(0, end.Line, end.Character);
                    items.Add(new ReferenceItem(filePath, startLoc, endLoc));
                }
            }

            var refItems = SymbolFinder.FindReferencesAsync(symbol, workspace.CurrentSolution).Result;
            foreach (var refItem in refItems) {
                foreach (var loc in refItem.Locations) {
                    var filePath = loc.Document.FilePath;
                    var start = loc.Location.GetLineSpan().StartLinePosition;
                    var end = loc.Location.GetLineSpan().EndLinePosition;
                    var startLoc = new Location(0, start.Line, start.Character);
                    var endLoc = new Location(0, end.Line, end.Character);
                    items.Add(new ReferenceItem(filePath, startLoc, endLoc));
                }
            }
            return items;
        }
    }
}
