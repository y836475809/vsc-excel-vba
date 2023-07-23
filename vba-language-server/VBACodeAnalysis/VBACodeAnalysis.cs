using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Recommendations;
using Microsoft.CodeAnalysis.Text;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.CodeAnalysis.Completion;
using Microsoft.CodeAnalysis.VisualBasic;

namespace VBACodeAnalysis {
    using locDiffDict = Dictionary<int, List<LocationDiff>>;

    public class VBACodeAnalysis {
        private AdhocWorkspace workspace;
        private Project project;
        private Dictionary<string, DocumentId> doc_id_dict;
        private Rewrite rewrite;
        private VBADiagnostic vbaDiagnostic;

        public Dictionary<string, Dictionary<int, int>> lineMappingDict;
        public Dictionary<string, locDiffDict> locationDiffDict;

        public VBACodeAnalysis() {
            var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
            workspace = new AdhocWorkspace(host);

            var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), "MyProject", "MyProject", LanguageNames.VisualBasic).
            WithMetadataReferences(new[] {
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
            });
            project = workspace.AddProject(projectInfo);

            doc_id_dict = new Dictionary<string, DocumentId>();
            lineMappingDict = new Dictionary<string, Dictionary<int, int>>();
            locationDiffDict = new Dictionary<string, locDiffDict>();
        }

        public void setSetting(RewriteSetting rewriteSetting) {
            rewrite = new Rewrite(rewriteSetting);
            vbaDiagnostic = new VBADiagnostic();
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
                locationDiffDict[name] = rewrite.locationDiffDict;
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
            locationDiffDict[name] = rewrite.locationDiffDict;
            lineMappingDict[name] = rewrite.lineMappingDict;
            workspace.TryApplyChanges(
                workspace.CurrentSolution.WithDocumentText(docId, reSourceText));
        }

        private bool IsCompletionItem(ISymbol symbol) {
            var names = new string[] {
                   "Int64", "Int32", "Double", "Byte",
                   "String", "Boolean", "Object", "ValueType"
            };
            var name = symbol.ContainingType?.Name;
            if(name != null) {
                if (names.Contains(name)) {
                    return false;
                }
            }
            return true;
        }

        public int GetCharaDiff(string name, int line, int chara) {
            if (!locationDiffDict.ContainsKey(name)) {
                return 0;
            }
            var diffDict = locationDiffDict[name];
            if (!diffDict.ContainsKey(line)) {
                return 0;
            }
            var diffs = diffDict[line];
            var sumDiff = diffs.Where(x => x.Chara <= chara).Select(x => x.Diff).Sum();
            return sumDiff;
        }

        public async Task<List<CompletionItem>> GetCompletions(string name, string text, int line, int chara) {
			var completions = new List<CompletionItem>();
            if (!doc_id_dict.ContainsKey(name)) {
                return completions;
            }

            var docId = doc_id_dict[name];
            var doc = workspace.CurrentSolution.GetDocument(docId);

            var position = GetPosition(doc, line, chara);
            if(position < 0) {
                return completions;
            }
            var symbols = await Recommender.GetRecommendedSymbolsAtPositionAsync(doc, position);
            foreach (var symbol in symbols) {
                if (!IsCompletionItem(symbol)) {
                    continue;
                }
                var completionItem = new CompletionItem();
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

            var completionService = CompletionService.GetService(doc);
            var results = await completionService.GetCompletionsAsync(doc, position);
            if (results.ItemsList.Any()) {
                var items = results.ItemsList.Where(x => {
                    return x.Tags.Contains("Keyword")
                        && !(completions.Exists(y => y.DisplayText == x.DisplayText));
                }).Select(x => {
                    var compItem = new CompletionItem {
                        DisplayText = x.DisplayText,
                        CompletionText = x.DisplayText,
                        Description = x.Properties.Values.ToString(),
                        Kind = "Keyword"
                    };
                    return compItem;
                });
                if (items.Any()) {
                    completions.AddRange(items);
                    completions.Add(new CompletionItem {
                        DisplayText = "Variant",
                        CompletionText = "Variant",
                        Description = "Variant",
                        Kind = "Keyword"
                    });
                }
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
                var doc = workspace.CurrentSolution.GetDocument(docId);
                var model = await doc.GetSemanticModelAsync();

                var position = GetPosition(doc, line, chara);
                if (position < 0) {
                    return items;
                }

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

        public (int, int, int) GetSignaturePosition(string name, int line, int chara) {
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

            var doc = workspace.CurrentSolution.GetDocument(docId);
            var position = GetPosition(doc, line, chara);
            if (position < 0) {
                return (procLine, procChara, -1);
            }

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

        public async Task<List<SignatureHelpItem>> GetSignatureHelp(string name, int line, int chara) {
            var items = new List<SignatureHelpItem>();

            if (!doc_id_dict.ContainsKey(name)) {
                return items;
            }
            var docId = doc_id_dict[name];
            if (!workspace.CurrentSolution.ContainsDocument(docId)) {
                return items;
            }

            var doc = workspace.CurrentSolution.GetDocument(docId);

            var position = GetPosition(doc, line, chara);
            if (position < 0) {
                return items;
            }

            var model = await doc.GetSemanticModelAsync();
            var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);
            if (symbol == null) {
                return items;
            }

            if (symbol.Kind == SymbolKind.Method) {
                var menbers = symbol.ContainingType.GetMembers(symbol.Name);
                foreach (var menber in menbers) {
                    var item = new SignatureHelpItem();
                    var methodSymbol = menber as IMethodSymbol;
                    foreach (var param in methodSymbol.Parameters) {
                        item.Args.Add(new ArgumentItem(
                            param.Name, ConvKind(param.Type.Name)));
                    }
                    var displayText = string.Join("", methodSymbol.ToDisplayParts().Select(x => {
                        return ConvKind(x.ToString());
                    }));
                    item.DisplayText = displayText;
                    item.Description = methodSymbol.GetDocumentationCommentXml();
                    item.Kind = methodSymbol.Kind.ToString();
                    item.ReturnType = ConvKind(methodSymbol.ReturnType.Name);
                    items.Add(item);
                }
            }
            if (symbol is ILocalSymbol localSymbol) {
                items = GetPropSignatureHelpItems(
                    localSymbol.Type.Name, localSymbol.Type.GetMembers());
            }
            if (symbol is IFieldSymbol filedSymbol) {
                items = GetPropSignatureHelpItems(
                    filedSymbol.Type.Name, filedSymbol.Type.GetMembers());
            }
            if (symbol is IPropertySymbol propSymbol) {
                items = GetPropSignatureHelpItems(
                    propSymbol.Type.Name, propSymbol.Type.GetMembers());
            }
            return items;
        }

        private List<SignatureHelpItem> GetPropSignatureHelpItems(string symbolTypeName, IEnumerable<ISymbol> members) {
            var items = new List<SignatureHelpItem>();

            var menberSymbols = members.Where(x => {
                return (x.Kind == SymbolKind.Property) && (x as IPropertySymbol).IsDefault();
            });
            foreach (var symbol in menberSymbols) {
                var item = new SignatureHelpItem();
                var propSymbol = symbol as IPropertySymbol;
                foreach (var param in propSymbol.Parameters) {
                    item.Args.Add(new ArgumentItem(
                        param.Name, ConvKind(param.Type.Name)));
                }
                var displayText = string.Join("", propSymbol.ToDisplayParts().Select(x => {
                    return ConvKind(x.ToString());
                }));
                var index = displayText.IndexOf("(");
                if (index >= 0) {
                    displayText = $"{symbolTypeName}{displayText[index..]}";
                }
                item.DisplayText = displayText;
                item.Description = propSymbol.GetDocumentationCommentXml();
                item.Kind = propSymbol.Kind.ToString();
                item.ReturnType = ConvKind(propSymbol.GetMethod.ReturnType.Name);
                items.Add(item);
            }
            return items;
        }

        public async Task<List<CompletionItem>> GetHover(string name, int line, int chara) {
            var items = new List<CompletionItem>();
            if (!doc_id_dict.ContainsKey(name)) {
                return items;
            }
            var docId = doc_id_dict[name];
            if (!workspace.CurrentSolution.ContainsDocument(docId)) {
                return items;
            }

            var doc = workspace.CurrentSolution.GetDocument(docId);
            var model = await doc.GetSemanticModelAsync();

            var position = GetPosition(doc, line, chara);
            if (position < 0) {
                return items;
            }

            var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);
            if (symbol == null) {
                return items;
            }
			var item = new CompletionItem {
				DisplayText = symbol.ToDisplayString(),
				CompletionText = symbol.MetadataName,
				Description = symbol.GetDocumentationCommentXml(),
				Kind = symbol.Kind.ToString(),
				ReturnType = ""
			};
			if (symbol.Kind == SymbolKind.Method) {
                var methodSymbol = symbol as IMethodSymbol;
                if(methodSymbol.MethodKind == MethodKind.Constructor) {
                    item.DisplayText = $"Class {methodSymbol.ContainingType.Name}";
                    item.Kind = TypeKind.Class.ToString();
				} else {
                    item.DisplayText = string.Join("", methodSymbol.ToDisplayParts().Select(x => {
                        return ConvKind(x.ToString());
                    }));
                }
                item.ReturnType = ConvKind(methodSymbol.ReturnType.Name);

                var menbersNum = symbol.ContainingType.GetMembers(symbol.Name).Length;
                if(menbersNum > 1) {
                    item.DisplayText = $"{item.DisplayText} (+{menbersNum - 1} overloads)";
                }
            }
            if (symbol is IPropertySymbol propSymbol) {
                item.DisplayText = string.Join("", propSymbol.ToDisplayParts().Select(x => {
                    return ConvKind(x.ToString());
                }));
                item.ReturnType = ConvKind(propSymbol.Type.Name);
            }
            if (symbol is INamedTypeSymbol namedType) {
                if (namedType.TypeKind == TypeKind.Class) {
                    item.DisplayText = $"Class {symbol.MetadataName}";
                    item.Kind = TypeKind.Class.ToString();
				}
            }
            if (symbol is IFieldSymbol fieldSymbol) {
                SetVariableItem(symbol, ref item);
            }
            if (symbol is ILocalSymbol localSymbol) {
                SetVariableItem(symbol, ref item);
            }

            items.Add(item);
            return items;
        }

        private void SetVariableItem(ISymbol symbol, ref CompletionItem item) {
            string symbolName = "";
            string typeName = "";
            string accessibility = "";
            bool isConst = false;
            object constValue = null;
            if (symbol is IFieldSymbol fieldSymbol) {
                symbolName = fieldSymbol.Name;
                typeName = ConvKind(fieldSymbol.Type.Name);
                accessibility = fieldSymbol.DeclaredAccessibility.ToString();
                isConst = fieldSymbol.IsConst;
                constValue = fieldSymbol.ConstantValue;
            }
            if (symbol is ILocalSymbol localSymbol) {
                symbolName = localSymbol.Name;
                typeName = ConvKind(localSymbol.Type.Name);
                accessibility = "Local";
                isConst = localSymbol.IsConst;
                constValue = localSymbol.ConstantValue;
            }
            var dispText = $"{accessibility} {symbolName} As {typeName}";
            if (isConst) {
                if (typeName.ToLower() == "string") {
                    constValue = @$"""{constValue}""";
                }
                dispText = $"{accessibility} Const {symbolName} As {typeName} = {constValue}";
            }

            item.DisplayText = dispText;
            item.CompletionText = typeName;
            item.Description = symbol.GetDocumentationCommentXml();
            item.Kind = symbol.Kind.ToString();
            item.ReturnType = typeName;
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
			return await vbaDiagnostic.GetDiagnostics(doc);
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

            var doc = workspace.CurrentSolution.GetDocument(docId);

            var position = GetPosition(doc, line, chara);
            if (position < 0) {
                return items;
            }

            var model = await doc.GetSemanticModelAsync();
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

        private int GetPosition(Document doc, int line, int chara) {
            var lines = doc.GetTextAsync().Result.Lines;
            if (lines.Count <= line || line < 0) {
                return -1;
            }
            if (lines[line].End <= chara || chara < 0) {
                return -1;
            }
            var position = lines.GetPosition(new LinePosition(line, chara));
            return position;
        }
    }
}
