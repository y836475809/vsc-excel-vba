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
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using LSP = Microsoft.VisualStudio.LanguageServer.Protocol;

namespace VBACodeAnalysis {
    public class VBACodeAnalysis {
        private AdhocWorkspace workspace;
        private Project project;
        private Dictionary<string, DocumentId> doc_id_dict;
		private PreprocVBA _preprocVBA;
		private VBADiagnostic vbaDiagnostic;

        public VBACodeAnalysis() {
            var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
            workspace = new AdhocWorkspace(host);

            var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), "MyProject", "MyProject", LanguageNames.VisualBasic).
            WithMetadataReferences(new[] {
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
            });
            project = workspace.AddProject(projectInfo);

            doc_id_dict = [];
        }

		public void setSetting(RewriteSetting rewriteSetting) {
            _preprocVBA = new PreprocVBA();
			vbaDiagnostic = new VBADiagnostic();
        }

        public string Rewrite(string name, string vbaCode) {
            return _preprocVBA.Rewrite(name, vbaCode);
		}

		public void AddDocument(string name, string text, bool applyChanges= true) {
            if (doc_id_dict.TryGetValue(name, out DocumentId docId)) {
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
				var text = doc.GetTextAsync().Result; 
				var rewriteCode = _preprocVBA.Rewrite(name, text.ToString());
				solution = solution.WithDocumentText(docId, SourceText.From(rewriteCode));
			}
            workspace.TryApplyChanges(solution);
        }

        public void DeleteDocument(string name)
        {
            if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
                return;
            }
            workspace.TryApplyChanges(
               workspace.CurrentSolution.RemoveDocument(docId));
            doc_id_dict.Remove(name);
        }

        public void ChangeDocument(string name, string vbCode) {
            if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
                return;
            }
            if (name.EndsWith(".d.vb")) {
                return;
            }
			workspace.TryApplyChanges(
                workspace.CurrentSolution.WithDocumentText(
                    docId, SourceText.From(vbCode)));
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
            var colShift = _preprocVBA.GetColShift(name, line, chara);
            return colShift;
		}

        public async Task<List<CompletionItem>> GetCompletions(string name, string text, int line, int chara) {
			var completions = new List<CompletionItem>();
            if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
                return completions;
            }
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
            if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
                return items;
            }
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
                        var mapLineIndex = _preprocVBA.GetReMapLineIndex(tree.FilePath, start.Line);
                        if(mapLineIndex >= 0) {
                            items.Add(new DefinitionItem(
                                tree.FilePath,
                                new Location(span.Value.Start, mapLineIndex, 0),
                                new Location(span.Value.End, mapLineIndex, 0),
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

        public (int, int, int) GetSignaturePosition(string name, int line, int chara) {
            var non = (-1, -1, -1);

            if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
                return non;
            }
            if (!workspace.CurrentSolution.ContainsDocument(docId)) {
                return non;
            }

            var doc = workspace.CurrentSolution.GetDocument(docId);
            var position = GetPosition(doc, line, chara);
            if (position < 0) {
                return non;
            }

            var rootNode = doc.GetSyntaxRootAsync().Result;
            var currentToken = rootNode.FindToken(position);
            var currentNode = rootNode.FindNode(currentToken.Span);

            if (currentNode is ArgumentListSyntax argList) {
				return GetProcAndArgPosition(argList, position);
			}

			if (currentNode is SimpleArgumentSyntax simpleArgs) {
                if (simpleArgs.Parent is not ArgumentListSyntax parentArgList) {
					return non;
				}
				return GetProcAndArgPosition(parentArgList, position);
			}

			return non;
        }

        private (int, int, int) GetProcAndArgPosition(ArgumentListSyntax argList, int position) {
			var argPosition = 0;
			var non = (-1, -1, -1);

			if (!argList.CloseParenToken.IsMissing) {
				var closeParent = argList.CloseParenToken;
				if (closeParent.Span.End <= position) {
					return non;
				}
			}

			var commaTokens = argList.DescendantTokens()
                .Where(x => x.IsKind(SyntaxKind.CommaToken));
			foreach (var item in commaTokens) {
				if (item.Span.End > position) {
					break;
				}
				argPosition++;
			}
			var procToken = argList.OpenParenToken.GetPreviousToken();
			var procSp = procToken.GetLocation().GetLineSpan().StartLinePosition;
			return (procSp.Line, procSp.Character, argPosition);
		}

        public async Task<List<SignatureHelpItem>> GetSignatureHelp(string name, int line, int chara) {
            var items = new List<SignatureHelpItem>();

            if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
                return items;
            }
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
            if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
                return items;
            }
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
                if (Util.Eq(typeName, "string")) {
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

        private string ConvKind(string typeName) {
            var convTypeName = typeName;
            if (Util.Eq(typeName, "int64")) {
                convTypeName = "Long";
            }
            if (Util.Eq(typeName, "int32")) {
                convTypeName = "Integer";
            }
            if (Util.Eq(typeName, "datetime")) {
                convTypeName = "Date";
            }
            if (Util.Eq(typeName, "object")) {
                convTypeName = "Variant";
            }
            return convTypeName;
        }

        public async Task<List<DiagnosticItem>> GetDiagnostics(string name) {
			if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
				return [];
			}
			var doc = workspace.CurrentSolution.GetDocument(docId);
			vbaDiagnostic.ignoreDs = _preprocVBA.GetIgnoreDiagnostics(name);
            var prepDiagnosticList = _preprocVBA.GetDiagnostics(name);
			var diagnosticList = await vbaDiagnostic.GetDiagnostics(doc);
            return [.. diagnosticList.Concat(prepDiagnosticList)];
         }

        public async Task<List<ReferenceItem>> GetReferences(string name, int line, int chara) {
            var items = new List<ReferenceItem>();
            if (!doc_id_dict.TryGetValue(name, out DocumentId value)) {
                return items;
            }
            var docId = value;
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

        public LSP.DocumentSymbol[] GetDocumentSymbols(string name, Uri uri) {
			if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
				return [];
			}
			var doc = workspace.CurrentSolution.GetDocument(docId);
			var node = doc.GetSyntaxRootAsync().Result;
            var docSymbols = DocumentSymbolProvider.GetDocumentSymbols(node, uri,  (int line) => {
                if(_preprocVBA.TryGetProperty(name, line, out string prefix, out string propName)) {
                    return (true, prefix, propName);
                }
                return (false, null, null);
			});
            return docSymbols;
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
