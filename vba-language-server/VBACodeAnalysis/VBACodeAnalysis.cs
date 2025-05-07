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
using Antlr4.Runtime.Misc;
using System.Reflection.Emit;
using Microsoft.CodeAnalysis.Elfie.Model;
using Microsoft.VisualBasic;

namespace VBACodeAnalysis {
    public class VBACodeAnalysis {
        private AdhocWorkspace workspace;
        private Project project;
        private Dictionary<string, DocumentId> doc_id_dict;
		private PreprocVBA _preprocVBA;
		private VBADiagnosticProvider vbaDiagnosticProvider;

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
			vbaDiagnosticProvider = new VBADiagnosticProvider();
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

        public async Task<List<VBACompletionItem>> GetCompletions(string name, int line, int chara) {
			if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
                return [];
            }
            var doc = workspace.CurrentSolution.GetDocument(docId);
			var adjChara = GetCharaDiff(name, line, chara) + chara;
			var position = GetPosition(doc, line, adjChara);
            if(position < 0) {
                return [];
            }
            var completionItems = new List<VBACompletionItem>();
			var symbols = await Recommender.GetRecommendedSymbolsAtPositionAsync(doc, position);
			var items = symbols.Where(x => IsCompletionItem(x)).Select(y => {
				var label = y.MetadataName;
                var display = y.ToDisplayString();
                if(label.ToLower() == "structure"){
                    label = "Type";
                    display = "Type";
                }
				
				var docment = y.GetDocumentationCommentXml();
				var kind = y.Kind.ToString();
				if (y is INamedTypeSymbol namedType) {
					if (namedType.TypeKind == Microsoft.CodeAnalysis.TypeKind.Class) {
						kind = Microsoft.CodeAnalysis.TypeKind.Class.ToString();
					}
				}
                return new VBACompletionItem {
                    Label = label,
                    Display = display,
                    Doc = docment,
                    Kind = kind
                };
				//return label;
			}).ToList();
            completionItems.AddRange(items);

			var completionService = CompletionService.GetService(doc);
            var completions = await completionService.GetCompletionsAsync(doc, position);
            if (completions.ItemsList.Any()) {
                var results = completions.ItemsList.Where(x => {
                    return x.Tags.Contains("Keyword")
                        && !(items.Exists(y => y.Display == x.DisplayText));
                }).Select(x => {
					var label = x.DisplayText;
					var display = x.DisplayText;
                    if(label.ToLower() == "structure"){
                        label = "Type";
                        display = "Type";
                    }
					var docment = x.Properties.Values.ToString();
					var kind = "Keyword";
					return new VBACompletionItem {
						Label = label,
						Display = display,
						Doc = docment,
						Kind = kind
					};
				});
				completionItems.AddRange(results);

				if (results.Any()) {
                    completionItems.Add(new() {
                        Label = "Variant",
                        Display = "Variant",
                        Doc = "Variant",
                        Kind = "Keyword"
                    });
				}
            }
            return completionItems;
		}

        public async Task<List<VBALocation>> GetDefinitions(string name, int line, int chara) {
			if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
                return [];
            }
            if (!workspace.CurrentSolution.ContainsDocument(docId)) {
				return [];
			}

			var doc = workspace.CurrentSolution.GetDocument(docId);
            var model = await doc.GetSemanticModelAsync();
			var adjChara = GetCharaDiff(name, line, chara) + chara;
			var position = GetPosition(doc, line, adjChara);
            if (position < 0) {
                return [];
            }
            var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);
            if (symbol == null) {
                return [];
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

            var vbaLocations = new List<VBALocation>();
			foreach (var loc in symbol.Locations) {
                var span = loc?.SourceSpan;
                var tree = loc?.SourceTree;
                if (span == null || tree == null) {
                    continue;
                }
                if (isClass) {
                    vbaLocations.Add(new() {
						Uri = new Uri(tree.FilePath),
						Start = (0, 0),
						End = (0, 0)
					});
                    continue;
				}
				var startLinePos = tree.GetLineSpan(span.Value).StartLinePosition;
                var endLinePos = tree.GetLineSpan(span.Value).EndLinePosition;
                var mapLineIndex = _preprocVBA.GetReMapLineIndex(tree.FilePath, startLinePos.Line);
                if (mapLineIndex >= 0) {
					vbaLocations.Add(new() {
						Uri = new Uri(tree.FilePath),
						Start = (mapLineIndex, 0),
						End = (mapLineIndex, 0)
					});
				} else {
					var adjStart = AdjustPosition(tree.FilePath, startLinePos.Line, startLinePos.Character);
					var adjEnd = AdjustPosition(tree.FilePath, endLinePos.Line, endLinePos.Character);
					vbaLocations.Add(new() {
						Uri = new Uri(tree.FilePath),
						Start = adjStart,
						End = adjEnd
					});
				}
			}
            return vbaLocations;
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

        public async Task<(int, List<VBASignatureInfo>)> GetSignatureHelp(
			string name, int line, int chara) {
			var adjChara = GetCharaDiff(name, line, chara) + chara;
			var (sigLine, sigChara, argPosition) = GetSignaturePosition(name, line, adjChara);
            if (sigLine < 0) {
                return (-1, []);
            }
            if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
                return (-1, []);
			}
            if (!workspace.CurrentSolution.ContainsDocument(docId)) {
                return (-1, []);
			}

            var doc = workspace.CurrentSolution.GetDocument(docId);
            var position = GetPosition(doc, sigLine, sigChara);
            if (position < 0) {
                return (-1, []);
			}
            var model = await doc.GetSemanticModelAsync();
            var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);
            if (symbol == null) {
                return (-1, []);
			}

            var vbaSignatureInfos = new List<VBASignatureInfo>();
			if (symbol.Kind == SymbolKind.Method) {
                var menbers = symbol.ContainingType.GetMembers(symbol.Name);
                foreach (var menber in menbers) {
					var parameters = new List<VBAParameterInfo>();
					var methodSymbol = menber as IMethodSymbol;
                    foreach (var param in methodSymbol.Parameters) {
                        parameters.Add(new() {
                            Label = param.Name, 
                            Doc = ConvKind(param.Type.Name)
                        });
                    }
                    var displayText = string.Join("", methodSymbol.ToDisplayParts().Select(x => {
                        return ConvKind(x.ToString());
                    }));
                    vbaSignatureInfos.Add(new() {
                        Label = displayText,
                        Doc = methodSymbol.GetDocumentationCommentXml(),
                        ParameterInfos = parameters
					});
				}
            }
            if (symbol is ILocalSymbol localSymbol) {
				vbaSignatureInfos.AddRange(GetPropSignatureHelpItems(
					localSymbol.Type.Name, localSymbol.Type.GetMembers()));
            }
            if (symbol is IFieldSymbol filedSymbol) {
				vbaSignatureInfos.AddRange(GetPropSignatureHelpItems(
                    filedSymbol.Type.Name, filedSymbol.Type.GetMembers()));
            }
            if (symbol is IPropertySymbol propSymbol) {
                vbaSignatureInfos.AddRange(GetPropSignatureHelpItems(
                    propSymbol.Type.Name, propSymbol.Type.GetMembers()));
            }
            return (argPosition, vbaSignatureInfos);
		}

        private List<VBASignatureInfo> GetPropSignatureHelpItems(
			string symbolTypeName, IEnumerable<ISymbol> members) {
			var vbaSignatureInfos = new List<VBASignatureInfo>();
			var menberSymbols = members.Where(x => {
                return (x.Kind == SymbolKind.Property) && (x as IPropertySymbol).IsDefault();
            });
            foreach (var symbol in menberSymbols) {
				var parameters = new List<VBAParameterInfo>();
				var propSymbol = symbol as IPropertySymbol;
                foreach (var param in propSymbol.Parameters) {
					parameters.Add(new() {
						Label = param.Name,
						Doc = ConvKind(param.Type.Name)
					});
				}
                var displayText = string.Join("", propSymbol.ToDisplayParts().Select(x => {
                    return ConvKind(x.ToString());
                }));
                var index = displayText.IndexOf("(");
                if (index >= 0) {
                    displayText = $"{symbolTypeName}{displayText[index..]}";
                }
				vbaSignatureInfos.Add(new() {
					Label = displayText,
					Doc = propSymbol.GetDocumentationCommentXml(),
					ParameterInfos = parameters
				});
			}
            return vbaSignatureInfos;
		}

        public async Task<VBAHover> GetHover(string name, int line, int chara) {
            if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
                return null;
            }
            if (!workspace.CurrentSolution.ContainsDocument(docId)) {
                return null;
            }

            var doc = workspace.CurrentSolution.GetDocument(docId);
            var model = await doc.GetSemanticModelAsync();
			var adjChara = GetCharaDiff(name, line, chara) + chara;
			var position = GetPosition(doc, line, adjChara);
            if (position < 0) {
                return null;
            }
            var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);
            if (symbol == null) {
                return null;
            }

			string label = "";
			string description = "";
			string returnType = "";
			string kind = "";
			if (symbol.Kind == SymbolKind.Method) {
                var methodSymbol = symbol as IMethodSymbol;
                if(methodSymbol.MethodKind == MethodKind.Constructor) {
					label = $"Class {methodSymbol.ContainingType.Name}";
					kind = TypeKind.Class.ToString();
				} else {
					label = string.Join("", methodSymbol.ToDisplayParts().Select(x => {
						return ConvKind(x.ToString());
					}));
				}
				returnType = ConvKind(methodSymbol.ReturnType.Name);

				var menbersNum = symbol.ContainingType.GetMembers(symbol.Name).Length;
                if(menbersNum > 1) {
					label = $"{label} (+{menbersNum - 1} overloads)";
				}

                description = symbol.GetDocumentationCommentXml();
			}
            if (symbol is IPropertySymbol propSymbol) {
				label = string.Join("", propSymbol.ToDisplayParts().Select(x => {
					return ConvKind(x.ToString());
				}));
				returnType = ConvKind(propSymbol.Type.Name);
			}
            if (symbol is INamedTypeSymbol namedType) {
                if (namedType.TypeKind == TypeKind.Class) {
					label = $"Class {symbol.MetadataName}";
					kind = TypeKind.Class.ToString();
				}
            }
			if (symbol is IFieldSymbol fieldSymbol || symbol is ILocalSymbol localSymbol) {
                SetVariableItem(symbol, ref label, ref description, ref returnType, ref kind);
			}

            var contents = new List<VBContent>();
			if (description != "") {
                contents.Add(new() {
					Language = "xml",
                    Value = description,
				});
			}
			if (label != "") {
				contents.Add(new() {
					Language = "vb",
					Value = label,
				});
			}
			if (returnType != "") {
				contents.Add(new() {
					Language = "vb",
					Value = $"@return {returnType}",
				});
			}
			if (kind != "") {
				contents.Add(new() {
					Language = "vb",
					Value = $"@kind {kind}",
				});
			}

			var hover = new VBAHover {
				Start = (line, chara),
				Contents = contents
			};
            return hover;
		}

		private void SetVariableItem(ISymbol symbol, ref string label, ref string description, ref string returnType, ref string kind) {
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

			label = dispText;
			description = symbol.GetDocumentationCommentXml();
			kind = symbol.Kind.ToString();
			returnType = typeName;
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

        public async Task<List<VBADiagnostic>> GetDiagnostics(string name) {
			if (!doc_id_dict.TryGetValue(name, out DocumentId docId)) {
				return [];
			}
			var doc = workspace.CurrentSolution.GetDocument(docId);
			vbaDiagnosticProvider.ignoreDs = _preprocVBA.GetIgnoreDiagnostics(name);
            var prepDiagnosticList = _preprocVBA.GetDiagnostics(name);
			var diagnosticList = await vbaDiagnosticProvider.GetDiagnostics(doc);
            var items = diagnosticList.Concat(prepDiagnosticList);
            return [..items];
		}

        public async Task<List<VBALocation>> GetReferences(
			string name, int line, int chara) {
            if (!doc_id_dict.TryGetValue(name, out DocumentId value)) {
                return [];
            }
            var docId = value;
            if (!workspace.CurrentSolution.ContainsDocument(docId)) {
                return [];
            }

            var doc = workspace.CurrentSolution.GetDocument(docId);
			var adjChara = GetCharaDiff(name, line, chara) + chara;
			var position = GetPosition(doc, line, adjChara);
            if (position < 0) {
                return [];
            }

            var model = await doc.GetSemanticModelAsync();
            var symbol = await SymbolFinder.FindSymbolAtPositionAsync(model, position, workspace);
            if (symbol == null) {
                return [];
            }

            var vbaLocations = new List<VBALocation>();
			if (symbol.IsDefinition) {
                foreach (var loc in symbol.Locations) {
                    var filePath = loc.SourceTree.FilePath;
                    var start = loc.GetLineSpan().StartLinePosition;
                    var end = loc.GetLineSpan().EndLinePosition;
					var adjStart = AdjustPosition(filePath, start.Line, start.Character);
					var adjEnd = AdjustPosition(filePath, end.Line, end.Character);
					var uri = new Uri(filePath);
                    vbaLocations.Add(new() {
                        Uri = uri,
                        Start = adjStart,
                        End = adjEnd
					});
				}
            }

            var refItems = SymbolFinder.FindReferencesAsync(symbol, workspace.CurrentSolution).Result;
            foreach (var refItem in refItems) {
                foreach (var loc in refItem.Locations) {
                    var filePath = loc.Document.FilePath;
                    var start = loc.Location.GetLineSpan().StartLinePosition;
                    var end = loc.Location.GetLineSpan().EndLinePosition;
					var adjStart = AdjustPosition(filePath, start.Line, start.Character);
					var adjEnd = AdjustPosition(filePath, end.Line, end.Character);
					var uri = new Uri(filePath);
					vbaLocations.Add(new() {
						Uri = uri,
						Start = adjStart,
						End = adjEnd
					});
				}
            }
            return vbaLocations;
        }

        public List<VBADocSymbol> GetDocumentSymbols(string name, Uri uri) {
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

		private (int, int) AdjustPosition(string filePath, int vbaLine, int vbaChara) {
			var charaDiff = GetCharaDiff(filePath, vbaLine, vbaChara);
			var line = vbaLine;
			var chara = vbaChara - charaDiff;
			if (line < 0) {
				line = 0;
			}
			if (chara < 0) {
				chara = 0;
			}
			return (line, chara);
		}
	}
}
