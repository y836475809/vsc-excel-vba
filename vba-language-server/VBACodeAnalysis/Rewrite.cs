using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VBACodeAnalysis {
    public class Rewrite {
        private RewriteSetting setting;
        public Dictionary<int, (int, int)> charaOffsetDict;
        public Dictionary<int, int> lineMappingDict;

        public Rewrite(RewriteSetting setting) {
            this.setting = setting;
            charaOffsetDict = new Dictionary<int, (int, int)>();
            lineMappingDict = new Dictionary<int, int>();
        }

        public SourceText RewriteStatement(Document doc) {
            var docRoot = doc.GetSyntaxRootAsync().Result;
            var nodes = docRoot.DescendantNodes();

            var prepChanges = Prep(nodes);
            docRoot = docRoot.SyntaxTree
			    .WithChangedText(docRoot.GetText().WithChanges(prepChanges))
			    .GetRootAsync().Result;

            var rewriteProp = new RewriteProperty();
            docRoot = rewriteProp.Rewrite(docRoot);
            charaOffsetDict = rewriteProp.charaOffsetDict;
            lineMappingDict = rewriteProp.lineMappingDict;

            docRoot = TypeStatement(docRoot);

           var updatedDoc =  doc.WithSyntaxRoot(docRoot);
            nodes = docRoot.DescendantNodes();
            var changes = VBAClassToFunction(updatedDoc, nodes);
            return docRoot.GetText().WithChanges(changes);
        }

        private List<TextChange> Prep(IEnumerable<SyntaxNode> nodes) {
            var changes = SetStatement(nodes);
            changes = changes.Concat(AsClauseStatement(nodes)).ToList();

            var changeDict = new Dictionary<string, TextChange>();
            foreach (var item in changes) {
                var key = item.Span.ToString();
                changeDict[key] = item;
            }
            return changeDict.Values.ToList();
        }

        public List<TextChange> VBAClassToFunction(
                    Document doc, IEnumerable<SyntaxNode> node) {
            SemanticModel model = null;
            var ns = setting.NameSpace;
            var rewriteDict = setting.getRewriteDict();

            var allChanges = new List<TextChange>();
            var set = new HashSet<string>();

            var invExpStmts = node.OfType<InvocationExpressionSyntax>();
            foreach (var stmt in invExpStmts) {
                var fiestToken = stmt.GetFirstToken();
                var text = fiestToken.Text;
                if (!rewriteDict.ContainsKey(text)) {
                    continue;
                }
                if (model == null) {
                    model = doc.GetSemanticModelAsync().Result;
                }
                var pos = fiestToken.GetLocation().SourceSpan.Start + 1;
                var symbol = SymbolFinder.FindSymbolAtPositionAsync(doc, pos).Result;
                if (!IsClass(symbol, text)) {
                    continue;
                }
                var sp = stmt.GetFirstToken().Span;
                var key = $"{sp.Start}-{sp.End}";
                if (!set.Contains(key)) {
                    var rename = $"{ns}.{rewriteDict[text]}";
                    allChanges.Add(new TextChange(stmt.GetFirstToken().Span, rename));
                }
                set.Add(key);
            }

            var forStmt = node.OfType<ForEachStatementSyntax>();
            foreach (var stmt in forStmt) {
                var identTokens = stmt.ChildNodes().Where(x => {
                    return rewriteDict.ContainsKey(x.ToString());
                });
                foreach (var ident in identTokens) {
                    var sp = ident.Span;
                    var key = $"{sp.Start}-{sp.End}";
                    if (!set.Contains(key)) {
                        var rename = $"{ns}.{rewriteDict[ident.ToString()]}";
                        allChanges.Add(new TextChange(sp, rename));
                    }
                    set.Add(key);
                }
            }
            return allChanges;
        }

        private bool IsClass(ISymbol symbol, string name) {
            if (symbol is INamedTypeSymbol namedType) {
                if (namedType.TypeKind == TypeKind.Class) {
                    if (namedType.Name == name) {
                        return true;
                    }
                }
            }
            return false;
        }

        private List<TextChange> SetStatement(IEnumerable<SyntaxNode> node) {
            var allChanges = new List<TextChange>();

            var emptyStmts = node.OfType<EmptyStatementSyntax>();
            // 'Let' および 'Set' 代入ステートメントはサポートされなくなりました。
            const string code = "BC30807";
			foreach (var stmt in emptyStmts) {
				var ds = stmt.Empty.TrailingTrivia.Where(x => {
					return x.GetDiagnostics().SingleOrDefault(x => x.Id == code) != null;
				});
				var changes = ds.Select(x => {
					return new TextChange(x.Span, new string(' ', x.Span.Length));
				});
			    if (changes.Any()) {
					allChanges = allChanges.Concat(changes).ToList();
				}
			}

			var setAccBlockStmt = node.Where(x => x.IsKind(SyntaxKind.SetAccessorBlock));
            foreach (var stmt in setAccBlockStmt) {
                var changes = stmt.ChildNodes()
                    .Where(x => x.IsKind(SyntaxKind.SetAccessorStatement))
                    .Select(x => {
                        return new TextChange(x.Span, new string(' ', x.Span.Length));
                });
				if (changes.Any()) {
                    var mm = allChanges.Concat(changes);
                    allChanges.Concat(mm);

                    allChanges = allChanges.Concat(changes).ToList();
				}
            }

            {
                var setAccStmt = node.Where(x => x.IsKind(SyntaxKind.SetAccessorStatement));
                var changes = setAccStmt.Select(x => new TextChange(x.Span, new string(' ', x.Span.Length)));
                if (changes.Any()) {
                    allChanges = allChanges.Concat(changes).ToList();
                }
            }
			return allChanges;
		}

        public SyntaxNode TypeStatement(SyntaxNode root) {
			// 'Type' ステートメントはサポートされなくなりました
			// 'Structure' ステートメントを使用してください
			var code = "BC30802";
			var diags = root.GetDiagnostics().Where(x => x.Id == code);
			if (!diags.Any()) {
				return root;
			}

			var typeCharaOffsetDict = new Dictionary<int, (int, int)>();
            
            var typeTokens = root.DescendantTokens().Where(x => {
                if (x.IsKind(SyntaxKind.EndKeyword) 
                    || x.IsKind(SyntaxKind.EmptyToken)
                    || x.IsKind(SyntaxKind.PublicKeyword)
                    || x.IsKind(SyntaxKind.PrivateKeyword)) {
                    var trivias = x.TrailingTrivia.Select(y => y.ToString().ToLower().Trim());
                    return trivias.Any(x => x.Contains("type"));
                }
                return false;
            });

            var lookup = new Dictionary<SyntaxToken, SyntaxToken>();
            foreach (var item in typeTokens) {
                if (item.IsKind(SyntaxKind.EmptyToken)) {
                    var rep = SyntaxFactory.Token(SyntaxKind.EmptyToken,
                        Regex.Replace(item.ToFullString(), "type", "Structure", RegexOptions.IgnoreCase));
                    lookup.Add(item, rep);
                } else {
                    var rep = SyntaxFactory.Token(item.Kind(),
                        Regex.Replace(item.ToFullString(), "type", "Structure", RegexOptions.IgnoreCase));
                    lookup.Add(item, rep);
                }
                var trivias = item.TrailingTrivia.Where(
                    x => x.ToString().ToLower().Trim().IndexOf("type") >= 0);
				if (trivias.Any()) {
                    var lp = trivias.First().GetLocation().GetLineSpan();
                    var ltrivia = item.LeadingTrivia.ToFullString();
                    var diff = "Structure".Length - "type".Length;
                    var sp = lp.StartLinePosition;
                    var ep = lp.EndLinePosition;
                    typeCharaOffsetDict[sp.Line] =
                        (sp.Character + "type".Length, diff);
                }
            }

            if (lookup.Count == 0) {
                return root;
            }

            var lookupMenber = new Dictionary<SyntaxToken, SyntaxToken>();
            var repNode = root.ReplaceTokens(lookup.Keys, (s, d) => lookup[s]);
            repNode = root.SyntaxTree.WithChangedText(repNode.GetText()).GetRootAsync().Result;
            var stNodes = repNode.DescendantNodes().OfType<StructureBlockSyntax>();
            foreach (var st in stNodes) {
                foreach (var menber in st.Members) {
                    var menberTokens = menber.ChildTokens();
					if (!menberTokens.Any()) {
                        continue;
					}
                    var token = menberTokens.First();
					if (token.IsKind(SyntaxKind.EmptyToken)) {
                        var rep = SyntaxFactory.Token(SyntaxKind.EmptyToken,
                            $"Public {menber.GetTrailingTrivia()}");
                        lookupMenber.Add(token, rep);
                    } else {
                        var rep = SyntaxFactory.Token(SyntaxKind.EmptyToken,
                            $"Public {token.ToFullString()}");
                        lookupMenber.Add(token, rep);
                    }
                    var lp = token.GetLocation().GetLineSpan();
                    var sp = lp.StartLinePosition;
                    typeCharaOffsetDict[sp.Line] = (sp.Character,  "Public ".Length);
                }
            }
            if(lookupMenber.Count == 0) {
				foreach (var item in typeCharaOffsetDict) {
                    charaOffsetDict[item.Key] = item.Value;
                }
                return repNode;
            }
			try {
                repNode = repNode.ReplaceTokens(lookupMenber.Keys, (s, d) => lookupMenber[s]);
            } catch (System.Exception) {
                return root;
			}

            updateCharaOffsetDict(typeCharaOffsetDict);

            var repTree = repNode.SyntaxTree.WithChangedText(repNode.GetText());
            return repTree.GetRootAsync().Result;
        }

        private void updateCharaOffsetDict(Dictionary<int, (int, int)> dict) {
            foreach (var item in dict) {
                charaOffsetDict[item.Key] = item.Value;
            }
        }

        private List<TextChange> LocalDeclarationStatement(IEnumerable<SyntaxNode> node) {
            var allChanges = new List<TextChange>();
            var forStmt = node.OfType<LocalDeclarationStatementSyntax>();
            const string code = "BC30804";
            foreach (var stmt in forStmt) {
                var ds = stmt.GetDiagnostics().Where(x => {
                    return x.Id == code;
                });
                var changes = ds.Select(x => {
                    return new TextChange(x.Location.SourceSpan, "Object ");
                });
                if (changes.Count() > 0) {
                    allChanges = allChanges.Concat(changes).ToList();
                }
            }
            return allChanges;
        }

        private List<TextChange> FieldDeclarationStatement(IEnumerable<SyntaxNode> node) {
            var allChanges = new List<TextChange>();
            var forStmt = node.OfType<FieldDeclarationSyntax>();
            const string code = "BC30804";
            foreach (var stmt in forStmt) {
                var ds = stmt.GetDiagnostics().Where(x => {
                    return x.Id == code;
                });
                var changes = ds.Select(x => {
                    return new TextChange(x.Location.SourceSpan, "Object ");
                });
                if (changes.Count() > 0) {
                    allChanges = allChanges.Concat(changes).ToList();
                }
            }
            return allChanges;
        }

        private List<TextChange> AsClauseStatement(IEnumerable<SyntaxNode> node) {
            var allChanges = new List<TextChange>();
            var forStmt = node.OfType<AsClauseSyntax>();
            const string code = "BC30804";
            foreach (var stmt in forStmt) {
                var ds = stmt.GetDiagnostics().Where(x => {
                    return x.Id == code;
                });
                var changes = ds.Select(x => {
                    return new TextChange(x.Location.SourceSpan, "Object ");
                });
                if (changes.Count() > 0) {
                    allChanges = allChanges.Concat(changes).ToList();
                }
            }
            return allChanges;
        }
    }
}
