using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System.Collections.Generic;
using System.Linq;

namespace VBACodeAnalysis {
    public class Rewrite {
        private RewriteSetting setting;
        public Dictionary<int, (int, int)> charaOffsetDict;
        public Dictionary<int, int> lineMappingDict;

        public Rewrite(RewriteSetting setting) {
            this.setting = setting;
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
