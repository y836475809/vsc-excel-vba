using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1 {
    public class Rewrite {
        private RewriteSetting setting;
        public Rewrite(RewriteSetting setting) {
            this.setting = setting;
        }

        public SourceText RewriteStatement(Document doc) {
            var docRoot = doc.GetSyntaxRootAsync().Result;
            var nodes = docRoot.DescendantNodes();
            var changes = new List<TextChange>();

            docRoot = PropertyStatement(docRoot);
            nodes = docRoot.DescendantNodes();

            changes = changes.Concat(ReplaceStatement(nodes)).ToList();
            changes = changes.Concat(SetStatement(nodes)).ToList();
            changes = changes.Concat(LocalDeclarationStatement(nodes)).ToList();
            changes = changes.Concat(FieldDeclarationStatement(nodes)).ToList();

            return docRoot.GetText().WithChanges(changes);
        }

        private SyntaxNode PropertyStatement(SyntaxNode docRoot) {
            var node = docRoot.DescendantNodes();
            var allChanges = new List<TextChange>();

            var forStmt2 = node.OfType<EndBlockStatementSyntax>()
                .Where(x => x.IsKind(SyntaxKind.EndPropertyStatement));
            foreach (var stmt in forStmt2) {
                var chTokens = stmt.ChildTokens();
                var propToken = chTokens.Where(x => x.IsKind(SyntaxKind.PropertyKeyword)).FirstOrDefault();
				if (propToken.IsKind(SyntaxKind.None)) {
                    continue;
				}
                allChanges.Add(
                    new TextChange(propToken.Span, "Function"));
            }

            var lines = docRoot.GetText().Lines;
            var forStmt = node.OfType<PropertyStatementSyntax>();

            var repDict = new SortedDictionary<int, string>();
            var propNameSet = new HashSet<string>();
            var repLetDict = new SortedDictionary<int, string>();
            var repGetDict = new SortedDictionary<int, (string, string)>();
            foreach (var stmt in forStmt) {
                var chTokens = stmt.ChildTokens();
                var propToken = chTokens.Where(x => x.IsKind(SyntaxKind.PropertyKeyword)).FirstOrDefault();
                if (propToken.IsKind(SyntaxKind.None)) {
                    continue;
                }
                var asTokens = stmt.ChildNodes()
                    .Where(x => x.IsKind(SyntaxKind.SimpleAsClause));
                var propNameToken = propToken.GetNextToken().TrailingTrivia
                    .Where(x => x.IsKind(SyntaxKind.SkippedTokensTrivia)).FirstOrDefault();
                if (propNameToken.IsKind(SyntaxKind.None)) {
                    continue;
                }
                var propName = propNameToken.ToString();
                if (asTokens.Any()) {
                    var lp = propToken.GetLocation().GetLineSpan();
                    var propLine = lp.StartLinePosition.Line - 1;
                    if (lines[propLine].Span.Length > 0) {
						propNameSet.Add(propName);
						continue;
                    }
                    var asdef = asTokens.First().ToString();
                    if (!propNameSet.Contains(propName)) {
                        var insp = docRoot.GetText().Lines[propLine].Span;
                        var insText = $"Public Property {propName} {asdef}";
                        repDict.Add(propLine, insText);
                    }
                    //propNameSet.Add(propName);

                    var insp2 = docRoot.GetText().Lines[propLine + 1].Span;
                    var mk = $"Private Function Get{propName}() {asdef}";
                    allChanges.Add(new TextChange(
                        new TextSpan(insp2.Start, mk.Length), mk));
                    repGetDict.Add(propLine + 1, (propName, $"Get{propName}"));
				} else {
                    var lp = propToken.GetLocation().GetLineSpan();
                    var propLine = lp.StartLinePosition.Line - 1;
                    //var propName = propNameToken.ToString();
                    if (lines[propLine].Span.Length > 0) {
						propNameSet.Add(propName);
						continue;
                    }
                    var propArgs = stmt.ChildNodes().Where(x => x.IsKind(SyntaxKind.ParameterList));
                    if (propArgs.Any()) {
                        var paramsSym = propArgs.First() as ParameterListSyntax;
                        var paramsAs = paramsSym.Parameters;
                        if (paramsAs.Any()) {
                            var asClause = paramsAs.First().ChildNodes().Where(
                                x => x.IsKind(SyntaxKind.SimpleAsClause)).FirstOrDefault();
							if (!asClause.IsKind(SyntaxKind.None)) {
                                var asClauseType = (asClause as SimpleAsClauseSyntax).Type;
                                if (!propNameSet.Contains(propName)) {
                                    var insp = docRoot.GetText().Lines[propLine].Span;
                                    var insText = $"Public Property {propName} As {asClauseType}";
                                    repDict.Add(propLine, insText);
                                }
                                //propNameSet.Add(propName);
                            }
                        }
                    }
                    var chSpan = docRoot.GetText().Lines[propLine + 1].Span;
                    var chText = $"Private Function Let{propName}";
                    allChanges.Add(new TextChange(
                        new TextSpan(chSpan.Start, chText.Length), chText));
                    repLetDict.Add(propLine + 1, "let");
                }
                propNameSet.Add(propName);
            }
			if (!repDict.Any()) {
                return docRoot;
            }

            var cchtext = docRoot.GetText().WithChanges(allChanges);
            docRoot = docRoot.SyntaxTree.WithChangedText(cchtext).GetRootAsync().Result;

            var repOtherDict = new SortedDictionary<int, string>();
            lines = docRoot.GetText().Lines;
            foreach (var line in repGetDict.Keys) {
                var targetNode = docRoot.FindNode(lines[line].Span);
                if(targetNode == null) {
                    continue;
				}
                var assigs = targetNode.Parent.ChildNodes()
					.Where(x => x.IsKind(SyntaxKind.SimpleAssignmentStatement))
                    .Cast<AssignmentStatementSyntax>();
                var (org, pre) = repGetDict[line];
                foreach (var assig in assigs) {
                    if (assig.Left.ToString() != org) {
                        continue;
                    }
                    var assgiLine = assig.GetLocation().GetLineSpan().StartLinePosition.Line;
                    var trivia = assig.GetLeadingTrivia().FirstOrDefault().ToString();
                    repOtherDict.Add(assgiLine, $"{trivia}{pre} = {assig.Right}");
                }
            }
            var otherChanges = new List<TextChange>();
            foreach (var line in repLetDict.Keys) {
                var targetNode = docRoot.FindNode(lines[line].Span);
                if(targetNode == null) {
                    continue;
				}
                var methodNodes = targetNode.Parent.ChildNodes();
                if(methodNodes.Count() < 2) {
                    continue;
				}
                var methodStartEndNodes = new SyntaxNode[] { methodNodes.First(), methodNodes.Last() };
                foreach (var methodNode in methodStartEndNodes) {
                    var funcToken = methodNode.ChildTokens()
                        .Where(x => x.IsKind(SyntaxKind.FunctionKeyword)).FirstOrDefault();
					if (funcToken.IsKind(SyntaxKind.None)) {
                        continue;
					}
                    otherChanges.Add(new TextChange(
                        new TextSpan(funcToken.Span.Start, funcToken.Span.Length), "Sub     "));
                }
            }
			if (otherChanges.Any()) {
                var cchtext2 = docRoot.GetText().WithChanges(otherChanges);
                docRoot = docRoot.SyntaxTree.WithChangedText(cchtext2)
                    .GetRootAsync().Result;
            }

            var sb = new StringBuilder();
			lines = docRoot.GetText().Lines;
			for (int i = 0; i < lines.Count(); i++) {
                if (repDict.ContainsKey(i)) {
					if (lines[i].Span.Length == 0) {
                        sb.AppendLine(repDict[i]);
                    }
                } else {
                    if (repOtherDict.ContainsKey(i)) {
                        sb.AppendLine(repOtherDict[i]);
                    } else {
                        sb.AppendLine(lines[i].ToString());
                    }
                }
            }
            docRoot = docRoot.SyntaxTree.
                WithChangedText(SourceText.From(sb.ToString())).GetRootAsync().Result;

			return docRoot;
		}

		public List<TextChange> ReplaceStatement(IEnumerable<SyntaxNode> node) {
            var ns = setting.NameSpace;
            var rewriteDict = setting.getRewriteDict();
            var allChanges = new List<TextChange>();
            //var docRoot = doc.GetSyntaxRootAsync().Result;
            var forStmt = node.OfType<InvocationExpressionSyntax>();
            var set = new HashSet<string>();
            foreach (var stmt in forStmt) {
                var tt = stmt.GetFirstToken().Text;
                if (rewriteDict.ContainsKey(tt) && stmt.ArgumentList != null) {
                    var sp = stmt.GetFirstToken().Span;
                    var k = $"{sp.Start}-{sp.End}";
                    if (!set.Contains(k)) {
                        var rename = $"{ns}.{rewriteDict[tt]}";
                        allChanges.Add(new TextChange(stmt.GetFirstToken().Span, rename));
                    }
                    set.Add(k);
                }
            }
            return allChanges;
        }

        private List<TextChange> SetStatement(IEnumerable<SyntaxNode> node) {
            var forStmt = node.OfType<EmptyStatementSyntax>();
            var allChanges = new List<TextChange>();
            // 'Let' および 'Set' 代入ステートメントはサポートされなくなりました。
            const string code = "BC30807";
            foreach (var stmt in forStmt) {
                var ds = stmt.Empty.TrailingTrivia.Where(x => {
                    return x.GetDiagnostics().SingleOrDefault(x => x.Id == code) != null;
                });
                var changes = ds.Select(x => {
                    return new TextChange(x.Span, new string(' ', x.Span.Length));
                });
                if (changes.Count() > 0) {
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
    }
}
