using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VBACodeAnalysis {
    using locDiffDict = Dictionary<int, List<LocationDiff>>;

    public class Rewrite {
        private RewriteSetting setting;
        public Dictionary<int, int> lineMappingDict;
        public locDiffDict locationDiffDict;

        public Rewrite(RewriteSetting setting) {
            this.setting = setting;
            lineMappingDict = new Dictionary<int, int>();
            locationDiffDict = new locDiffDict();
        }

        public SourceText RewriteStatement(Document doc) {
            lineMappingDict = new Dictionary<int, int>();
            locationDiffDict = new locDiffDict();

            var docRoot = doc.GetSyntaxRootAsync().Result;
            var nodes = docRoot.DescendantNodes();

            var prepChanges = Prep(nodes);
            docRoot = docRoot.SyntaxTree
			    .WithChangedText(docRoot.GetText().WithChanges(prepChanges))
			    .GetRootAsync().Result;

            var rewriteProp = new RewriteProperty();
            docRoot = rewriteProp.Rewrite(docRoot);
            lineMappingDict = rewriteProp.lineMappingDict;

            var typeResult = TypeStatement(docRoot);
            docRoot = typeResult.root;
            ApplyLocationDict(ref locationDiffDict, typeResult.dict);

            docRoot = AnnotationAs(docRoot);

            var result = VBAClassToFunction(docRoot);
            docRoot = result.root;
            ApplyLocationDict(ref locationDiffDict, result.dict);

            return docRoot.GetText();
        }

        public void ApplyLocationDict(ref locDiffDict srcDict, locDiffDict dict) {
            foreach (var item in dict) {
                if (srcDict.ContainsKey(item.Key)) {
                    var mlist = srcDict[item.Key];
                    mlist.Sort((a, b) => {
                        return a.Chara - b.Chara;
                    });
                    foreach (var item2 in item.Value) {
                        var sumDiff = 0;
                        foreach (var item3 in mlist) {
                            if (item3.Chara < item2.Chara) {
                                sumDiff += item3.Diff;
                            } else {
                                break;
                            }
                        }
                        item2.Chara -= sumDiff;
                    }
                    mlist.AddRange(item.Value);
                } else {
                    srcDict.Add(item.Key, item.Value);
                }
            }
            foreach (var item in srcDict) {
                item.Value.Sort((a, b) => {
                    return a.Chara - b.Chara;
                });
            }
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

        public (SyntaxNode root, locDiffDict dict) VBAClassToFunction(SyntaxNode root) {
            var ns = setting.NameSpace;
            var rewriteDict = setting.getRewriteDict();

            var locDiffDict = new locDiffDict();
            var lookup = new Dictionary<SyntaxNode, SyntaxNode>();

            var node = root.DescendantNodes();
            var invExpStmts = node.OfType<InvocationExpressionSyntax>();
            foreach (var stmt in invExpStmts) {
                var fiestToken = stmt.GetFirstToken();
                var text = fiestToken.Text;
                if (!rewriteDict.ContainsKey(text)) {
                    continue;
                }
                var oldName = stmt.ToString();
                var newName = $"{ns}.{oldName}";
                var newNode = SyntaxFactory.ParseExpression(newName)
					.WithLeadingTrivia(stmt.GetLeadingTrivia())
					.WithTrailingTrivia(stmt.GetTrailingTrivia());
                lookup.Add(stmt, newNode);

				var lp = stmt.GetLocation().GetLineSpan();
                var slp = lp.StartLinePosition;
                if (!locDiffDict.ContainsKey(slp.Line)) {
                    locDiffDict.Add(slp.Line, new List<LocationDiff>());
                }
                if(!locDiffDict[slp.Line].Exists(x => x.Line == slp.Line && x.Chara == slp.Character)) {
                    locDiffDict[slp.Line].Add(new LocationDiff(slp.Line, slp.Character, newName.Length - oldName.Length));
                }
            }
            
            var keys = rewriteDict.Keys.Select(x => x.ToLower());
            var forStmt = node.OfType<ForEachStatementSyntax>();
            foreach (var stmt in forStmt) {
                var fiestToken = stmt.Expression.GetFirstToken();
                var fiestNode = stmt.FindNode(fiestToken.Span);
                if (!keys.Contains(fiestNode.ToString().ToLower())) {
                    continue;
				}

                var text = fiestNode.ToString();
                var newName = $"{ns}.{text}";
                var newNode = SyntaxFactory.ParseExpression(newName)
                    .WithLeadingTrivia(fiestNode.GetLeadingTrivia())
                    .WithTrailingTrivia(fiestNode.GetTrailingTrivia());
                lookup.Add(fiestNode, newNode);
                var lp = fiestNode.GetLocation().GetLineSpan();
                var slp = lp.StartLinePosition;
                if (!locDiffDict.ContainsKey(slp.Line)) {
                    locDiffDict.Add(slp.Line, new List<LocationDiff>());
                }
                if (!locDiffDict[slp.Line].Exists(x => x.Line == slp.Line && x.Chara == slp.Character)) {
                    locDiffDict[slp.Line].Add(new LocationDiff(slp.Line, slp.Character, newName.Length - text.Length));
                }
            }
            var repNode = root.ReplaceNodes(lookup.Keys, (s, d) => lookup[s]);
            return (repNode, locDiffDict);
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

        public SyntaxNode AnnotationAs(SyntaxNode root) {
            var lookup = new Dictionary<SyntaxNode, SyntaxNode>();
            var items = root.DescendantNodes().OfType<VariableDeclaratorSyntax>();
            foreach (var item in items) {
                var comments = item.Parent.GetLeadingTrivia().Where(x => x.IsKind(SyntaxKind.CommentTrivia));
                if (!comments.Any()) {
                    continue;
                }
                var value = comments.First().ToString();
                var mc = Regex.Match(value, @"@as\s+(\S+)", RegexOptions.IgnoreCase);
                if (!mc.Success) {
                    continue;
                }
				if (!item.AsClause.ChildNodes().Any()) {
                    continue;
                }
                var typeName = mc.Groups[1].Value;
                var oldNode = item.AsClause.ChildNodes().First();
                var newNode = SyntaxFactory.IdentifierName(typeName)
                    .WithTrailingTrivia(oldNode.GetTrailingTrivia());
                lookup.Add(oldNode, newNode);
            }
            if(lookup.Count == 0) {
                return root;
            }
            var repNode = root.ReplaceNodes(lookup.Keys, (s, d) => lookup[s]);
            var repRoot  = root.SyntaxTree.WithChangedText(repNode.GetText());
            return repRoot.GetRootAsync().Result;
        }

        public (SyntaxNode root, locDiffDict dict) TypeStatement(SyntaxNode root) {
            var locDiffDict = new locDiffDict();
            // 'Type' ステートメントはサポートされなくなりました
            // 'Structure' ステートメントを使用してください
            var code = "BC30802";
			var diags = root.GetDiagnostics().Where(x => x.Id == code);
			if (!diags.Any()) {
				return (root, locDiffDict);
			}

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
                var value = item.ToFullString();
                var mc1 = Regex.Match(
                        value, @"\s+(type)\s+",
                        RegexOptions.IgnoreCase);
                var mc2 = Regex.Match(
                        value, @"^(type)\s+",
                        RegexOptions.IgnoreCase);
                if (!mc1.Success && !mc2.Success) {
					continue;
				}
                Match mc = null;
                if (mc1.Success) {
                    mc = mc1;
                }
                if (mc2.Success) {
                    mc = mc2;
                }
                var pre = value.Substring(0, mc.Groups[1].Index);
                var post = value.Substring(mc.Groups[1].Index + mc.Groups[1].Length);
                var newValue = $"{pre}Structure{post}";
                if (item.IsKind(SyntaxKind.EmptyToken)) {
                    var rep = SyntaxFactory.Token(SyntaxKind.EmptyToken, newValue);
                    lookup.Add(item, rep);
                } else {
                    var rep = SyntaxFactory.Token(item.Kind(), newValue);
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
                    if (!locDiffDict.ContainsKey(sp.Line)) {
                        locDiffDict.Add(sp.Line, new List<LocationDiff>());
                    }
                    locDiffDict[sp.Line].Add(
                        new LocationDiff(sp.Line, sp.Character + "type".Length, diff));
                }
            }

            if (lookup.Count == 0) {
                return (root, locDiffDict);
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
                        var menTrivia = menber.GetTrailingTrivia();
                        var skippedTokens = menTrivia.TakeWhile(x => !x.IsKind(SyntaxKind.SkippedTokensTrivia));
                        var trailingTrivia = menTrivia.Skip(skippedTokens.Count()).Take(menTrivia.Count);
                        var rep = SyntaxFactory.Token(SyntaxKind.EmptyToken, $"Public ")
                            .WithLeadingTrivia(skippedTokens)
                            .WithTrailingTrivia(trailingTrivia);
                        lookupMenber.Add(token, rep);
                        var menToken = trailingTrivia.First().GetLocation();
                        var lp = menToken.GetLineSpan();
                        var sp = lp.StartLinePosition;
                        if (!locDiffDict.ContainsKey(sp.Line)) {
                            locDiffDict.Add(sp.Line, new List<LocationDiff>());
                        }
                        locDiffDict[sp.Line].Add(
                            new LocationDiff(sp.Line, sp.Character, "Public ".Length));
                    } else if (token.IsKind(SyntaxKind.PublicKeyword)
                            || token.IsKind(SyntaxKind.PrivateKeyword)) {
                            var rep = SyntaxFactory.Token(SyntaxKind.EmptyToken, $"{token} {token}")
                                .WithLeadingTrivia(token.LeadingTrivia)
                                .WithTrailingTrivia(token.TrailingTrivia);
                            lookupMenber.Add(token, rep);
                        var lp = token.GetLocation().GetLineSpan();
                        var sp = lp.StartLinePosition;
                        if (!locDiffDict.ContainsKey(sp.Line)) {
                            locDiffDict.Add(sp.Line, new List<LocationDiff>());
                        }
                        locDiffDict[sp.Line].Add(
                            new LocationDiff(sp.Line, sp.Character, "Public ".Length));
                    }
                }
            }
            if(lookupMenber.Count == 0) {
                return (repNode, locDiffDict);
            }
			try {
                repNode = repNode.ReplaceTokens(lookupMenber.Keys, (s, d) => lookupMenber[s]);
            } catch (System.Exception) {
                return (root, new locDiffDict());
			}

            var repTree = repNode.SyntaxTree.WithChangedText(repNode.GetText());
            return (repTree.GetRootAsync().Result, locDiffDict);
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
