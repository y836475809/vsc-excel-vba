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
            var propResult = rewriteProp.Rewrite(docRoot);
            docRoot = propResult.root;
            MergeLocationDiffDict(ref locationDiffDict, propResult.dict);
            lineMappingDict = rewriteProp.lineMappingDict;

            var typeResult = TypeStatement(docRoot);
            docRoot = typeResult.root;
            MergeLocationDiffDict(ref locationDiffDict, typeResult.dict);

            docRoot = AnnotationAs(docRoot);

            var result = VBAClassToFunction(docRoot);
            docRoot = result.root;
            MergeLocationDiffDict(ref locationDiffDict, result.dict);

			var predefinedResult = Predefined(docRoot);
            docRoot = predefinedResult.root;
            MergeLocationDiffDict(ref locationDiffDict, predefinedResult.dict);

			return docRoot.GetText();
        }

        public void MergeLocationDiffDict(ref locDiffDict srcDict, locDiffDict inDict) {
            foreach (var inItem in inDict) {
                if (srcDict.ContainsKey(inItem.Key)) {
                    var srclist = srcDict[inItem.Key];
                    srclist.Sort((a, b) => {
                        return a.Chara - b.Chara;
                    });
                    var cloneList = srclist.Select(x => x.Clone()).ToList();
                    var diff = cloneList[0].Diff;
                    for (int i = 1; i < cloneList.Count(); i++) {
                        cloneList[i].Chara += diff;
                        diff += cloneList[i].Diff;
                    }

                    foreach (var inLocDiff in inItem.Value) {
                        var sumDiff = 0;
                        foreach (var clineLocDiff in cloneList) {
                            if (clineLocDiff.Chara < inLocDiff.Chara) {
                                sumDiff += clineLocDiff.Diff;
                            } else {
                                break;
                            }
                        }
                        inLocDiff.Chara -= sumDiff;
                    }
                    srclist.AddRange(inItem.Value);
                } else {
                    srcDict.Add(inItem.Key, inItem.Value);
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
            var mn = setting.VBAClassToFunction.ModuleName;
            var vbaClasses = setting.VBAClassToFunction.VBAClasses;

            var locDiffDict = new locDiffDict();
            var lookup = new Dictionary<SyntaxNode, SyntaxNode>();

            var node = root.DescendantNodes();
            var invExpStmts = node.OfType<InvocationExpressionSyntax>();
            foreach (var stmt in invExpStmts) {
                var firstToken = stmt.GetFirstToken();
                var text = firstToken.Text;
                if (!vbaClasses.Contains(text.ToLower())) {
                    continue;
                }
                var childNodes = stmt.ChildNodes();
                if (!childNodes.Any()) {
                    continue;
                }
                var firstNode = childNodes.First();
                var oldName = firstNode.ToString();
                var newName = $"{mn}.{oldName}";
                var newNode = SyntaxFactory.ParseExpression(newName)
					.WithLeadingTrivia(firstNode.GetLeadingTrivia())
					.WithTrailingTrivia(firstNode.GetTrailingTrivia());
                if (lookup.ContainsKey(firstNode)) {
                    continue;
                }
                lookup.Add(firstNode, newNode);

                var lp = firstNode.GetLocation().GetLineSpan();
                var slp = lp.StartLinePosition;
                if (!locDiffDict.ContainsKey(slp.Line)) {
                    locDiffDict.Add(slp.Line, new List<LocationDiff>());
                }
                if(!locDiffDict[slp.Line].Exists(x => x.Line == slp.Line && x.Chara == slp.Character)) {
                    locDiffDict[slp.Line].Add(new LocationDiff(slp.Line, slp.Character, newName.Length - oldName.Length));
                }
            }
            
            var forStmt = node.OfType<ForEachStatementSyntax>();
            foreach (var stmt in forStmt) {
                var firstToken = stmt.Expression.GetFirstToken();
                var firstNode = stmt.FindNode(firstToken.Span);
                if (!vbaClasses.Contains(firstToken.ToString().ToLower())) {
                    continue;
				}

                var text = firstNode.ToString();
                var newName = $"{mn}.{text}";
                var newNode = SyntaxFactory.ParseExpression(newName)
                    .WithLeadingTrivia(firstNode.GetLeadingTrivia())
                    .WithTrailingTrivia(firstNode.GetTrailingTrivia());
                if (lookup.ContainsKey(firstNode)) {
                    continue;
                }
                lookup.Add(firstNode, newNode);
                var lp = firstNode.GetLocation().GetLineSpan();
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

        public (SyntaxNode root, locDiffDict dict) Predefined(SyntaxNode root) {
            var nm = this.setting.VBAPredefined.ModuleName;
            var dict = new locDiffDict();
            var lookup = new Dictionary<SyntaxToken, SyntaxToken>();
            var items = root.DescendantNodes().OfType<PredefinedCastExpressionSyntax>();
            foreach (var item in items) {
                var fiestToken = item.GetFirstToken();
                var oldNode = item.FindNode(fiestToken.Span);
                var text = fiestToken.Text;      
                var newName = $"{nm}.{text}";
                var newToken = SyntaxFactory.Token(
                    SyntaxKind.KeyKeyword, newName)
                    .WithLeadingTrivia(oldNode.GetLeadingTrivia());
                lookup.Add(fiestToken, newToken);

                var ls = fiestToken.GetLocation().GetLineSpan();
                var slp = ls.StartLinePosition;
				if (!dict.ContainsKey(slp.Line)) {
                    dict.Add(slp.Line, new List<LocationDiff>());
                }
                dict[slp.Line].Add(
                    new LocationDiff(slp.Line, slp.Character, newName.Length-text.Length));
            }
            if (lookup.Count == 0) {
                return (root, new locDiffDict());
            }
            var repNode = root.ReplaceTokens(lookup.Keys, (s, d) => lookup[s]);
            var tree = root.SyntaxTree.WithChangedText(repNode.GetText());
            return (tree.GetRootAsync().Result, dict);
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
