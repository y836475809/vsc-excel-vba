using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace VBACodeAnalysis {
    using locDiffDict = Dictionary<int, List<LocationDiff>>;

    public class RewriteProperty {
        // line - (line start, offset)
        // 書き換えで追加された文字を考慮するのが
        // どのラインのどの位置から、どれだけの文字数を調整で追加するか
        public locDiffDict locationDiffDict;

        // line - line
        // 追加したVB.net形式のprop文のライン位置と元の(VBA形式の) prop文のライン位置との対応
        public Dictionary<int, int> lineMappingDict;

		private void Init() {
            locationDiffDict = new locDiffDict();
            lineMappingDict = new Dictionary<int, int>();
        }

        public (SyntaxNode root, locDiffDict dict) Rewrite(SyntaxNode docRoot) {
            Init();

            var (changes, lineGetLetDict) = GetPropStmtChanges(docRoot);
            if (!changes.Any()) {
                return (docRoot,  new locDiffDict());
            }

            docRoot = docRoot.SyntaxTree
                .WithChangedText(docRoot.GetText()
                .WithChanges(changes))
                .GetRootAsync().Result;

            var (PropNamePropStmtLineDict, lineRepStmtDict) = GetReplaceDict(docRoot, lineGetLetDict);
            if (!PropNamePropStmtLineDict.Any() || !lineRepStmtDict.Any()) {
                return (docRoot, new locDiffDict());
            }
           
            docRoot = ApplyReplace(docRoot, PropNamePropStmtLineDict, lineRepStmtDict);
            return (docRoot, locationDiffDict);
        }

        private (List<TextChange>, Dictionary<int, string>) GetPropStmtChanges(SyntaxNode docRoot) {
            var changes = new List<TextChange>();
            // prop文のライン番号 - get or set or let
            // 書き換えたprop文の名前変更や内部の変数名の書き換え位置の取得に使用
            var lineGetLetDict = new Dictionary<int, string>();

            var keywordsGetLet = new List<string>() {
                 "get",  "let", "set"
            };
            var propPairs = new List<(PropertyStatementSyntax, EndBlockStatementSyntax)>();
            var props = docRoot.DescendantTokens()
                .Where(x => x.IsKind(SyntaxKind.PropertyKeyword))
                .Select(x => x.Parent);
            for (int i = 0; i < props.Count() / 2; i++) {
                propPairs.Add((
                    props.ElementAt(i * 2) as PropertyStatementSyntax,
                    props.ElementAt(i * 2 + 1) as EndBlockStatementSyntax));
            }
            foreach (var pair in propPairs) {
                if (pair.Item1 == null || pair.Item2 == null) {
                    continue;
                }
                var tokensGetLet = pair.Item1.ChildTokens().Where(x => keywordsGetLet.Contains(x.ToString().ToLower()));
                if (!tokensGetLet.Any()) {
                    continue;
                }
                var tokenGetLet = tokensGetLet.First();
                lineGetLetDict.Add(
                    tokenGetLet.GetLocation().GetLineSpan().StartLinePosition.Line,
                    tokenGetLet.ToString().ToLower());
                changes.Add(
                    new TextChange(tokenGetLet.Span, new string(' ', tokenGetLet.Span.Length))); ;

                var propTokens = pair.Item1.ChildTokens()
                    .Where(x => x.IsKind(SyntaxKind.PropertyKeyword));
                if (!propTokens.Any()) {
                    continue;
                }
                var prop = propTokens.First();
                if (tokenGetLet.ToString().ToLower() == "get") {
                    changes.Add(new TextChange(prop.Span, "Function"));
                    var propEndToknes = pair.Item2.ChildTokens()
                        .Where(x => x.IsKind(SyntaxKind.PropertyKeyword));
                    if (propEndToknes.Any()) {
                        changes.Add(new TextChange(propEndToknes.First().Span, "Function"));
                    }
                } else {
                    var adj = new string(' ', "Property".Length - "Sub".Length);
                    changes.Add(new TextChange(prop.Span, $"Sub{adj}"));
                    var propEndToknes = pair.Item2.ChildTokens()
                        .Where(x => x.IsKind(SyntaxKind.PropertyKeyword));
                    if (propEndToknes.Any()) {
                        changes.Add(new TextChange(propEndToknes.First().Span, $"Sub{adj}"));
                    }
                }
            }
            return (changes, lineGetLetDict);
        }

        private (Dictionary<string, (string, int)>, Dictionary<int, string>) GetReplaceDict(
            SyntaxNode docRoot, Dictionary<int, string> lineGetLetDict) {
            // Prop文の名前 - VB.net式のProp文, Prop文のライン番号 
            var PropNamePropStmtLineDict = new Dictionary<string, (string, int)>();
            // ライン番号 - ライン番号で置き換える文
            // prop文から書き換えたメソッド、メソッド内部の書き換え位置と文を記録
            var lineRepStmtDict = new Dictionary<int, string>();

            var lines = docRoot.GetText().Lines;
            foreach (var lineNum in lineGetLetDict.Keys) {
                var line = lines[lineNum];
                var propNode = docRoot.FindNode(line.Span);
                if (!(propNode is MethodStatementSyntax mathodnode)) {
                    continue;
                }
                var text = line.ToString();
                var funcName = mathodnode.Identifier.ToString();
                if (mathodnode.IsKind(SyntaxKind.FunctionStatement)) {
					if (!PropNamePropStmtLineDict.ContainsKey(funcName)) {
                        var asCaseType = mathodnode.AsClause?.Type;
                        if(asCaseType == null) {
                            PropNamePropStmtLineDict[funcName] =
                                ($"Public Property {funcName}", lineNum);
                        } else {
                            PropNamePropStmtLineDict[funcName] =
                                ($"Public Property {funcName} As {mathodnode.AsClause.Type}", lineNum);
                        }
                    }
                }
                if (mathodnode.IsKind(SyntaxKind.SubStatement)) {
                    var paramsAs = mathodnode.ParameterList.Parameters;
                    if (paramsAs.Any()) {
                        var asClause = paramsAs.First().ChildNodes().Where(
                            x => x.IsKind(SyntaxKind.SimpleAsClause)).FirstOrDefault();
                        if (!asClause.IsKind(SyntaxKind.None)) {
                            var asClauseType = (asClause as SimpleAsClauseSyntax).Type;
                            if (!PropNamePropStmtLineDict.ContainsKey(funcName)) {
                                PropNamePropStmtLineDict[funcName] = 
                                    ($"Public Property {funcName} As {asClauseType}", lineNum);
                            }
                        }
                    }
                }

                var getlet = lineGetLetDict[lineNum];
                var funcNameIndex = text.IndexOf(funcName);
                if (funcNameIndex > 0) {    
                    if ((funcNameIndex - getlet.Length) >= 0) {
                        text = text.Remove(funcNameIndex - getlet.Length, getlet.Length);
                    }
                }
                // prorp文をSub or Functionに書き換えた文の
                // メソッド名をget or let+メソッド名する
                // また、Privateにする
                text = text.Replace(funcName, $"{getlet}{funcName}");
                if (mathodnode.Modifiers.Any()) {
                    var modify = mathodnode.Modifiers.First();
                    if (modify.IsKind(SyntaxKind.PublicKeyword)) {
                        text = text.Replace("Public", "Private");
                    }
                    var len = "Private".Length - modify.Text.Length;
                    if (!locationDiffDict.ContainsKey(lineNum)) {
                        locationDiffDict.Add(lineNum, new List<LocationDiff>());
                    }
                    locationDiffDict[lineNum].Add(
                        new LocationDiff(lineNum, "Private".Length, len));
                } else {
                    text = $"Private {text}";
                    var len = "Private".Length;
                    if (!locationDiffDict.ContainsKey(lineNum)) {
                        locationDiffDict.Add(lineNum, new List<LocationDiff>());
                    }
                    locationDiffDict[lineNum].Add(
                        new LocationDiff(lineNum, len, len));
                }
				lineRepStmtDict.Add(lineNum, text);

                // prop文を書き換えたメソッド内でメソッド名=値としている箇所を
                // get+メソッド名=値に書き換える
                var propChNodes = propNode.Parent.ChildNodes();
				if (propChNodes.Any()) {
                    var assigs = propChNodes
                        .Where(x => x.IsKind(SyntaxKind.SimpleAssignmentStatement))
                        .Cast<AssignmentStatementSyntax>();
                    if (assigs.Any()) {
                        foreach (var assigStmt in assigs) {
                            if (assigStmt.Left.ToString() != funcName) {
                                continue;
                            }
                            var repText = $"{getlet}{funcName}";
                            var assigLineNum = assigStmt.GetLocation().GetLineSpan().StartLinePosition.Line;
                            //var assigText = lines[assigLineNum].ToString();
                            var assigLeft = assigStmt.Left;
                            var newen = SyntaxFactory.IdentifierName(repText)
                                    .WithLeadingTrivia(assigLeft.GetLeadingTrivia())
                                    .WithTrailingTrivia(assigLeft.GetTrailingTrivia());
                            var newasg = assigStmt.ReplaceNode(assigLeft, newen)
                                    .WithLeadingTrivia(assigStmt.GetLeadingTrivia())
                                    .WithTrailingTrivia(assigStmt.GetTrailingTrivia());

                            lineRepStmtDict.Add(assigLineNum, Regex.Replace(newasg.ToFullString(), Environment.NewLine, ""));
                            var assigTrivia = assigStmt.GetLeadingTrivia().FirstOrDefault().ToString();
                            if (!locationDiffDict.ContainsKey(assigLineNum)) {
                                locationDiffDict.Add(assigLineNum, new List<LocationDiff>());
                            }
                            locationDiffDict[assigLineNum].Add(
                                new LocationDiff(assigLineNum, assigTrivia.Length + repText.Length, repText.Length - funcName.Length));
                        }
					}
				}
            }

            return (PropNamePropStmtLineDict, lineRepStmtDict);
        }

        private SyntaxNode ApplyReplace(SyntaxNode docRoot,
            Dictionary<string, (string, int)> PropNamePropStmtLineDict,
            Dictionary<int, string> lineRepStmtDict) {

            var sb = new StringBuilder();
            var lines = docRoot.GetText().Lines;
            // 書き換え
            for (int i = 0; i < lines.Count() - 1; i++) { 
                if (lineRepStmtDict.ContainsKey(i)) {
                    sb.AppendLine(lineRepStmtDict[i]);
                } else {
                    sb.AppendLine(lines[i].ToString());
                }
            }
            // VB.net形式のprop文をコードの最後辺りに追加する
            // (元のコードと書き換え後でラインがずれないようにするため)
            var count = lines.Count() - 1;
            foreach (var item in PropNamePropStmtLineDict) {
                var propStmt = item.Value.Item1;
                sb.AppendLine(propStmt);
                lineMappingDict[count] = item.Value.Item2;
                count++;
            }
            sb.AppendLine(lines.Last().ToString());
            return docRoot.SyntaxTree.
                WithChangedText(SourceText.From(sb.ToString())).GetRootAsync().Result;
        }
    }
}
