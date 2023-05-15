using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConsoleApp1 {
    public class RewriteProperty {
        private List<TextChange> allChanges;
        private SortedDictionary<int, string> repDict;
        private HashSet<string> propNameSet;
        private SortedDictionary<int, string> repLetDict;
        private SortedDictionary<int, (string, string)> repGetDict;
        private SortedDictionary<int, string> repOtherDict;

        private void Init() {
            allChanges = new List<TextChange>();
            repDict = new SortedDictionary<int, string>();
            propNameSet = new HashSet<string>();
            repLetDict = new SortedDictionary<int, string>();
            repGetDict = new SortedDictionary<int, (string, string)>();
            repOtherDict = new SortedDictionary<int, string>();
        }

        public SyntaxNode Rewrite(SyntaxNode docRoot) {
            Init();

            var node = docRoot.DescendantNodes();

            EndPropertyStatement(node);

            var lines = docRoot.GetText().Lines;
            var forStmt = node.OfType<PropertyStatementSyntax>();

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
                var lineSpan = propToken.GetLocation().GetLineSpan();
                var propLineNum = lineSpan.StartLinePosition.Line;
                var propChildNodes = stmt.ChildNodes();

                GetPropertyStatement(stmt, propName, lines, propLineNum);
                LetPropertyStatement(stmt, propName, lines, propLineNum);
            }
            if (!repDict.Any()) {
                return docRoot;
            }

            if (allChanges.Any()) {
                var changes = docRoot.GetText().WithChanges(allChanges);
                docRoot = docRoot.SyntaxTree.WithChangedText(changes).GetRootAsync().Result;
            }

            UpdateGetProp(docRoot);
            var otherChanges = UpdateLetProp(docRoot);

            if (otherChanges.Any()) {
                var cchtext2 = docRoot.GetText().WithChanges(otherChanges);
                docRoot = docRoot.SyntaxTree.WithChangedText(cchtext2)
                    .GetRootAsync().Result;
            }

            docRoot = ApplyReplace(docRoot);
            return docRoot;
        }

        private void GetPropertyStatement(PropertyStatementSyntax stmt, string propName, TextLineCollection lines, int propLineNum) {
            var asTokens = stmt.ChildNodes().Where(x => x.IsKind(SyntaxKind.SimpleAsClause));
            if (!asTokens.Any()) {
                return;
            }

            var prePropLineNum = propLineNum - 1;
            var prePropSpan = lines[prePropLineNum].Span;
            if (prePropSpan.Length > 0) {
                return;
            }

            var asdef = asTokens.First().ToString();
            if (!propNameSet.Contains(propName)) {
                var insText = $"Public Property {propName} {asdef}";
                repDict.Add(prePropLineNum, insText);
            }
            propNameSet.Add(propName);

            var propSpan = lines[propLineNum].Span;
            var repCode = $"Private Function Get{propName}() {asdef}";
            allChanges.Add(new TextChange(
                new TextSpan(propSpan.Start, repCode.Length), repCode));
            repGetDict.Add(propLineNum, (propName, $"Get{propName}"));
        }

        private void LetPropertyStatement(PropertyStatementSyntax stmt, string propName, TextLineCollection lines, int propLineNum) {
            var prePropLineNum = propLineNum - 1;
            var prePropSpan = lines[prePropLineNum].Span;

            if (prePropSpan.Length > 0) {
                return;
            }

            var propArgs = stmt.ChildNodes().Where(x => x.IsKind(SyntaxKind.ParameterList));
            if (!propArgs.Any()) {
                return;
            }

            var paramsSym = propArgs.First() as ParameterListSyntax;
            var paramsAs = paramsSym.Parameters;
            if (!paramsAs.Any()) {
                return;
            }
            var asClause = paramsAs.First().ChildNodes().Where(
                x => x.IsKind(SyntaxKind.SimpleAsClause)).FirstOrDefault();
            if (asClause.IsKind(SyntaxKind.None)) {
                return;
            }
            var asClauseType = (asClause as SimpleAsClauseSyntax).Type;
            if (!propNameSet.Contains(propName)) {
                var insText = $"Public Property {propName} As {asClauseType}";
                repDict.Add(prePropLineNum, insText);
            }
            propNameSet.Add(propName);

            var propSpan = lines[propLineNum].Span;
            var chText = $"Private Function Let{propName}";
            allChanges.Add(new TextChange(
                new TextSpan(propSpan.Start, chText.Length), chText));
            repLetDict.Add(propLineNum, "let");
        }

        private void EndPropertyStatement(IEnumerable<SyntaxNode> node) {
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
        }

        private void UpdateGetProp(SyntaxNode docRoot) {
            //var repDict = new SortedDictionary<int, string>();
            var lines = docRoot.GetText().Lines;
            foreach (var lineNum in repGetDict.Keys) {
                var targetNode = docRoot.FindNode(lines[lineNum].Span);
                if (targetNode == null) {
                    continue;
                }

                var (org, pre) = repGetDict[lineNum];
                var assigs = targetNode.Parent.ChildNodes()
                    .Where(x => x.IsKind(SyntaxKind.SimpleAssignmentStatement))
                    .Cast<AssignmentStatementSyntax>();
                foreach (var assig in assigs) {
                    if (assig.Left.ToString() != org) {
                        continue;
                    }
                    var assgiLineNum = assig.GetLocation().GetLineSpan().StartLinePosition.Line;
                    var assigTrivia = assig.GetLeadingTrivia().FirstOrDefault().ToString();
                    repOtherDict.Add(assgiLineNum, $"{assigTrivia}{pre} = {assig.Right}");
                }
            }
        }

        private List<TextChange> UpdateLetProp(SyntaxNode docRoot) {
            var changes = new List<TextChange>();
            var lines = docRoot.GetText().Lines;
            foreach (var line in repLetDict.Keys) {
                var targetNode = docRoot.FindNode(lines[line].Span);
                if (targetNode == null) {
                    continue;
                }
                var methodNodes = targetNode.Parent.ChildNodes();
                if (methodNodes.Count() < 2) {
                    continue;
                }
                var methodStartEndNodes = new SyntaxNode[] { methodNodes.First(), methodNodes.Last() };
                foreach (var methodNode in methodStartEndNodes) {
                    var funcToken = methodNode.ChildTokens()
                        .Where(x => x.IsKind(SyntaxKind.FunctionKeyword)).FirstOrDefault();
                    if (funcToken.IsKind(SyntaxKind.None)) {
                        continue;
                    }
                    changes.Add(new TextChange(
                        new TextSpan(funcToken.Span.Start, funcToken.Span.Length), "Sub     "));
                }
            }
            return changes;
        }

        private SyntaxNode ApplyReplace(SyntaxNode docRoot) {
            var sb = new StringBuilder();
            var lines = docRoot.GetText().Lines;
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
            return docRoot.SyntaxTree.
                WithChangedText(SourceText.From(sb.ToString())).GetRootAsync().Result;
        }
    }
}
