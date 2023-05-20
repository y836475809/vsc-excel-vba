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
        // line - (chara start, offset)
        public Dictionary<int, (int, int)> charaOffsetDict;

		// line - line num
		public Dictionary<int, int> lineMappingDict;

		private void Init() {
            charaOffsetDict = new Dictionary<int, (int, int)>();
            lineMappingDict = new Dictionary<int, int>();
        }

        public SyntaxNode Rewrite(SyntaxNode docRoot) {
            Init();

            var props = docRoot.DescendantTokens()
                .Where(x => x.IsKind(SyntaxKind.PropertyKeyword))
                .Select(x => x.Parent);

           var prppreChanges = new List<TextChange>();
            var pairs = new List<(PropertyStatementSyntax, EndBlockStatementSyntax)>();
			for (int i = 0; i < props.Count()/2; i++) {
                pairs.Add((
                    props.ElementAt(i*2) as PropertyStatementSyntax, 
                    props.ElementAt(i*2 + 1) as EndBlockStatementSyntax));
            }
            var map = new Dictionary<SyntaxToken, SyntaxToken>();
            
            var replinemap = new Dictionary<int, string>();
            foreach (var pair in pairs) {
                if(pair.Item1 == null || pair.Item2 == null) {
                    continue;
				}
                var propsetletkeys = pair.Item1.ChildTokens()
                .Where(x =>
                    x.ToString().ToLower() == "get"
                    || x.ToString().ToLower() == "let"
                    || x.ToString().ToLower() == "set");
                if (propsetletkeys.Any()) {
                    var getToken = propsetletkeys.First();
                    var attname = getToken.ToString().ToLower();
                    replinemap.Add(
                        getToken.GetLocation().GetLineSpan().StartLinePosition.Line,
                        attname.ToLower());
                    prppreChanges.Add(
                        new TextChange(getToken.Span, new string(' ', getToken.Span.Length))); ;

                    var propkeys = pair.Item1.ChildTokens()
                        .Where(x => x.IsKind(SyntaxKind.PropertyKeyword));
                    if (propkeys.Any()) {
                        var pToken = propkeys.First();
                        if (getToken.ToString().ToLower() == "get") {
                            prppreChanges.Add(new TextChange(pToken.Span, "Function"));
                            var endpropkeys = pair.Item2.ChildTokens()
                                .Where(x => x.IsKind(SyntaxKind.PropertyKeyword));
                            if (endpropkeys.Any()) {
                                prppreChanges.Add(new TextChange(endpropkeys.First().Span, "Function"));
                            }
                        } else {
                            prppreChanges.Add(new TextChange(pToken.Span, "Sub     "));
                            var endpropkeys = pair.Item2.ChildTokens()
                                .Where(x => x.IsKind(SyntaxKind.PropertyKeyword));
                            if (endpropkeys.Any()) {
                                prppreChanges.Add(new TextChange(endpropkeys.First().Span, "Sub     "));
                            }
                        }
                    }
                }
            }

            docRoot = docRoot.SyntaxTree.WithChangedText(
                docRoot.GetText().WithChanges(prppreChanges))
                .GetRootAsync().Result;

            var newPropLineDict = new Dictionary<string, int>();
            var newPropStateDict = new Dictionary<string, string>();
            var repLineMap = new Dictionary<int, string>();
            var lines = docRoot.GetText().Lines;
            foreach (var linenum in replinemap.Keys) {
                var line = lines[linenum];
                var propnode = docRoot.FindNode(line.Span);
				if (!(propnode is MethodStatementSyntax mathodnode)) {
					continue;
				}
				var text = line.ToString();
                var pp = replinemap[linenum];
                var funcname = mathodnode.Identifier.ToString();
				if (mathodnode.IsKind(SyntaxKind.FunctionStatement)){
                    newPropStateDict[funcname] =
                        $"Public Property {funcname} As {mathodnode.AsClause.Type}";
                }
                if (mathodnode.IsKind(SyntaxKind.SubStatement)) {
                    var paramsAs = mathodnode.ParameterList.Parameters;
                    if (paramsAs.Any()) {
                        var asClause = paramsAs.First().ChildNodes().Where(
                            x => x.IsKind(SyntaxKind.SimpleAsClause)).FirstOrDefault();
                        if (!asClause.IsKind(SyntaxKind.None)) {
                            var asClauseType = (asClause as SimpleAsClauseSyntax).Type;
                            newPropStateDict[funcname] =
                                $"Public Property {funcname} As {asClauseType}";
                        }
                    }
                } 

                if (!newPropLineDict.ContainsKey(funcname)) {
                    newPropLineDict[funcname] = linenum;
                }

                var spinde = text.IndexOf(funcname);
                if (spinde > 0) {
                    if((spinde - pp.Length) >= 0) {
                        text = text.Remove(spinde - pp.Length, pp.Length);
                    }
                }
                text = text.Replace(funcname, $"{pp}{funcname}");
                if (mathodnode.Modifiers.Any()) {
                    var mod = mathodnode.Modifiers.First();
					if (mod.IsKind(SyntaxKind.PublicKeyword)) {
                        text = text.Replace("Public", "Private");
                    }
                    var m = mod.Text;
                    var d = m.Length - "Private".Length;
                    charaOffsetDict[linenum] = ("Private".Length, d);
                } else {
                    text = $"Private {text}";
                    var d = "Private".Length;
                    charaOffsetDict[linenum] = (d, d);
                }
                repLineMap.Add(linenum, text);

                var propChNodes = propnode.Parent.ChildNodes();
                var assigs = propChNodes
                    .Where(x => x.IsKind(SyntaxKind.SimpleAssignmentStatement))
                    .Cast<AssignmentStatementSyntax>();
				if (assigs.Any()) {
					foreach (var item in assigs) {
                        if (item.Left.ToString() != funcname) {
                            continue;
                        }
                        var asLine = item.GetLocation().GetLineSpan().StartLinePosition.Line;
                        var asline = lines[asLine].ToString();
                        var repval = $"{pp}{funcname}";
                        repLineMap.Add(asLine, asline.Replace(funcname, repval));

                        var assigTrivia = item.GetLeadingTrivia().FirstOrDefault().ToString();
                        charaOffsetDict[asLine] = (assigTrivia.Length + repval.Length, funcname.Length - repval.Length);
                    }
				}
            }

            docRoot = ApplyReplace3(docRoot, newPropLineDict, newPropStateDict, repLineMap);
            return docRoot;
        }

        private SyntaxNode ApplyReplace3(SyntaxNode docRoot,
            Dictionary<string, int> newPropLineDict,
            Dictionary<string, string> newPropStateDict,
            Dictionary<int, string> repdict) {
           var sb = new StringBuilder();
            var lines = docRoot.GetText().Lines;
            for (int i = 0; i < lines.Count() - 1; i++) { 
                if (repdict.ContainsKey(i)) {
                    sb.AppendLine(repdict[i]);
                } else {
                    sb.AppendLine(lines[i].ToString());
                }
            }

            var count = lines.Count() - 1;
            foreach (var item in newPropStateDict) {
                sb.AppendLine(item.Value);
                if (newPropLineDict.ContainsKey(item.Key)) {
                    lineMappingDict[count] = newPropLineDict[item.Key];
                }
                count++;
            }
            sb.AppendLine(lines.Last().ToString());
            return docRoot.SyntaxTree.
                WithChangedText(SourceText.From(sb.ToString())).GetRootAsync().Result;
        }
    }
}
