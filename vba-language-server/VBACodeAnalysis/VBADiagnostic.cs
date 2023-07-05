using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace VBACodeAnalysis {
	public class VBADiagnostic {
        private RewriteSetting rewriteSetting;

        public VBADiagnostic(RewriteSetting setting) {
            this.rewriteSetting = setting;
        }

        public void SetSetting(RewriteSetting setting) {
            this.rewriteSetting = setting;
        }

        public async Task<List<DiagnosticItem>> GetDiagnostics(Document doc) {
            var codes = new string[] {
                "BC35000",  // ランタイム ライブラリ関数 が定義されていないため、
                                   // 要求された操作を実行できません。
                "BC30627", // 'Option' ステートメントは、宣言または 'Imports' ステートメントの前に記述しなければなりません
                //"BC30431", //  'End Property' の前には、対応する 'Property' を指定しなければなりません
                //"BC36759", // 自動実装プロパティはパラメーターを持つことができません
            };
            var AddItems = new List<DiagnosticItem>();
            var node = doc.GetSyntaxRootAsync().Result;
            var result = await doc.GetSemanticModelAsync();
            var diagnostics = result.GetDiagnostics();
            var items1 = diagnostics.Where(x => {
                if (codes.Contains(x.Id)) {
                    return false;
                }
                if(IsRewriteFunction(x, node)) {
                    return false;
				}

                if (x.Id == "BC30451") {
                    //  BC30451 宣言されていません。アクセスできない保護レベルになっています
                    if (IsFileInOutStatement1(x, node, ref AddItems)) {
                        return false;
                    }
                }
                if (x.Id == "BC30800" || x.Id == "BC32017") {
                    // BC30800 メソッドの引数はかっこで囲む必要があります。
                    // BC32017 コンマ、')'、または有効な式の継続文字が必要です。
                    if (IsFileInOutStatement2(x, node, ref AddItems)) {
                        return false;
                    }
                }
                if (x.Id == "BC30201") {
                    // BC30201 #1 式が必要です
                    if (IsFileInOutStatement3(x, node, ref AddItems)) {
                        return false;
                    }
                }

				if (x.Id == "BC30800") {
                    // メソッドの引数は、かっこで囲む必要があります
                    var targetNode = node.FindNode(x.Location.SourceSpan);
                    if (targetNode.Parent is ArgumentListSyntax arg) {
                        var op = arg.OpenParenToken;
                        var cp = arg.CloseParenToken;
                        if (op.IsMissing && cp.IsMissing) {
                            var token = targetNode.GetFirstToken();
                            var preToken = token.GetPreviousToken();
                            if (!preToken.IsKind(SyntaxKind.CallKeyword)) {
                                // testArgs 123
                                // Add testArgs(1,2)
                                // Add testArgs(123)
                                // Eq testArgs(1,2) = eqret
                                // Eq testArgs(123) = 1
                                // Eq testArgs(1,2) = 1
                                return false;
                            }
                        }
					} else {
                        var token = targetNode.GetFirstToken();
                        var preToken = token.GetPreviousToken();
                        if (!preToken.IsKind(SyntaxKind.CallKeyword)) {
                            // testArgs 1, 2
                            // Add testArgs 123
                            // Add testArgs 1,2
                            return false;
                        }
                    }
                }

                return true;
            });
            var items = items1.Select(x => {
                // Hidden = 0,
                // Info = 1,
                // Warning = 2,
                // Error = 3
                var severity = x.Severity.ToString();
                var msg = x.GetMessage();
                var s = x.Location.GetLineSpan().StartLinePosition;
                var e = x.Location.GetLineSpan().EndLinePosition;
                return new DiagnosticItem(severity, msg,
                    s.Line, s.Character,
                    e.Line, e.Character);
            }).ToList();

            AddMultiArgMethodDiag(node, ref AddItems);

            items.AddRange(AddItems);
            return items;
        }

        private bool IsRewriteFunction(Diagnostic x, SyntaxNode node) {
            if (x.Id != "BC30068") {
                return false;
            }
            var targetNode = node.FindNode(x.Location.SourceSpan);
			if (targetNode == null) {
                return false;
			}
            var tokens = targetNode.DescendantTokens().ToList();
            if (tokens.Count < 4) {
                return false;
            }
            var ns = rewriteSetting.NameSpace;
            if (tokens.First().ToString() != ns) {
                return false;
            }
            var dict = rewriteSetting.getRestoreDict();
            if (dict.ContainsKey(tokens[2].ToString()) && tokens[3].ToString() == "(") {
                return true;
            }
            return false;
        }

		private void AddMultiArgMethodDiag(SyntaxNode node, ref List<DiagnosticItem> dls) {
            // 引数が複数でCallがないsub, function呼び出しをエラーにする
            var forStmt = node.DescendantNodes().OfType<InvocationExpressionSyntax>();
            foreach (var stmt in forStmt) {
                if(stmt.ArgumentList == null) {
                    continue;
				}
                if (stmt.ArgumentList.Arguments.Count <= 1) {
					continue;
				}
                var isCallStmt = stmt.Parent.IsKind(SyntaxKind.CallStatement);
                var isExpressStmt = stmt.Parent.IsKind(SyntaxKind.ExpressionStatement);
                var isAssignStmt = stmt.Parent.IsKind(SyntaxKind.SimpleAssignmentStatement);
                if (!isCallStmt && !isAssignStmt && isExpressStmt) {
                    var opToken = stmt.ArgumentList.OpenParenToken;
                    if (opToken.Text.Length == 0) {
                        continue;
                    }
                    var lineSpan = stmt.GetLocation().GetLineSpan();
                    var startPos = lineSpan.StartLinePosition;
                    var endPos = lineSpan.EndLinePosition;
                    dls.Add(new DiagnosticItem(
                        DiagnosticSeverity.Error.ToString(),
                        "Call is required.",
                        startPos.Line, startPos.Character,
                        endPos.Line, endPos.Character));
                }
            }
		}

        private bool IsFileInOutStatement1(Diagnostic x, SyntaxNode node, ref List<DiagnosticItem> dls) {
            //  BC30451 宣言されていません。アクセスできない保護レベルになっています
            var idefNode = node.FindNode(x.Location.SourceSpan);
            var name = idefNode.ToString().ToLower();
            var targets = new string[] { "open", "close", "print", "write" };
            if (!targets.Contains(name)) {
                return false;
            }
            var parentNode = idefNode.Parent;
            var childNodes = parentNode?.ChildNodes();
            if (childNodes == null) {
                return false;
            }
            if (childNodes.Count() < 2) {
                return false;
            }
            var agsNode = childNodes.ElementAt(1);
            var arg = agsNode as ArgumentListSyntax;
            if (arg is null) {
                return false;
            }
            if (name == "open") {
                // Open fname For Output As #1
                var result = Regex.IsMatch(arg.ToFullString(),
                    @"\S+\s+For\s+(Input|Output|Append|Random|Binary)\s+As\s+(#\d+|\S+)",
                    RegexOptions.IgnoreCase);
                if (result) {
                    return true;
                }
                var lsp = parentNode.GetLocation().GetLineSpan();
                var sp = lsp.StartLinePosition;
                var ep = lsp.EndLinePosition;
                dls.Add(new DiagnosticItem(
                    DiagnosticSeverity.Error.ToString(),
                    "Requires 4 arguments. Open {FilePath} For Input | Output | Append As #Filenumber",
                    sp.Line, sp.Character,
                    ep.Line, ep.Character));
            }
            return false;
        }

        private bool IsFileInOutStatement2(Diagnostic x, SyntaxNode node, ref List<DiagnosticItem> dls) {
            // BC30800 メソッドの引数はかっこで囲む必要があります。
            // BC32017 コンマ、')'、または有効な式の継続文字が必要です。
            var openParent = node.FindNode(x.Location.SourceSpan).Parent;
            var openChild = openParent.ChildNodes();
            if (!openChild.Any()) {
                return false;
            }
            var name = openChild.First().ToString().ToLower();
            var targets = new string[] { "open" };
            if (targets.Contains(name)) {
                return true;
            }
            return false;
        }
        private bool IsFileInOutStatement3(Diagnostic x, SyntaxNode node, ref List<DiagnosticItem> dls) {
            // BC30201 #1 式が必要です
            var fileNum = node.FindNode(x.Location.SourceSpan);
            var result = Regex.IsMatch(fileNum.ToFullString(),
                @"#\d+", RegexOptions.IgnoreCase);
			if (!result) {
                return false;
			}
            var parent = fileNum;
			for (int i = 0; i < 2; i++) {
                if(parent == null) {
                    return false;
                }
                if (parent.IsKind(SyntaxKind.ArgumentList)) {
                    parent = parent.Parent;
                    break;
                }
                parent = parent.Parent;
            }
            var names = new List<string> { "close", "print", "write" };
            var childnodes = parent?.ChildNodes();
            if(childnodes != null && childnodes.Any()) {
                var name = childnodes.First().ToString().ToLower();
                return names.Contains(name);
            }
            return false;
        }
    }
}
