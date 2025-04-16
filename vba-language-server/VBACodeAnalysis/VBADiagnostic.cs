using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace VBACodeAnalysis {
	public class VBADiagnostic {
        public List<IgnoreDiagnostic> ignoreDs;
        private Dictionary<string, string> _errorMsgDict;
        public VBADiagnostic() {
            ignoreDs = [];
			_errorMsgDict = new Dictionary<string, string> {
				["open"] = Properties.Resources.OpenErrorMsg,
				["print"] = Properties.Resources.PrintErrorMsg,
				["write"] = Properties.Resources.WriteErrorMsg,
				["close"] = Properties.Resources.CloseErrorMsg,
				["input"] = Properties.Resources.InputErrorMsg,
				["line_input"] = Properties.Resources.LineInputErrorMsg,
			};
		}

		public async Task<List<DiagnosticItem>> GetDiagnostics(Document doc) {
            var codes = new string[] {
                "BC35000",  // ランタイム ライブラリ関数 が定義されていないため、
                                   // 要求された操作を実行できません。
            };
            var AddItems = new List<DiagnosticItem>();
            var node = doc.GetSyntaxRootAsync().Result;
            var result = await doc.GetSemanticModelAsync();
            var diagnostics = result.GetDiagnostics();
            var items1 = diagnostics.Where(x => {
                if (codes.Contains(x.Id)) {
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
                return new DiagnosticItem(x.Id, severity, msg,
                    s.Line, s.Character,
                    e.Line, e.Character);
            }).ToList();

            AddMultiArgMethodDiag(node, ref AddItems);

            items.AddRange(AddItems);
            return items;
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
                        "VBA_call",
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
            var ss = x.Location.GetLineSpan();
            var ssp = ss.StartLinePosition;
            var sep = ss.EndLinePosition;

			foreach (var item in ignoreDs) {
                if(item.Text == idefNode.ToString()
					&& ssp.Line == item.StartLine
                    && ssp.Character == item.StartCol
					&& sep.Line == item.EndLine
					&& sep.Character == item.EndCol) {
					return true;
				}
			}
            var name = idefNode.ToString();
			var targets = new List<string> { "open", "close", "print", "write", "input", "line_input" };
            if (Util.Contains(name, targets)) {
                return true;
            }
            return false;
        }

        private bool IsFileInOutStatement2(Diagnostic x, SyntaxNode node, ref List<DiagnosticItem> dls) {
            // BC30800 メソッドの引数はかっこで囲む必要があります。
            // BC32017 コンマ、')'、または有効な式の継続文字が必要です。
			var op = node.FindNode(x.Location.SourceSpan).GetFirstToken().GetPreviousToken();
			var ss = op.GetLocation().GetLineSpan();
			var ssp = ss.StartLinePosition;
			var sep = ss.EndLinePosition;
			foreach (var item in ignoreDs) {
				if (Util.Eq(item.Text, op.Text)
					&& ssp.Line == item.StartLine
					&& ssp.Character == item.StartCol
					&& sep.Line == item.EndLine
					&& sep.Character == item.EndCol) {
					return true;
				}
			}

			var name = op.Text.ToLower();
            if (_errorMsgDict.ContainsKey(name) && op.GetPreviousToken().Text == ".") {
				// open, print, write,...がエラーで何かのメソッドの場合は
				// Properties.Resourcesのエラーではないので抜ける
				return false;
            }
			if (_errorMsgDict.TryGetValue(name, out string value)) {
				var msg = value;
                var id = $"VBA_{name}";
				var endCol = sep.Character;
                var diagno = new DiagnosticItem(
					id,
                    DiagnosticSeverity.Error.ToString(),
                    msg,
                    ssp.Line, ssp.Character,
                    sep.Line, endCol);
				if (!Contains(dls, diagno)) {
					dls.Add(diagno);
				}
				return true;
			}
            return false;
        }

        private bool Contains(List<DiagnosticItem> dls, DiagnosticItem d) {
            foreach (var item in dls) {
                if (item.Eq(d)) {
                    return true;
                }
			}
            return false;
        }
    }
}
