﻿using Microsoft.CodeAnalysis;
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
	public class MyDiagnostic {
        private RewriteSetting rewriteSetting;

        public MyDiagnostic(RewriteSetting setting) {
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

                if (IsOpen(x, node, ref AddItems)) {
                    return false;
                }
                if (IsClose(x, node, ref AddItems)) {
                    return false;
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

        private bool IsOpen(Diagnostic x, SyntaxNode node, ref  List<DiagnosticItem> dls) {
            // Open fname For Output As #1
            if (x.Id == "BC30451") {
                // Open 宣言されていません。アクセスできない保護レベルになっています
                var token = node.FindToken(x.Location.SourceSpan.Start);
                if (token.Text == "Open") {        
					try {
                        // fname
                        var neToken1 = node.FindToken(x.Location.SourceSpan.End + 1);
                        //  For Output As #1
                        var neToken2 = node.FindToken(neToken1.GetLocation().SourceSpan.End + 1);
                        if (neToken2.LeadingTrivia.Count == 0) {
                            var lineSpan1 = neToken1.GetLocation().GetLineSpan();
                            //var slp1 = lineSpan1.StartLinePosition;
							var elp1 = lineSpan1.EndLinePosition;
							dls.Add(new DiagnosticItem(
                                DiagnosticSeverity.Error.ToString(),
                                "Requires 4 arguments, Open {FilePath} For Input | Output | Append As #{Number}",
                                elp1.Line, elp1.Character+1,
                                elp1.Line, elp1.Character+1));
                        }
                        var argsText = neToken2.LeadingTrivia[0].ToFullString();
                        var args = argsText.Split(" ").Where(x => x.Length > 0).ToList();
                        var argLineSpan = token.GetLocation().GetLineSpan();
                        var argslp = argLineSpan.StartLinePosition;
                        var argelp = argLineSpan.EndLinePosition;
                        //Input Output Append
                        if (args.Count != 4) {

                            dls.Add(new DiagnosticItem(
                                DiagnosticSeverity.Error.ToString(),
                                "Requires 4 arguments, Open {FilePath} For Input | Output | Append As #{Number}",
                                argslp.Line, argslp.Character,
                                argelp.Line, argelp.Character));
                        }
                        if (args.Count == 4) {
                            if(args[0].ToLower() != "for") {
                                dls.Add(new DiagnosticItem(
                                     DiagnosticSeverity.Error.ToString(),
                                     "\"For\" is required, Open {FilePath} For Input | Output | Append As #{Number}",
                                     argslp.Line, argslp.Character,
                                     argelp.Line, argelp.Character));
                            }
                            var otypes = new string[] { "Input", "Output", "Append"};
                            if (!otypes.Contains(args[1])) {
                                dls.Add(new DiagnosticItem(
                                     DiagnosticSeverity.Error.ToString(),
                                     "\"Access\" is required, Input or Output or Append, Open {FilePath} For Input | Output | Append As #{Number}",
                                     argslp.Line, argslp.Character,
                                     argelp.Line, argelp.Character));
                            }
                            if (args[2].ToLower() != "as") {
                                dls.Add(new DiagnosticItem(
                                     DiagnosticSeverity.Error.ToString(),
                                     "\"As\" is required, Open {FilePath} For Input | Output | Append As #{Number}",
                                     argslp.Line, argslp.Character,
                                     argelp.Line, argelp.Character));
                            }
                        }
                    } catch (Exception ex) {
						//throw;
					}
                    return true;
                }
            }
            if (x.Id == "BC30800") {
                // メソッドの引数はかっこで囲む必要があります。
                var token = node.FindToken(x.Location.SourceSpan.Start);
                var pre = token.GetPreviousToken();
                if (pre.Text == "Open") {
                    return true;
                }
            }
            if (x.Id == "BC32017") {
                // コンマ、')'、または有効な式の継続文字が必要です。
                var token = node.FindToken(x.Location.SourceSpan.Start);
                var pre = token.GetPreviousToken().GetPreviousToken();
                if (pre.Text == "Open") {
                    return true;
                }
            }
            return false;
        }

        private bool IsClose(Diagnostic x, SyntaxNode node, ref List<DiagnosticItem> dls) {
            // Close #1
            if (x.Id == "BC30451") {
                // Open 宣言されていません。アクセスできない保護レベルになっています
                var token = node.FindToken(x.Location.SourceSpan.Start);
                if (token.Text == "Close") {
                    //Close #1    '1番のファイルを閉じます
                    //Close       '現在開いているすべてのファイルを閉じます
                    var neToken = node.FindToken(x.Location.SourceSpan.End + 1);
                    var neTrivia = neToken.TrailingTrivia;
                    if (neTrivia.Count == 2) {
                        var hasFileNum = Regex.IsMatch(neTrivia[0].ToFullString(), "#[0-9]+$");
                        if (!hasFileNum || !neTrivia[1].IsKind(SyntaxKind.EndOfLineTrivia)) {
                            var lineSpan = neToken.GetLocation().GetLineSpan();
                            var slp = lineSpan.StartLinePosition;
                            var elp = lineSpan.EndLinePosition;
                            dls.Add(new DiagnosticItem(
                             DiagnosticSeverity.Error.ToString(),
                             "Close #{Number}",
                             slp.Line, slp.Character,
                             elp.Line, elp.Character));
                        }
                    } else if (neTrivia.Count > 2) {
                        var lineSpan = neToken.GetLocation().GetLineSpan();
                        var slp = lineSpan.StartLinePosition;
                        var elp = lineSpan.EndLinePosition;
                        dls.Add(new DiagnosticItem(
                         DiagnosticSeverity.Error.ToString(),
                         "Close #{Number}",
                         slp.Line, slp.Character,
                         elp.Line, elp.Character));
                    }
                    return true;
                }
            }
            if (x.Id == "BC30800") {
                // メソッドの引数はかっこで囲む必要があります。
                var token = node.FindToken(x.Location.SourceSpan.Start);
                var pre = token.GetPreviousToken();
                if (pre.Text == "Close") {
                    return true;
                }
            }
            if (x.Id == "BC32017") {
                // コンマ、')'、または有効な式の継続文字が必要です。
                var token = node.FindToken(x.Location.SourceSpan.Start);
                var pre = token.GetPreviousToken().GetPreviousToken();
                if (pre.Text == "Close") {
                    return true;
                }
            }
            if (x.Id == "BC30201") {
                // #1 式が必要です
                var token = node.FindToken(x.Location.SourceSpan.Start);
                if (token.GetPreviousToken().Text == "Close") {
                    return true;
                }
            }
            return false;
        }
    }
}
