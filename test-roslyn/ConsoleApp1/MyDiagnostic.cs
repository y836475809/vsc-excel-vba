using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1 {
	class MyDiagnostic {
        private RewriteSetting rewriteSetting;

        public MyDiagnostic(RewriteSetting setting) {
            this.rewriteSetting = setting;
        }

        public void SetSetting(RewriteSetting setting) {
            this.rewriteSetting = setting;
        }

        public async Task<List<DiagnosticItem>> GetDiagnostics(Document doc) {
            var sourceText = doc.GetTextAsync().Result;
            var codes = new string[] {
                "BC35000",  // ランタイム ライブラリ関数 が定義されていないため、
                                   // 要求された操作を実行できません。
                "BC30627", // 'Option' ステートメントは、宣言または 'Imports' ステートメントの前に記述しなければなりません
            };
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
                if (x.Id == "BC30800") {
                    var sym = SymbolFinder.FindSymbolAtPositionAsync(
                        doc, x.Location.SourceSpan.Start - 1).Result;
                    if (sym is IMethodSymbol mth) {
                        if (mth.ReturnsVoid) {
                            var token = node.FindNode(x.Location.SourceSpan).GetFirstToken();
                            var preToken = token.GetPreviousToken();
                            if (!preToken.IsKind(SyntaxKind.CallKeyword)) {
                                var neToken = token.GetNextToken();
                                if (!neToken.IsKind(SyntaxKind.OpenParenToken)) {
                                    // method "test"の場合(Callなしカッコなしの呼び出し)のエラーはエラーにしない
                                    return false;
                                }
                            }
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

            var diagnosticCall = await getCallStatementAsync(doc);
            items.AddRange(diagnosticCall);
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

        private async Task<List<DiagnosticItem>> getCallStatementAsync(Document document) {
            var d = new DiagnosticCallStatement();
            var locations = await d.mmAsync(document);
            return locations.Select(x => {
                var severity = DiagnosticSeverity.Error.ToString();
                var msg = "Call is required";
                var positon = x.Positon;
                return new DiagnosticItem(severity, msg,
                    x.Line, x.Character,
                    x.Line, x.Character);
            }).ToList();
        }

        public async Task<List<Location>> mmAsync(Document document) {
            var locations = new List<Location>();
            var workspace = document.Project.Solution.Workspace;
            var model = await document.GetSemanticModelAsync();

            var syntaxRoot = await document.GetSyntaxRootAsync();
            var forStmt = syntaxRoot.DescendantNodes().OfType<InvocationExpressionSyntax>();
            foreach (var stmt in forStmt) {
                var node = stmt.ChildNodes().First();
                var position = (int)(node.Span.Start + node.Span.End) / 2;
                var symbol = await SymbolFinder.FindSymbolAtPositionAsync(
                    model, position, workspace);
                if (symbol == null) {
                    continue;
                }
                if (symbol is IMethodSymbol mth) {
                    if (mth.ReturnsVoid) {
                        continue;
                    }
                    if (node.Parent.IsKind(SyntaxKind.CallStatement)) {
                        continue;
                    }
                    var loc = node.GetLocation();
                    var span = loc?.SourceSpan;
                    var tree = loc?.SourceTree;
                    if (span != null && tree != null) {
                        var start = tree.GetLineSpan(span.Value).StartLinePosition;
                        locations.Add(new Location(span.Value.Start, start.Line, start.Character));
                    }
                }
            }
            return locations;
        }
    }
}
