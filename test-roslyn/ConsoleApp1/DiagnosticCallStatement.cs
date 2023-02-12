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
    public class DiagnosticCallStatement {
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
                        if (!node.Parent.IsKind(SyntaxKind.CallStatement)) {
                            var loc = node.GetLocation();
                            var span = loc?.SourceSpan;
                            var tree = loc?.SourceTree;
                            if (span != null && tree != null) {
                                var start = tree.GetLineSpan(span.Value).StartLinePosition;
                                //Console.WriteLine($"line={start.Line}, ch={start.Character}{node.ToFullString()}");
                                locations.Add(new Location(span.Value.Start, start.Line, start.Character));
                            }
                        }
                    }
                }
            }
            return locations;
        }
    }
}
