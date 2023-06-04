using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBACodeAnalysis {
    public class Rewrite {
        private RewriteSetting setting;
        public Dictionary<int, (int, int)> charaOffsetDict;
        public Dictionary<int, int> lineMappingDict;

        public Rewrite(RewriteSetting setting) {
            this.setting = setting;
        }

        public SourceText RewriteStatement(Document doc) {
            var docRoot = doc.GetSyntaxRootAsync().Result;
            var nodes = docRoot.DescendantNodes();
            var changes = new List<TextChange>();

            var rewriteProp = new RewriteProperty();
            docRoot = rewriteProp.Rewrite(docRoot);
            charaOffsetDict = rewriteProp.charaOffsetDict;
            lineMappingDict = rewriteProp.lineMappingDict;

            nodes = docRoot.DescendantNodes();

            changes = changes.Concat(ReplaceStatement(nodes)).ToList();
            changes = changes.Concat(SetStatement(nodes)).ToList();
            changes = changes.Concat(LocalDeclarationStatement(nodes)).ToList();
            changes = changes.Concat(FieldDeclarationStatement(nodes)).ToList();

            return docRoot.GetText().WithChanges(changes);
        }

		public List<TextChange> ReplaceStatement(IEnumerable<SyntaxNode> node) {
            var ns = setting.NameSpace;
            var rewriteDict = setting.getRewriteDict();
            var allChanges = new List<TextChange>();
            //var docRoot = doc.GetSyntaxRootAsync().Result;
            var forStmt = node.OfType<InvocationExpressionSyntax>();
            var set = new HashSet<string>();
            foreach (var stmt in forStmt) {
                var tt = stmt.GetFirstToken().Text;
                if (rewriteDict.ContainsKey(tt) && stmt.ArgumentList != null) {
                    var sp = stmt.GetFirstToken().Span;
                    var k = $"{sp.Start}-{sp.End}";
                    if (!set.Contains(k)) {
                        var rename = $"{ns}.{rewriteDict[tt]}";
                        allChanges.Add(new TextChange(stmt.GetFirstToken().Span, rename));
                    }
                    set.Add(k);
                }
            }
            return allChanges;
        }

        private List<TextChange> SetStatement(IEnumerable<SyntaxNode> node) {
            var forStmt = node.OfType<EmptyStatementSyntax>();
            var allChanges = new List<TextChange>();
            // 'Let' および 'Set' 代入ステートメントはサポートされなくなりました。
            const string code = "BC30807";
            foreach (var stmt in forStmt) {
                var ds = stmt.Empty.TrailingTrivia.Where(x => {
                    return x.GetDiagnostics().SingleOrDefault(x => x.Id == code) != null;
                });
                var changes = ds.Select(x => {
                    return new TextChange(x.Span, new string(' ', x.Span.Length));
                });
                if (changes.Count() > 0) {
                    allChanges = allChanges.Concat(changes).ToList();
                }
            }
            return allChanges;
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
    }
}
