using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1 {
    public class DiagnosticRewrite {

        public List<TextChange> RewriteStatement(IEnumerable<SyntaxNode> node) {
            var changes = new List<TextChange>();
            changes = changes.Concat(RewriteSetStatement(node)).ToList();
            changes = changes.Concat(LocalDeclarationStatement(node)).ToList();
            changes = changes.Concat(FieldDeclarationStatement(node)).ToList();
            return changes;
        }

        private List<TextChange> RewriteSetStatement(IEnumerable<SyntaxNode> node) {
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
