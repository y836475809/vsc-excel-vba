using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using Microsoft.VisualStudio.LanguageServer.Protocol;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using LSP = Microsoft.VisualStudio.LanguageServer.Protocol;

namespace VBACodeAnalysis {
	class DocumentSymbolProvider {
		public static DocumentSymbol[] GetDocumentSymbols(SyntaxNode node, Uri uri, Func<int, (bool, string, string)> propMapFunc) {
			var children = new List<DocumentSymbol>();
			var varChildren = GetFieldVarSymbols(node, LSP.SymbolKind.Field, true);
			var methodStmts = node.DescendantNodes().OfType<MethodBlockSyntax>();
			foreach (var stmt in methodStmts) {
				var s = stmt.SubOrFunctionStatement;
				var name = s.Identifier.Text;
				var lineSpan = s.GetLocation().GetLineSpan();
				var sp = lineSpan.StartLinePosition;
				var ep = lineSpan.EndLinePosition;
				var (ret, prefix, propName) = propMapFunc(sp.Line);
				if (ret) {
					var range = new LSP.Range {
						Start = new Position { Line = sp.Line, Character = 0 },
						End = new Position { Line = ep.Line, Character = ep.Character },
					};
					name = $"{prefix} {propName}";
					children.Add(new DocumentSymbol {
						Name = name.Trim(),
						Deprecated = false,
						Kind = LSP.SymbolKind.Property,
						Range = range,
						SelectionRange = range
					});
				} else {
					var range = new LSP.Range {
						Start = new Position { Line = sp.Line, Character = 0 },
						End = new Position { Line = ep.Line, Character = ep.Character },
					};
					var varChildren2 = GetLocalVarSymbols(stmt);
					children.Add(new DocumentSymbol {
						Name = name.Trim(),
						Deprecated = false,
						Kind = LSP.SymbolKind.Method,
						Range = range,
						SelectionRange = range,
						Children = [.. varChildren2]
					});
				}
			}

			var varStmts = node.DescendantNodes().OfType<DeclareStatementSyntax>();
			foreach (var stmt in varStmts) {
				var name = stmt.Identifier.Text;
				var lineSpan = stmt.GetLocation().GetLineSpan();
				var sp = lineSpan.StartLinePosition;
				var ep = lineSpan.EndLinePosition;
				var range = new LSP.Range {
					Start = new Position { Line = sp.Line, Character = 0 },
					End = new Position { Line = ep.Line, Character = ep.Character },
				};
				children.Add(new DocumentSymbol {
					Name = name.Trim(),
					Deprecated = false,
					Kind = LSP.SymbolKind.Variable,
					Range = range,
					SelectionRange = range
				});
			}

			var rootName = Path.GetFileNameWithoutExtension(uri.LocalPath);
			var ext = Path.GetExtension(uri.LocalPath);
			LSP.SymbolKind rootKind;
			if (ext == ".bas") {
				rootKind = LSP.SymbolKind.Module;
			} else if (ext == ".cls") {
				rootKind = LSP.SymbolKind.Class;
			} else {
				rootKind = LSP.SymbolKind.Module;
			}
			var rootLineSpen = node.GetLocation().GetLineSpan();
			var rootSp = rootLineSpen.StartLinePosition;
			var rootEp = rootLineSpen.EndLinePosition;
			var rootRange = new LSP.Range {
				Start = new Position { Line = rootSp.Line, Character = 0 },
				End = new Position { Line = rootEp.Line, Character = rootEp.Character },
			};
			var typeChildren = GetTypeSymbols(node);
			var enumChildren = GetEnumSymbols(node);
			var nchildren = children.Concat(typeChildren).Concat(enumChildren).Concat(varChildren);
			var docSymbol = new DocumentSymbol {
				Name = rootName.Trim(),
				Kind = rootKind,
				Deprecated = false,
				Range = rootRange,
				SelectionRange = rootRange,
				Children = [.. nchildren]
			};
			
			return [docSymbol];
		}

		private static List<DocumentSymbol> GetFieldVarSymbols(SyntaxNode node, LSP.SymbolKind kind, bool root) {
			var symbols = new List<DocumentSymbol>();
			var varSyntaxes = node.DescendantNodes().OfType<FieldDeclarationSyntax>();
			foreach (var syntax in varSyntaxes) {
				if (root) {
					if (syntax.Parent != null && syntax.Parent.IsKind(Microsoft.CodeAnalysis.VisualBasic.SyntaxKind.StructureBlock)) {
						continue;
					}
				} else {
					if (node.IsKind(Microsoft.CodeAnalysis.VisualBasic.SyntaxKind.ModuleBlock)) {
						continue;
					}
					if (node.IsKind(Microsoft.CodeAnalysis.VisualBasic.SyntaxKind.ClassBlock)) {
						continue;
					}
				}
				foreach (var declarator in syntax.Declarators) {
					var ident = declarator.Names.ToFullString();
					var spen = syntax.GetLocation().GetLineSpan();
					var sp = spen.StartLinePosition;
					var ep = spen.EndLinePosition;
					var range = new LSP.Range {
						Start = new Position { Line = sp.Line, Character = 0 },
						End = new Position { Line = ep.Line, Character = ep.Character },
					};
					symbols.Add(new DocumentSymbol {
						Name = ident.Trim(),
						Kind = kind,
						Deprecated = false,
						Range = range,
						SelectionRange = range
					});
				}
			}
			return symbols;
		}

		private static List<DocumentSymbol> GetTypeSymbols(SyntaxNode node) {
			var symbols = new List<DocumentSymbol>();
			var structureSyntaxes = node.DescendantNodes().OfType<StructureBlockSyntax>();
			foreach (var syntax in structureSyntaxes) {
				var stmt = syntax.StructureStatement;
				var ident = stmt.Identifier.Text;
				var spen = stmt.GetLocation().GetLineSpan();
				var sp = spen.StartLinePosition;
				var ep = spen.EndLinePosition;
				var range = new LSP.Range {
					Start = new Position { Line = sp.Line, Character = 0 },
					End = new Position { Line = ep.Line, Character = ep.Character },
				};
				var varChildren = GetFieldVarSymbols(stmt.Parent, LSP.SymbolKind.Variable, false);
				symbols.Add(new DocumentSymbol {
					Name = ident.Trim(),
					Kind = LSP.SymbolKind.Struct,
					Deprecated = false,
					Range = range,
					SelectionRange = range,
					Children = [..varChildren]
				});
			}
			return symbols;
		}

		private static List<DocumentSymbol> GetEnumSymbols(SyntaxNode node) {
			var symbols = new List<DocumentSymbol>();
			var enumSyntaxes = node.DescendantNodes().OfType<EnumBlockSyntax>();
			foreach (var syntax in enumSyntaxes) {
				var stmt = syntax.EnumStatement;
				var ident = stmt.Identifier.Text;
				var spen = stmt.GetLocation().GetLineSpan();
				var sp = spen.StartLinePosition;
				var ep = spen.EndLinePosition;
				var range = new LSP.Range {
					Start = new Position { Line = sp.Line, Character = 0 },
					End = new Position { Line = ep.Line, Character = ep.Character },
				};
				var varChildren = GetEnumVarSymbols(stmt.Parent);
				symbols.Add(new DocumentSymbol {
					Name = ident.Trim(),
					Kind = LSP.SymbolKind.Enum,
					Deprecated = false,
					Range = range,
					SelectionRange = range,
					Children = [.. varChildren]
				});
			}
			return symbols;
		}

		private static List<DocumentSymbol> GetLocalVarSymbols(SyntaxNode node) {
			var symbols = new List<DocumentSymbol>();
			var varSyntaxes = node.DescendantNodes().OfType<LocalDeclarationStatementSyntax>();
			foreach (var syntax in varSyntaxes) {
				foreach (var declarator in syntax.Declarators) {
					var ident = declarator.Names.ToFullString();
					var spen = syntax.GetLocation().GetLineSpan();
					var sp = spen.StartLinePosition;
					var ep = spen.EndLinePosition;
					var range = new LSP.Range {
						Start = new Position { Line = sp.Line, Character = 0 },
						End = new Position { Line = ep.Line, Character = ep.Character },
					};
					symbols.Add(new DocumentSymbol {
						Name = ident.Trim(),
						Kind = LSP.SymbolKind.Variable,
						Deprecated = false,
						Range = range,
						SelectionRange = range
					});
				}
			}
			return symbols;
		}

		private static List<DocumentSymbol> GetEnumVarSymbols(SyntaxNode node) {
			var symbols = new List<DocumentSymbol>();
			var varSyntaxes = node.DescendantNodes().OfType<EnumMemberDeclarationSyntax>();
			foreach (var syntax in varSyntaxes) {
				var ident = syntax.Identifier.Text;
				var spen = syntax.GetLocation().GetLineSpan();
				var sp = spen.StartLinePosition;
				var ep = spen.EndLinePosition;
				var range = new LSP.Range {
					Start = new Position { Line = sp.Line, Character = 0 },
					End = new Position { Line = ep.Line, Character = ep.Character },
				};
				symbols.Add(new DocumentSymbol {
					Name = ident.Trim(),
					Kind = LSP.SymbolKind.EnumMember,
					Deprecated = false,
					Range = range,
					SelectionRange = range
				});
			}
			return symbols;
		}
	}
}
