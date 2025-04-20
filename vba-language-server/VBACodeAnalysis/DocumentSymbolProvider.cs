using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Elfie.Model;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using Microsoft.VisualStudio.LanguageServer.Protocol;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using LSP = Microsoft.VisualStudio.LanguageServer.Protocol;
using CAVB = Microsoft.CodeAnalysis.VisualBasic;

namespace VBACodeAnalysis {
	class DocumentSymbolProvider {
		public static DocumentSymbol[] GetDocumentSymbols(SyntaxNode node, Uri uri, Func<int, (bool, string, string)> propMapFunc) {
			var children = new List<DocumentSymbol>();
			children.AddRange(GetVarSymbols(node, LSP.SymbolKind.Field, true));
			children.AddRange(GetMethodSymbols(node, propMapFunc));
			children.AddRange(GetTypeSymbols(node));
			children.AddRange(GetEnumSymbols(node));

			var symbolName = Path.GetFileNameWithoutExtension(uri.LocalPath);
			var ext = Path.GetExtension(uri.LocalPath);
			LSP.SymbolKind kind;
			if (ext == ".bas") {
				kind = LSP.SymbolKind.Module;
			} else if (ext == ".cls") {
				kind = LSP.SymbolKind.Class;
			} else {
				kind = LSP.SymbolKind.Module;
			}
			var rootSymbol = GetSymbol(node, symbolName, kind);
			rootSymbol.Children = [.. children];
			return [rootSymbol];
		}

		private static List<DocumentSymbol> GetMethodSymbols(SyntaxNode node, Func<int, (bool, string, string)> propMapFunc) {
			var symbols = new List<DocumentSymbol>();
			var methodSyntaxes = node.DescendantNodes().OfType<MethodBlockSyntax>();
			foreach (var syntax in methodSyntaxes) {
				var stmt = syntax.SubOrFunctionStatement;
				var name = stmt.Identifier.Text;
				var lineSpan = stmt.GetLocation().GetLineSpan();
				var sp = lineSpan.StartLinePosition;
				var (isPorp, prefix, propName) = propMapFunc(sp.Line);
				if (isPorp) {
					name = $"{prefix} {propName}";
					symbols.Add(GetSymbol(stmt, name, LSP.SymbolKind.Property));
				} else {
					var methodSymbol = GetSymbol(stmt, name, LSP.SymbolKind.Method);
					methodSymbol.Children = [.. GetLocalVarSymbols(syntax)];
					symbols.Add(methodSymbol);
				}
			}
			return symbols;
		}

		private static List<DocumentSymbol> GetVarSymbols(SyntaxNode node, LSP.SymbolKind kind, bool root) {
			var symbols = new List<DocumentSymbol>();
			var varSyntaxes = node.DescendantNodes().OfType<FieldDeclarationSyntax>();
			foreach (var syntax in varSyntaxes) {
				if (root) {
					if (syntax.Parent != null && syntax.Parent.IsKind(CAVB.SyntaxKind.StructureBlock)) {
						continue;
					}
				} else {
					if (node.IsKind(CAVB.SyntaxKind.ModuleBlock)) {
						continue;
					}
					if (node.IsKind(CAVB.SyntaxKind.ClassBlock)) {
						continue;
					}
				}
				foreach (var declarator in syntax.Declarators) {
					var name = declarator.Names.ToFullString();
					symbols.Add(GetSymbol(syntax, name, kind));
				}
			}
			return symbols;
		}

		private static List<DocumentSymbol> GetTypeSymbols(SyntaxNode node) {
			var symbols = new List<DocumentSymbol>();
			var structureSyntaxes = node.DescendantNodes().OfType<StructureBlockSyntax>();
			foreach (var syntax in structureSyntaxes) {
				var stmt = syntax.StructureStatement;
				var name = stmt.Identifier.Text;
				var typeSymbol = GetSymbol(stmt, name, LSP.SymbolKind.Struct);
				typeSymbol.Children = [.. GetVarSymbols(stmt.Parent, LSP.SymbolKind.Variable, false)];
				symbols.Add(typeSymbol);
			}
			return symbols;
		}

		private static List<DocumentSymbol> GetEnumSymbols(SyntaxNode node) {
			var symbols = new List<DocumentSymbol>();
			var enumSyntaxes = node.DescendantNodes().OfType<EnumBlockSyntax>();
			foreach (var syntax in enumSyntaxes) {
				var stmt = syntax.EnumStatement;
				var name = stmt.Identifier.Text;
				var enumSymbol = GetSymbol(stmt, name, LSP.SymbolKind.Enum);
				enumSymbol.Children = [..GetEnumVarSymbols(stmt.Parent)];
				symbols.Add(enumSymbol);
			}
			return symbols;
		}

		private static List<DocumentSymbol> GetLocalVarSymbols(SyntaxNode node) {
			var symbols = new List<DocumentSymbol>();
			var varSyntaxes = node.DescendantNodes().OfType<LocalDeclarationStatementSyntax>();
			foreach (var syntax in varSyntaxes) {
				foreach (var declarator in syntax.Declarators) {
					var name = declarator.Names.ToFullString();
					symbols.Add(GetSymbol(syntax, name, LSP.SymbolKind.Variable));
				}
			}
			return symbols;
		}

		private static List<DocumentSymbol> GetEnumVarSymbols(SyntaxNode node) {
			var symbols = new List<DocumentSymbol>();
			var varSyntaxes = node.DescendantNodes().OfType<EnumMemberDeclarationSyntax>();
			foreach (var syntax in varSyntaxes) {
				var name = syntax.Identifier.Text;
				symbols.Add(GetSymbol(node, name, LSP.SymbolKind.EnumMember));
			}
			return symbols;
		}

		private static DocumentSymbol GetSymbol(SyntaxNode node, string name, LSP.SymbolKind kind) {
			var spen = node.GetLocation().GetLineSpan();
			var sp = spen.StartLinePosition;
			var ep = spen.EndLinePosition;
			var range = new LSP.Range {
				Start = new Position { Line = sp.Line, Character = 0 },
				End = new Position { Line = ep.Line, Character = ep.Character },
			};
			return new DocumentSymbol {
				Name = name.Trim(),
				Kind = kind,
				Deprecated = false,
				Range = range,
				SelectionRange = range
			};
		}
	}
}
