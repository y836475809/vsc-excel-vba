﻿using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Elfie.Model;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace VBACodeAnalysis {
	class DocumentSymbolProvider {
		public static List<VBADocSymbol> GetDocumentSymbols(
			SyntaxNode node, Uri uri, 
			Func<int, (bool, string, string)> propMapFunc) {
			var children = new List<VBADocSymbol>();
			children.AddRange(GetVarSymbols(node, "Field", true));
			children.AddRange(GetMethodSymbols(node, propMapFunc));
			children.AddRange(GetTypeSymbols(node));
			children.AddRange(GetEnumSymbols(node));

			var symbolName = Path.GetFileNameWithoutExtension(uri.LocalPath);
			var ext = Path.GetExtension(uri.LocalPath);
			string kind = "Module";
			if (ext == ".bas") {
				kind = "Module";
			} else if (ext == ".cls") {
				kind = "Class";
			}
			var rootSymbol = GetSymbol(node, symbolName, kind);
			rootSymbol.Children = [.. children];
			return [rootSymbol];
		}

		private static List<VBADocSymbol> GetMethodSymbols(SyntaxNode node, Func<int, (bool, string, string)> propMapFunc) {
			var symbols = new List<VBADocSymbol>();
			var methodSyntaxes = node.DescendantNodes().OfType<MethodBlockSyntax>();
			foreach (var syntax in methodSyntaxes) {
				var stmt = syntax.SubOrFunctionStatement;
				var name = stmt.Identifier.Text;
				var lineSpan = stmt.GetLocation().GetLineSpan();
				var sp = lineSpan.StartLinePosition;
				var (isPorp, prefix, propName) = propMapFunc(sp.Line);
				if (isPorp) {
					name = $"{prefix} {propName}";
					symbols.Add(GetSymbol(stmt, name, "Property"));
				} else {
					var methodSymbol = GetSymbol(stmt, name, "Method");
					methodSymbol.Children = [.. GetLocalVarSymbols(syntax)];
					symbols.Add(methodSymbol);
				}
			}
			return symbols;
		}

		private static List<VBADocSymbol> GetVarSymbols(SyntaxNode node, string kind, bool root) {
			var symbols = new List<VBADocSymbol>();
			var varSyntaxes = node.DescendantNodes().OfType<FieldDeclarationSyntax>();
			foreach (var syntax in varSyntaxes) {
				if (root) {
					if (syntax.Parent != null && syntax.Parent.IsKind(SyntaxKind.StructureBlock)) {
						continue;
					}
				} else {
					if (node.IsKind(SyntaxKind.ModuleBlock)) {
						continue;
					}
					if (node.IsKind(SyntaxKind.ClassBlock)) {
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

		private static List<VBADocSymbol> GetTypeSymbols(SyntaxNode node) {
			var structureSyntaxes = node.DescendantNodes().OfType<StructureBlockSyntax>();
			var symbols = structureSyntaxes.Select(x => {
				var stmt = x.StructureStatement;
				var name = stmt.Identifier.Text;
				var typeSymbol = GetSymbol(stmt, name, "Struct");
				typeSymbol.Children = GetVarSymbols(stmt.Parent, "Variable", false);
				return typeSymbol;
			});
			return [.. symbols];
		}

		private static List<VBADocSymbol> GetEnumSymbols(SyntaxNode node) {
			var enumSyntaxes = node.DescendantNodes().OfType<EnumBlockSyntax>();
			var symbols = enumSyntaxes.Select(x => {
				var stmt = x.EnumStatement;
				var name = stmt.Identifier.Text;
				var enumSymbol = GetSymbol(stmt, name, "Enum");
				enumSymbol.Children = GetEnumVarSymbols(stmt.Parent);
				return enumSymbol;
			});
			return [..symbols];
		}

		private static List<VBADocSymbol> GetLocalVarSymbols(SyntaxNode node) {
			var varSyntaxes = node.DescendantNodes().OfType<LocalDeclarationStatementSyntax>();
			var symbols = new List<VBADocSymbol>();
			foreach (var syntax in varSyntaxes) {
				foreach (var declarator in syntax.Declarators) {
					var name = declarator.Names.ToFullString();
					symbols.Add(GetSymbol(syntax, name, "Variable"));
				}
			}
			return symbols;
		}

		private static List<VBADocSymbol> GetEnumVarSymbols(SyntaxNode node) {
			var varSyntaxes = node.DescendantNodes().OfType<EnumMemberDeclarationSyntax>();
			var symbols = varSyntaxes.Select(x => {
				var name = x.Identifier.Text;
				return GetSymbol(node, name, "EnumMember");
			});
			return [..symbols];
		}

		private static VBADocSymbol GetSymbol(SyntaxNode node, string name, string kind) {
			var spen = node.GetLocation().GetLineSpan();
			var sp = spen.StartLinePosition;
			var ep = spen.EndLinePosition;
			return new() {
				Name = name.Trim(),
				Kind = kind,
				Start = (sp.Line, 0),
				End = (ep.Line, ep.Character),
				Children = []
			};
		}
	}
}
