using Microsoft.VisualStudio.LanguageServer.Protocol;
using System;
using System.Collections.Generic;
using System.Linq;
using VBACodeAnalysis;
using Xunit;
using LSP = Microsoft.VisualStudio.LanguageServer.Protocol;


namespace TestProject {
	public class TestDocumentSymbolProvider {
		private static LSP.DocumentSymbol[] GetDocumentSymbols(string fileName) {
			var filePath = Helper.getPath(fileName);
			var vbaCode = Helper.getCode(fileName);
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			vbaca.setSetting(new RewriteSetting());
			var vbCode = vbaca.Rewrite(filePath, vbaCode);
			vbaca.AddDocument(filePath, vbCode);
			var symbols = vbaca.GetDocumentSymbols(filePath, new Uri(filePath));
			return symbols;
		}

		private static void AssertSymbol(DocumentSymbol[] symbols, List<(string, LSP.SymbolKind)> nameKindList) {
			Assert.Equal(nameKindList.Count, symbols.Length);
			var symNameKind = symbols.Zip(nameKindList, (first, second) => (first, second.Item1, second.Item2));
			foreach (var item in symNameKind) {
				var (symbol, name, kind) = item;
				Assert.Equal(name, symbol.Name);
				Assert.Equal(kind, symbol.Kind);
			}
		}

		[Fact]
		public void TestMethod() {
			var fileName = "test_document_symbol.bas";
			var symbols = GetDocumentSymbols(fileName);
			Assert.Equal(1, symbols.Length);
			var rootSym = symbols[0];
			Assert.Equal(LSP.SymbolKind.Module, rootSym.Kind);
			Assert.Equal("test_document_symbol", rootSym.Name);

			var rootSymChildren = rootSym.Children.OrderBy(x => x.Range.Start.Line).ToArray();
			Assert.Equal(11, rootSymChildren.Length);

			AssertSymbol(rootSymChildren, [
				("fieldvar1", LSP.SymbolKind.Field),
				("fieldvar2", LSP.SymbolKind.Field),
				("fieldvar3", LSP.SymbolKind.Field),

				("testEnum", LSP.SymbolKind.Enum),

				("type1", LSP.SymbolKind.Struct),
				("type2", LSP.SymbolKind.Struct),
				("type3", LSP.SymbolKind.Struct),

				("Get prop_get1", LSP.SymbolKind.Property),
				("Let prop_let1", LSP.SymbolKind.Property),
				("Set prop_set1", LSP.SymbolKind.Property),

				("Main", LSP.SymbolKind.Method),
			]);

			var enumSym = rootSymChildren[3];
			var type1Sym = rootSymChildren[4];
			var type2Sym = rootSymChildren[5];
			var type3Sym = rootSymChildren[6];
			var prop1Sym = rootSymChildren[7];
			var prop2Sym = rootSymChildren[8];
			var prop3Sym = rootSymChildren[9];
			var methodSym = rootSymChildren[10];

			AssertSymbol(enumSym.Children, [
				("e1", LSP.SymbolKind.EnumMember),
				("e2", LSP.SymbolKind.EnumMember),
				("e3", LSP.SymbolKind.EnumMember),
			]);

			AssertSymbol(type1Sym.Children, [
				("num1", LSP.SymbolKind.Variable),
				("name1", LSP.SymbolKind.Variable),
				("utc1()", LSP.SymbolKind.Variable),
			]);

			AssertSymbol(type2Sym.Children, [
				("num2", LSP.SymbolKind.Variable),
				("name2", LSP.SymbolKind.Variable),
			]);

			AssertSymbol(type3Sym.Children, [
				("num3", LSP.SymbolKind.Variable),
				("name3", LSP.SymbolKind.Variable),
			]);

			Assert.Null(prop1Sym.Children);
			Assert.Null(prop2Sym.Children);
			Assert.Null(prop3Sym.Children);

			AssertSymbol(methodSym.Children, [
				("num1", LSP.SymbolKind.Variable),
				("name1", LSP.SymbolKind.Variable),
				("End_Row", LSP.SymbolKind.Variable),
				("End_Col", LSP.SymbolKind.Variable),
			]);
		}
	}
}
