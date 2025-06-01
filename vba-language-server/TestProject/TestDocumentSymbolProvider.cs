using Microsoft.VisualStudio.LanguageServer.Protocol;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using VBACodeAnalysis;
using VBADocumentSymbol;
using Xunit;

namespace TestProject {
	public class TestDocumentSymbolProvider {
		private static IDocumentSymbol GetDocumentSymbol(string fileName) {
			var filePath = Helper.getPath(fileName);
			var vbaCode = Helper.getCode(fileName);
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			//vbaca.setSetting(new RewriteSetting());
			//var vbCode = vbaca.Rewrite(filePath, vbaCode);
			//vbaca.AddDocument(filePath, vbCode);
			//var symbols = vbaca.GetDocumentSymbols(filePath, new Uri(filePath));
			var symbol = vbaca.GetDocumentSymbols(new Uri(filePath), vbaCode);
			return symbol;
		}

		private static void AssertSymbol(List<VBADocSymbol> symbols, List<(string, string)> nameKindList) {
			Assert.Equal(nameKindList.Count, symbols.Count);
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
			var rootSym = GetDocumentSymbol(fileName);
			//Assert.Single(symbols);
			var vbaSymbol = VBALanguageServer.Util.ToLspDocumentSymbol(rootSym);
			//var rootSym = symbols[0];
			Assert.Equal(SymbolKind.Module, vbaSymbol.Kind);
			Assert.Equal("test_document_symbol", vbaSymbol.Name);

			var rootSymChildren = vbaSymbol.Children.OrderBy(x => x.Range.Start.Line).ToList();
			Assert.Equal(13, rootSymChildren.Count);

			//AssertSymbol(rootSymChildren, [
			//	("fieldvar1", "Field"),
			//	("fieldvar2", "Field"),
			//	("fieldvar3", "Field"),

			//	("testEnum", "Enum"),

			//	("type1", "Struct"),
			//	("type2", "Struct"),
			//	("type3", "Struct"),

			//	("Get prop_get_set1", "Property"),
			//	("Set prop_get_set1", "Property"),
			//	("Get prop_get1", "Property"),
			//	("Let prop_let1", "Property"),
			//	("Set prop_set1", "Property"),

			//	("Main", "Method"),
			//]);

			//var enumSym = rootSymChildren[3];
			//var type1Sym = rootSymChildren[4];
			//var type2Sym = rootSymChildren[5];
			//var type3Sym = rootSymChildren[6];
			//var prop1Sym = rootSymChildren[7];
			//var prop2Sym = rootSymChildren[8];
			//var prop3Sym = rootSymChildren[9];
			//var prop4Sym = rootSymChildren[10];
			//var prop5Sym = rootSymChildren[11];
			//var methodSym = rootSymChildren[12];

			//AssertSymbol(enumSym.Children, [
			//	("e1", "EnumMember"),
			//	("e2", "EnumMember"),
			//	("e3", "EnumMember"),
			//]);

			//AssertSymbol(type1Sym.Children, [
			//	("num1", "Variable"),
			//	("name1", "Variable"),
			//	("utc1()", "Variable"),
			//]);

			//AssertSymbol(type2Sym.Children, [
			//	("num2", "Variable"),
			//	("name2", "Variable"),
			//]);

			//AssertSymbol(type3Sym.Children, [
			//	("num3", "Variable"),
			//	("name3", "Variable"),
			//]);

			//Assert.Empty(prop1Sym.Children);
			//Assert.Empty(prop2Sym.Children);
			//Assert.Empty(prop3Sym.Children);
			//Assert.Empty(prop4Sym.Children);
			//Assert.Empty(prop5Sym.Children);

			//AssertSymbol(methodSym.Children, [
			//	("num1","Variable"),
			//	("name1","Variable"),
			//	("End_Row", "Variable"),
			//	("End_Col","Variable"),
			//]);
		}
	}
}
