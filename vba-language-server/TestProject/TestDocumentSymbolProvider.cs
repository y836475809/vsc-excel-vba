using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Xml.Linq;
using VBACodeAnalysis;
using VBADocumentSymbol;
using Xunit;

namespace TestProject {
	public class TestDocumentSymbolProvider {
		private static IDocumentSymbol GetDocumentSymbol(string fileName) {
			var filePath = Helper.getPath(fileName);
			var vbaCode = Helper.getCode(fileName);
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			var symbol = vbaca.GetDocumentSymbols(new Uri(filePath), vbaCode);
			return symbol;
		}

		private static void AssertSymbol(List<IDocumentSymbol> symbols, 
			List<(string, string)> nameKindList) {
			Assert.Equal(nameKindList.Count, symbols.Count);
			var symNameKind = symbols.Zip(nameKindList, (first, second) => (first, second.Item1, second.Item2));
			foreach (var item in symNameKind) {
				var (symbol, name, kind) = item;
				Assert.Equal(name, symbol.Name);
				Assert.Equal(kind, symbol.Kind);
			}
		}

		private static void AssertSymbol(List<IDocumentSymbol> actSymbolList, 
			List<(string, string, (int, int), (int, int))> expSymbolList) {
			Assert.Equal(expSymbolList.Count, actSymbolList.Count);
			for (int i = 0; i < expSymbolList.Count; i++) {
				var (expName, expKind, expStart, expEnd)  = expSymbolList[i];
				var actSymbol = actSymbolList[i];
				Assert.Equal(expName, actSymbol.Name);
				Assert.Equal(expKind, actSymbol.Kind);
				Assert.Equal(expStart.Item1, actSymbol.StartLine);
				Assert.Equal(expStart.Item2, actSymbol.StartColumn);
				Assert.Equal(expEnd.Item1, actSymbol.EndLine);
				Assert.Equal(expEnd.Item2, actSymbol.EndColumn);
			}
		}

		[Fact]
		public void TestDocumentSymbol1() {
			var fileName = "test_document_symbol1.bas";
			var rootSym = GetDocumentSymbol(fileName);

			Assert.Equal("Module", rootSym.Kind);
			Assert.Equal("test_document_symbol1", rootSym.Name);
			Assert.Equal(16, rootSym.Variables.Count);

			var children = rootSym.Variables;
			AssertSymbol(children, [
				("field_var1", "Variable"),
				("field_var2", "Variable"),
				("field_var3", "Variable"),
				("field_dim_var4", "Variable"),
				("field_const_var5", "Variable"),

				("testEnum", "Enum"),

				("type1", "Struct"),
				("type2", "Struct"),
				("type3", "Struct"),

				("Get prop_get_set1", "Property"),
				("Set prop_get_set1", "Property"),
				("Get prop_get1", "Property"),
				("Let prop_let1", "Property"),
				("Set prop_set1", "Property"),

				("Function func1", "Method"),
				("Sub sub1", "Method"),
			]);

			var enumSym = children[5];
			var type1Sym = children[6];
			var type2Sym = children[7];
			var type3Sym = children[8];
			var prop1Sym = children[9];
			var prop2Sym = children[10];
			var prop3Sym = children[11];
			var prop4Sym = children[12];
			var prop5Sym = children[13];
			var funcSym = children[14];
			var subSym = children[15];

			AssertSymbol(enumSym.Variables, [
				("e1", "EnumMember"),
				("e2", "EnumMember"),
				("e3", "EnumMember"),
			]);

			AssertSymbol(type1Sym.Variables, [
				("num1", "Variable"),
				("name1", "Variable"),
				("utc1", "Variable"),
			]);

			AssertSymbol(type2Sym.Variables, [
				("num2", "Variable"),
				("name2", "Variable"),
			]);

			AssertSymbol(type3Sym.Variables, [
				("num3", "Variable"),
				("name3", "Variable"),
			]);

			Assert.Empty(prop1Sym.Variables);
			Assert.Empty(prop2Sym.Variables);
			Assert.Empty(prop3Sym.Variables);
			Assert.Empty(prop4Sym.Variables);
			Assert.Empty(prop5Sym.Variables);
			Assert.Empty(funcSym.Variables);
			Assert.Empty(subSym.Variables);
		}

		[Fact]
		public void TestDocumentSymbol2() {
			var fileName = "test_document_symbol2.bas";
			var rootSym = GetDocumentSymbol(fileName);

			Assert.Equal("Module", rootSym.Kind);
			Assert.Equal("test_document_symbol2", rootSym.Name);
			Assert.Equal(11, rootSym.Variables.Count);

			var children = rootSym.Variables;
			AssertSymbol(children, [
				("field_var1", "Variable", (3, 8), (3, 18)),
				("field_var2", "Variable", (4, 8), (4, 18)),
				("field_var3", "Variable", (4, 30), (4, 40)),
				("field_dim_var4", "Variable", (5, 4), (5, 18)),
				("field_const_var5", "Variable", (6, 6), (6, 22)),

				("testEnum", "Enum", (8, 0), (11, 8)),

				("type1", "Struct", (13, 0), (16, 8)),

				("Get prop_get_set1", "Property", (18, 0), (22, 12)),
				("Set prop_get_set1", "Property", (24, 0), (27, 12)),

				("Function func1", "Method", (29, 0), (32, 12)),
				("Sub sub1", "Method", (34, 0), (37, 7)),
			]);

			var enumSym = children[5];
			var type1Sym = children[6];
			var prop1Sym = children[7];
			var prop2Sym = children[8];
			var funcSym = children[9];
			var subSym = children[10];

			AssertSymbol(enumSym.Variables, [
				("e1", "EnumMember", (9, 2), (9, 4)),
				("e2", "EnumMember", (10, 2), (10, 4)),
			]);

			AssertSymbol(type1Sym.Variables, [
				("num1", "Variable", (14, 2), (14, 6)),
				("utc1", "Variable", (15, 2), (15, 6)),
			]);

			Assert.Empty(prop1Sym.Variables);
			Assert.Empty(prop2Sym.Variables);
			Assert.Empty(funcSym.Variables);
			Assert.Empty(subSym.Variables);
		}
	}
}
