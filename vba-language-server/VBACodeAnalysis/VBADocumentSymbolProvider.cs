using Antlr4.Runtime.Atn;
using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using VBADocumentSymbol;

namespace VBACodeAnalysis {
	internal class DocumentSymbol : IDocumentSymbol {
		public DocumentSymbol() {
			Children = [];
		}
	}

	class DocumentSymbolProvider {
		public static IDocumentSymbol GetRoot(Uri uri, string vbaCode) {
			var symbolName = Path.GetFileNameWithoutExtension(uri.LocalPath);
			var ext = Path.GetExtension(uri.LocalPath);
			string kind = "Module";
			if (ext == ".bas") {
				kind = "Module";
			} else if (ext == ".cls") {
				kind = "Class";
			}
			var lines = vbaCode.Split(Environment.NewLine);
			var symbols = GetSymbols(vbaCode);
			var root = new DocumentSymbol {
				Name = symbolName,
				Kind = kind,
				StartLine = 0,
				StartColumn = 0,
				EndLine = lines.Length - 1,
				EndColumn = lines.Last().Length,
				Children = symbols
			};
			return root;
		}

		private static List<IDocumentSymbol> GetSymbols(string vbaCode) {
			var lexer = new VBADocumentSymbolLexer(new AntlrInputStream(vbaCode));
			var tokens = new CommonTokenStream(lexer);
			var parser = new VBADocumentSymbolParser(tokens);
			parser.Interpreter.PredictionMode = PredictionMode.SLL;
			lexer.RemoveErrorListeners();
			parser.RemoveErrorListeners();

			var nn = new VBADocumentSymbolListener();
			parser.AddParseListener(nn);
			try {
				parser.startRule();
			} catch (Exception) {
				tokens.Reset();
				parser.Reset();
				parser.Interpreter.PredictionMode = PredictionMode.LL;
				parser.startRule();
			}
			return nn.SymbolList;
		}
	}
}
