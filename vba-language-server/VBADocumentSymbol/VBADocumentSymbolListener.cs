using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design.Serialization;
using System.Linq;
using System.Xml.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using static System.Net.Mime.MediaTypeNames;
using static VBADocumentSymbol.VBADocumentSymbolParser;


namespace VBADocumentSymbol {
	public class DocumentSymbol : IDocumentSymbol {
		public DocumentSymbol() {
			Variables = [];
		}
	}


	public class VBADocumentSymbolListener :  VBADocumentSymbolBaseListener {
		private List<IDocumentSymbol> VariantList;
		public List<IDocumentSymbol>	 SymbolList;

		public VBADocumentSymbolListener() {
			VariantList = [];
			SymbolList = [];
		}

		public override void ExitStartRule([NotNull] StartRuleContext context) {
			base.ExitStartRule(context);
		}

		public override void ExitFiledVariant([NotNull] FiledVariantContext context) {
			var name = context.identifier().GetText();
			var text = context.GetText();
			var start = context.Start;
			var stop = context.Stop;
			VariantList.Add(new DocumentSymbol {
				Name = name,
				Kind = "FieldVariant",
				StartLine = start.Line - 1,
				StartColumn = start.Column,
				EndLine = stop.Line - 1,
				EndColumn = start.Column + text.Length
			});
		}

		public override void ExitDimStmt([NotNull] DimStmtContext context) {
			var name = context.identifier().GetText();
			var text = context.GetText();
			var start = context.Start;
			var stop = context.Stop;
			VariantList.Add(new DocumentSymbol {
				Name = name,
				Kind = "Variable",
				StartLine = start.Line - 1,
				StartColumn = start.Column,
				EndLine = stop.Line - 1,
				EndColumn = start.Column + text.Length
			});
		}

		public override void ExitConstStmt([NotNull] ConstStmtContext context) {
			var name = context.identifier().First().GetText();
			var text = context.GetText();
			var start = context.Start;
			var stop = context.Stop;
			VariantList.Add(new DocumentSymbol {
				Name = name,
				Kind = "Variable",
				StartLine = start.Line - 1,
				StartColumn = start.Column,
				EndLine = stop.Line - 1,
				EndColumn = start.Column + text.Length
			});
		}

		public override void ExitPropertyGetStmt([NotNull] PropertyGetStmtContext context) {
			SymbolList.AddRange(VariantList);
			VariantList.Clear();

			var name = $"Get {context.identifier().GetText()}";
			var text = context.GetText();
			var start = context.Start;
			var stop = context.Stop;
			SymbolList.Add(new DocumentSymbol {
				Name = name,
				Kind = "Property",
				StartLine = start.Line - 1,
				StartColumn = start.Column,
				EndLine = stop.Line - 1,
				EndColumn = start.Column + text.Length
			});
		}

		public override void ExitPropertySetStmt([NotNull] PropertySetStmtContext context) {
			SymbolList.AddRange(VariantList);
			VariantList.Clear();

			var propType = "Set";
			if(context.LET() != null) {
				propType = "Let";
			}
			var name = $"{propType} {context.identifier().GetText()}";
			var text = context.GetText();
			var start = context.Start;
			var stop = context.Stop;
			SymbolList.Add(new DocumentSymbol {
				Name = name,
				Kind = "Property",
				StartLine = start.Line - 1,
				StartColumn = start.Column,
				EndLine = stop.Line - 1,
				EndColumn = start.Column + text.Length
			});
		}

		public override void ExitEndPropertyStmt([NotNull] EndPropertyStmtContext context) {
			VariantList.Clear();

			if (!SymbolList.Any()) {
				return;
			}
			var symbol = SymbolList[SymbolList.Count - 1];
			if (symbol.Kind != "Property") {
				return;
			}
			symbol.EndLine = context.Start.Line - 1;
			symbol.EndColumn = context.Start.Column + context.GetText().Length;
		}

		public override void ExitSubStmt([NotNull] SubStmtContext context) {
			SymbolList.AddRange(VariantList);
			VariantList.Clear();

			var name = $"Sub {context.identifier().GetText()}";
			var text = context.GetText();
			var start = context.Start;
			var stop = context.Stop;
			SymbolList.Add(new DocumentSymbol {
				Name = name,
				Kind = "Method",
				StartLine = start.Line - 1,
				StartColumn = start.Column,
				EndLine = stop.Line - 1,
				EndColumn = start.Column + text.Length
			});
		}

		public override void ExitEndSubStmt([NotNull] EndSubStmtContext context) {
			VariantList.Clear();
			if (!SymbolList.Any()) {
				return;
			}
			var symbol = SymbolList[SymbolList.Count - 1];
			if(!symbol.Name.StartsWith("Sub ")){
				return;
			}
			symbol.EndLine = context.Start.Line - 1;
			symbol.EndColumn = context.Start.Column + context.GetText().Length;
		}

		public override void ExitFunctionStmt([NotNull] FunctionStmtContext context) {
			SymbolList.AddRange(VariantList);
			VariantList.Clear();

			var name = $"Function {context.identifier().GetText()}";
			var text = context.GetText();
			var start = context.Start;
			var stop = context.Stop;
			SymbolList.Add(new DocumentSymbol {
				Name = name,
				Kind = "Method",
				StartLine = start.Line - 1,
				StartColumn = start.Column,
				EndLine = stop.Line - 1,
				EndColumn = start.Column + text.Length
			});
		}

		public override void ExitEndFunctionStmt([NotNull] EndFunctionStmtContext context) {
			VariantList.Clear();
			if (!SymbolList.Any()) {
				return;
			}
			var symbol = SymbolList[SymbolList.Count - 1];
			if (!symbol.Name.StartsWith("Function ")){
				return;
			}
			symbol.EndLine = context.Start.Line - 1;
			symbol.EndColumn = context.Start.Column + context.GetText().Length;
		}

		public override void ExitTypeStmt([NotNull] TypeStmtContext context) {
			SymbolList.AddRange(VariantList);
			VariantList.Clear();

			var variables = new List<IDocumentSymbol>();
			foreach (var stmt in context.blockTypeStmt()) {
				var name = stmt.identifier().GetText();
				var text = stmt.GetText();
				var start = stmt.Start;
				var stop = stmt.Stop;
				variables.Add(new DocumentSymbol {
					Name = name,
					Kind = "Variable",
					StartLine = start.Line - 1,
					StartColumn = start.Column,
					EndLine = stop.Line - 1,
					EndColumn = start.Column + text.Length
				});
			}
			{
				var name = $"{context.identifier().GetText()}";
				var text = context.GetText();
				var start = context.Start;
				var stop = context.Stop;
				SymbolList.Add(new DocumentSymbol {
					Name = name,
					Kind = "Struct",
					StartLine = start.Line - 1,
					StartColumn = start.Column,
					EndLine = stop.Line - 1,
					EndColumn = start.Column + text.Length,
					Variables = variables
				});
			}
		}

		public override void ExitEnumStmt([NotNull] EnumStmtContext context) {
			SymbolList.AddRange(VariantList);
			VariantList.Clear();

			var variables = new List<IDocumentSymbol>();
			foreach (var stmt in context.blockEnumStmt()) {
				var name = stmt.identifier().GetText();
				var text = stmt.GetText();
				var start = stmt.Start;
				var stop = stmt.Stop;
				variables.Add(new DocumentSymbol {
					Name = name,
					Kind = "Enum",
					StartLine = start.Line - 1,
					StartColumn = start.Column,
					EndLine = stop.Line - 1,
					EndColumn = start.Column + text.Length
				});
			}
			{
				var name = $"{context.identifier().GetText()}";
				var text = context.GetText();
				var start = context.Start;
				var stop = context.Stop;
				SymbolList.Add(new DocumentSymbol {
					Name = name,
					Kind = "EnumMember",
					StartLine = start.Line - 1,
					StartColumn = start.Column,
					EndLine = stop.Line - 1,
					EndColumn = start.Column + text.Length,
					Variables = variables
				});
			}
		}
	}
}
