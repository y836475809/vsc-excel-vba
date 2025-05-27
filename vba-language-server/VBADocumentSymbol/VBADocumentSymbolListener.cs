using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using static VBADocumentSymbol.VBADocumentSymbolParser;


namespace VBADocumentSymbol {	
	public class VBADocumentSymbolListener :  VBADocumentSymbolBaseListener {
		private List<(string, int)> FieldVariantList;
		private List<(string, string, int, int)> PropertyList;
		private List<(string, int, int, List<(string, int)>)> TypeList;
		private List<(string, int, int)> MethodList;

		public VBADocumentSymbolListener() {
			FieldVariantList = [];
			PropertyList = [];
			TypeList = [];
			MethodList = [];
		}

		public override void ExitStartRule([NotNull] StartRuleContext context) {
			base.ExitStartRule(context);
		}

		public override void ExitDimStmt([NotNull] DimStmtContext context) {
			var name = context.identifier().GetText();
			FieldVariantList.Add((name, context.Start.Line));
		}

		public override void ExitConstStmt([NotNull] ConstStmtContext context) {
			var name = context.identifier()[0].GetText();
			FieldVariantList.Add((name, context.Start.Line));
		}

		public override void ExitPropertyGetStmt([NotNull] PropertyGetStmtContext context) {
			var name = context.identifier().GetText();
			PropertyList.Add(("get", name, context.Start.Line, context.Stop.Line));
		}

		public override void ExitPropertySetStmt([NotNull] PropertySetStmtContext context) {
			var name = context.identifier().GetText();
			var propType = "set";
			if(context.LET() != null) {
				propType = "let";
			}
			PropertyList.Add((propType, name, context.Start.Line, context.Stop.Line));
		}

		public override void ExitSubStmt([NotNull] SubStmtContext context) {
			var name = context.identifier()?.GetText();
			MethodList.Add((name, context.Start.Line, context.Stop.Line));
		}

		public override void ExitFunctionStmt([NotNull] FunctionStmtContext context) {
			var name = context.identifier()?.GetText();
			MethodList.Add((name, context.Start.Line, context.Stop.Line));
		}

		public override void ExitTypeStmt([NotNull] TypeStmtContext context) {

		}
	}
}
