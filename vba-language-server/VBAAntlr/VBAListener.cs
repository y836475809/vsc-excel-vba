using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using AntlrTemplate;
using static VBAAntlr.VBAParser;


namespace VBAAntlr {
	public enum ModuleType {
		Cls,
		Bas,
	}

	public interface IPropertyDiagnostic {
		string Id { get; }
		string Code { get; }
		string Severity { get; }
		int Line { get; }
	}

	public interface IRewriteVBA {
		/// <summary>
		/// 部分置換
		/// </summary>
		/// <param name="lineIndex"></param>
		/// <param name="repColRange"></param>
		/// <param name="text"></param>
		/// <param name="startCol"></param>
		/// <param name="enableShift"></param>
		void AddChange(int lineIndex, (int, int) repColRange, string text, int startCol, bool enableShift=true);

		void AddChange(int lineIndex, (int, int) repColRange, string text, int startCol, int shiftCol);

		/// <summary>
		/// 行ごと置換、ColShifなし
		/// </summary>
		/// <param name="lineIndex"></param>
		/// <param name="text"></param>
		void AddChange(int lineIndex, string text);

		void AddLineMap(int srcLine, int toLine);

		void AddPropertyMember(string text, int srcLine);

		void AddModuleAttribute(int lastLineIndex, string vbName, ModuleType type);

		void SetAttributeVBName(int line, int startChara, int endChara, string vbName);

		void FoundOption();

		void AddIgnoreDiagnostic((int, int) start, (int, int) end, string text);

		void AddIgnoreDiagnostic(IPropertyDiagnostic propertyDiagnostic);
	}

	public class VBAListener : VBABaseListener {
		public IRewriteVBA rewriteVBA;
		private RewriteDynamicArray rewriteDynamicArray;
		private RewriteProperty rewriteProperty;

		public VBAListener(IRewriteVBA rewriteVBA) {
			this.rewriteVBA = rewriteVBA;
			rewriteDynamicArray = new();
			rewriteProperty = new();
		}
		public override void ExitDimStmt([NotNull] VBAParser.DimStmtContext context) {
			rewriteDynamicArray.Add(context);
		}

		public override void ExitRedimStmt([NotNull] VBAParser.RedimStmtContext context) {
			rewriteDynamicArray.Add(context);
		}

		public override void ExitSubStmt([NotNull] VBAParser.SubStmtContext context) {
			rewriteDynamicArray.StartBlockStmt();
		}

		public override void ExitEndSubStmt([NotNull] VBAParser.EndSubStmtContext context) {
			rewriteDynamicArray.EndBlockStmt();
		}

		public override void ExitFunctionStmt([NotNull] VBAParser.FunctionStmtContext context) {
			rewriteDynamicArray.StartBlockStmt();
		}

		public override void ExitEndFunctionStmt([NotNull] VBAParser.EndFunctionStmtContext context) {
			rewriteDynamicArray.EndBlockStmt();
		}

		public override void ExitStartRule([NotNull] VBAParser.StartRuleContext context) {
			base.ExitStartRule(context);
			var setTokens = context.GetTokens(VBAParser.SET);
			var letTokens = context.GetTokens(VBAParser.LET);
			var tokens = setTokens.Concat( letTokens );
			GetLetSet(tokens);
			GetVBAFunction(context.children);
			GetPredefined(context.children);
			GetFilenumber(context);
			GetVariant(context);

			rewriteProperty.Rewrite(rewriteVBA);
			rewriteDynamicArray.Rewrite(rewriteVBA);
		}
		public override void ExitModuleAttributes([NotNull] VBAParser.ModuleAttributesContext context) {
			var attrs = context.attributeStmt();
			var s = attrs.Where(x => Util.Eq(x.identifier()[0].GetText(), "VB_Name"));
			if (s.Any()) {
				var lastLineIndex = attrs.Max(x => x.Start.Line) - 1;
				var vbNameIdent = s.First().identifier()[1];
				var vbName = vbNameIdent.GetText();
				var vbNameStart = vbNameIdent.Start;
				var type = attrs.Length > 1 ? ModuleType.Cls : ModuleType.Bas;
				rewriteVBA.AddModuleAttribute(lastLineIndex, vbName, type);
				rewriteVBA.SetAttributeVBName(
					vbNameStart.Line,
					vbNameStart.Column, vbNameStart.Column + vbName.Length, 
					vbName.Trim('"'));
			}
		}

		public override void ExitModuleOption([NotNull] VBAParser.ModuleOptionContext context) {
			var st = context.Start;
			rewriteVBA.AddChange(st.Line - 1, "");
			rewriteVBA.FoundOption();
		}

		public override void ExitTypeStmt([NotNull] VBAParser.TypeStmtContext context) {
			GetTypeMember(context);

			var typeStmt = context.TYPE();
			var sym = typeStmt.Symbol;

			var name = context.identifier();
			var st = name.Start;
			var s = sym.Column;
			var e = name.Start.Column;
			var ws_count = name.Start.Column - (s + sym.Text.Length);
			var text = $"Structure{new string(' ', ws_count)}";
			rewriteVBA.AddChange(st.Line - 1, (s, e), text, st.Column);

			var end_stm = context.typeEndStmt();
			var end_s = end_stm.Start;
			rewriteVBA.AddChange(end_s.Line - 1, "End Structure");
		}

		private void GetTypeMember(VBAParser.TypeStmtContext context) {
			var typeVariables = context.blockTypeStmt();
			foreach (var item in typeVariables) {
				if (item.visibility() != null) {
					continue;
				}
				var ident = item.identifier();
				if (ident == null) {
					continue;
				}
				{
					var st = ident.Start;
					var s = st.Column;
					rewriteVBA.AddChange(st.Line - 1, (s, s), "Public ", s);
				}

				var asType = item.asTypeClause()?.identifier();
				if (asType == null) {
					continue;
				}
				var asTypeName = asType.GetText();
				if (Util.Eq(asTypeName, "variant")){
					var st = asType.Start;
					var s = st.Column;
					var e = s + asTypeName.Length;
					rewriteVBA.AddChange(st.Line - 1, (s, e), "Object ", s, false);
				}
			}
		}

		public override void ExitPropertyGetStmt([NotNull] VBAParser.PropertyGetStmtContext context) {
			var asType = context.asTypeClause()?.identifier()?.GetText();
			if (GetVariant(context.asTypeClause()?.identifier())) {
				asType = "Object";
			} 

			rewriteProperty.AddProperty(PropertyType.Get, context);
			rewriteDynamicArray.StartBlockStmt();
		}
		public override void ExitPropertySetStmt([NotNull] VBAParser.PropertySetStmtContext context) {
			rewriteProperty.AddProperty(PropertyType.Set, context);
			rewriteDynamicArray.StartBlockStmt();
		}

		public override void ExitEndPropertyStmt([NotNull] VBAParser.EndPropertyStmtContext context) {
			rewriteProperty.AddProperty(PropertyType.End, context);
			rewriteDynamicArray.EndBlockStmt();
		}

		public override void ExitOpenStmt([NotNull] VBAParser.OpenStmtContext context) {
			var open = context.OPEN();
			var line = open.Symbol.Line - 1;
			var startCol = open.Symbol.Column;
			var endtCol = startCol + open.Symbol.Text.Length;
			rewriteVBA.AddIgnoreDiagnostic(
				(line, startCol), (line, endtCol),
				open.GetText());

			var fileNum = context.fileNumber();
			var fnText = fileNum.identifier().GetText();
			if (fnText.StartsWith('#')) {
				var st = fileNum.Start;
				rewriteVBA.AddChange(st.Line - 1,
					(st.Column, st.Column + 1), ",", st.Column, false);
			} else {
				var st = fileNum.Start;
				rewriteVBA.AddChange(st.Line - 1,
					(st.Column-1, st.Column), ",", st.Column, false);
			}
		}

		public override void ExitOutputStmt([NotNull] VBAParser.OutputStmtContext context) {
			var method = context.OUTPUT();
			var line = method.Symbol.Line - 1;
			var startCol = method.Symbol.Column;
			var endtCol = startCol + method.Symbol.Text.Length;
			rewriteVBA.AddIgnoreDiagnostic(
				(line, startCol), (line, endtCol),
				method.GetText());

			var fileNum = context.fileNumber();
			var fnText = fileNum.identifier().GetText();
			if (fnText.StartsWith('#')) {
				var st = fileNum.Start;
				rewriteVBA.AddChange(st.Line - 1,
					(st.Column, st.Column + 1), " ", st.Column, false);
			}
		}

		public override void ExitInputStmt([NotNull] VBAParser.InputStmtContext context) {
			var method = context.INPUT();
			var line = method.Symbol.Line - 1;
			var startCol = method.Symbol.Column;
			var endtCol = startCol + method.Symbol.Text.Length;
			rewriteVBA.AddIgnoreDiagnostic(
				(line, startCol), (line, endtCol),
				method.GetText());

			var fileNum = context.fileNumber();
			var fnText = fileNum.identifier().GetText();
			if (fnText.StartsWith('#')) {
				var st = fileNum.Start;
				rewriteVBA.AddChange(st.Line - 1,
					(st.Column, st.Column + 1), " ", st.Column, false);
			}
		}

		public override void ExitLineInputStmt([NotNull] VBAParser.LineInputStmtContext context) {
			var method = context.LINE_INPUT();
			var line = method.Symbol.Line - 1;
			var startCol = method.Symbol.Column;
			var endtCol = startCol + method.Symbol.Text.Length;
			rewriteVBA.AddIgnoreDiagnostic(
				(line, startCol), (line, endtCol),
				"Line_Input");

			rewriteVBA.AddChange(line,
				(startCol, endtCol), "Line_Input", startCol, false);

			var fileNum = context.fileNumber();
			var fnText = fileNum.identifier().GetText();
			if (fnText.StartsWith('#')) {
				var st = fileNum.Start;
				rewriteVBA.AddChange(st.Line - 1,
					(st.Column, st.Column + 1), " ", st.Column, false);
			}
		}

		public override void ExitCloseStmt([NotNull] VBAParser.CloseStmtContext context) {
			var method = context.CLOSE();
			var line = method.Symbol.Line - 1;
			var startCol = method.Symbol.Column;
			var endtCol = startCol + method.Symbol.Text.Length;
			rewriteVBA.AddIgnoreDiagnostic(
				(line, startCol), (line, endtCol),
				method.GetText());

			var fileNums = context.fileNumber();
			foreach (var item in fileNums) {
				var fnText = item.identifier().GetText();
				if (fnText.StartsWith('#')) {
					var st = item.Start;
					rewriteVBA.AddChange(st.Line - 1,
						(st.Column, st.Column + 1), " ", st.Column, false);
				}
			}
		}

		private void GetLetSet(IEnumerable<ITerminalNode> tokens) {
			foreach (var item in tokens) {
				var sym = item.Symbol;
				var lineIndex = sym.Line - 1;
				var s = sym.Column;
				var e = s + sym.Text.Length;
				rewriteVBA.AddChange(lineIndex, (s, e), new string(' ', e - s), e);
			}
		}

		private void GetVBAFunction(IList<IParseTree> children) {
			var nl = Environment.NewLine;
			var traget_list = new List<string> { 
				"range", "cells", "columns", "workbooks", "worksheets" };
			var traget_set = new List<string> { "as", "new" };
			var vbaFuncList = children.Where((x, index) => {
				if(!Util.Contains(x.GetText(), traget_list)) {
					return false;
				}
				var ct = x.Payload as CommonToken;
				var idx1 = index - 2;
				if (idx1 < 0) {
					return false;
				}
				if (Util.Contains(children.ElementAt(idx1).GetText(), traget_set)) {
					return false;
				}

				var pre_ch1 = children.ElementAt(index - 1);
				var pre_text1 = pre_ch1.GetText().Replace(nl, "").Trim();
				if (pre_text1 ==  ".") {
					return false;
				}
				var pre_ch2 = children.ElementAt(index - 2);
				var pre_text2 = pre_ch2.GetText().Replace(nl, "").Trim();
				if (pre_text1 == "_" && pre_text2 == ".") {
					return false;
				}

				return true;
			});
			foreach (var item in vbaFuncList) {
				var ct = item.Payload as CommonToken;
				var s = ct.Column;
				rewriteVBA.AddChange(ct.Line - 1, (s, s), "f.", s);
			}
		}

		private void GetPredefined(IList<IParseTree> children) {
			var traget_list = new List<string> { 
				"cstr", "cdbl", "cbyte", "cint", "clng", "cbool",
				"strconv", "val", "createobject" };

			var predefinedList = children.Where(x => Util.Contains(x.GetText(), traget_list));
			foreach (var item in predefinedList) {
				var ct = item.Payload as CommonToken;
				var s = ct.Column;
				rewriteVBA.AddChange(ct.Line - 1, (s, s), "ExcelVBAFunctions.", s);
			}
			// TODO
			var lineInputItems = children.Where(
				x => (x.Payload as CommonToken)?.Type == VBAParser.LINE_INPUT);
			foreach (var item in lineInputItems) {
				var ct = item.Payload as CommonToken;
				var text = ct.Text;
				var s = ct.Column;
				var e = ct.Column + text.Length;
				var len =  text.Length - "line_input".Length;
				var repText = $"line_input{new string(' ', len)}";
				rewriteVBA.AddChange(ct.Line - 1, (s, e), repText, s, false);
			}
		}

		private void GetFilenumber(ParserRuleContext context) {
			var fnTokens = context.children.Select((x, Index) => (Index, x))
				.Where(x => x.x.Payload is CommonToken { Type: VBAParser.IDENTIFIER })
				.Where(x => x.x.GetText().StartsWith('#'));
			foreach (var item in fnTokens) {
				var ni = item.Index + 1;
				var neide = context.GetChild(ni);
				var next_ct = neide?.Payload as CommonToken;
				if (next_ct?.Type != VBAParser.IDENTIFIER) {
					var ct = item.x.Payload as CommonToken;
					var s = ct.Column;
					rewriteVBA.AddChange(ct.Line - 1, (s, s+1), " ", s+1);
				}
			}
		}

		private void GetVariant(ParserRuleContext context) {
			var fnTokens = context.children
				.Where(x => x.Payload is CommonToken { Type: VBAParser.IDENTIFIER })
				.Where(x => Util.Eq(x.GetText(), "variant"));
			foreach (var item in fnTokens) {
				var ct = item.Payload as CommonToken;
				var s = ct.Column;
				var e = s + ct.Text.Length;
				rewriteVBA.AddChange(ct.Line - 1, (s, e), "Object ", s, false);
			}
		}
		private bool GetVariant(VBAParser.IdentifierContext context) {
			const string val = "variant";
			if(context == null) {
				return false;
			}
			if (!Util.Eq(context.GetText(), val)){
				return false;
			}
			var st = context.Start;
			var s = st.Column;
			var e = s + context.GetText().Length;
			rewriteVBA.AddChange(st.Line - 1, (s, e), "Object ", s, false);
			return true;
		}
	}
}
