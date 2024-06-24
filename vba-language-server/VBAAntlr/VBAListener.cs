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


namespace VBAAntlr {
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

		/// <summary>
		/// 行ごと置換、ColShifなし
		/// </summary>
		/// <param name="lineIndex"></param>
		/// <param name="text"></param>
		void AddChange(int lineIndex, string text);

		void AddPropertyName(int lineIndex, string text, string asType);
	}

	public class VBAListener : VBABaseListener {
		public IRewriteVBA rewriteVBA;

		public VBAListener(IRewriteVBA rewriteVBA) {
			this.rewriteVBA = rewriteVBA;
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
            base.ExitPropertyGetStmt(context);
			GetVBAFunction(context.children);
			GetPredefined(context.children);
			GetFilenumber(context);
			GetVariant(context);

			var name = context.identifier();
			var asType = context.asTypeClause()?.identifier()?.GetText();
			var elems = context.blockLetSetStmt();
			var end_stm = context.endPropertyStmt();

			if (GetVariant(context.asTypeClause()?.identifier())) {
				asType = "Object";
			} 

			rewriteVBA.AddPropertyName(name.Start.Line - 1, name.GetText(), asType);

			foreach ( var e in elems) {
				var let_stmt = e.letStmt();
				if ( let_stmt != null && let_stmt.identifier().Length > 0) {
					var idens = let_stmt.identifier();
					GetFilenumber(idens[1]);
					var iden = idens[0];
					if (name.GetText() != iden.GetText()) {
						if (let_stmt.LET() != null) {
							var sym = let_stmt.LET().Symbol;
							var lineIndex = sym.Line - 1;
							var sym_s = sym.Column;
							var sym_e = sym_s + sym.Text.Length;
							rewriteVBA.AddChange(lineIndex, (sym_s, sym_e), new string(' ', sym_e - sym_s), sym_e);
						}
						continue;
					}
					var s = iden.Start;
					rewriteVBA.AddChange(
						s.Line - 1, (s.Column, s.Column), 
						"Get", s.Column);
				}
				var set_stmt = e.setStmt();
				if (set_stmt != null && set_stmt.identifier().Length > 0) {
					var idens = set_stmt.identifier();
					GetFilenumber(idens[1]);
					var iden = idens[0];
					if (set_stmt.SET() != null) {
						var sym = set_stmt.SET().Symbol;
						var lineIndex = sym.Line - 1;
						var sym_s = sym.Column;
						var sym_e = sym_s + sym.Text.Length;
						rewriteVBA.AddChange(lineIndex, (sym_s, sym_e), new string(' ', sym_e - sym_s), sym_e);
					}
					if (name.GetText() != iden.GetText()) {
						continue;
					}
					var s = iden.Start;
					rewriteVBA.AddChange(
						s.Line - 1, (s.Column, s.Column),
						"Get", s.Column);
				}
			}

			var name_s = name.Start;
			rewriteVBA.AddChange(name_s.Line - 1, (0, name_s.Column), 
				"Private Function Get", name_s.Column);

			var end_s = end_stm.Start;
			rewriteVBA.AddChange(end_s.Line - 1, "End Function");
		}

		public override void ExitPropertyLetStmt([NotNull] VBAParser.PropertyLetStmtContext context) {
			var setTokens = context.GetTokens(VBAParser.SET);
			var letTokens = context.GetTokens(VBAParser.LET);
			var tokens = setTokens.Concat(letTokens);
			GetLetSet(tokens);
			GetVBAFunction(context.children);
			GetPredefined(context.children);
			GetFilenumber(context);
			GetVariant(context);

			var name = context.identifier();
			var end_stm = context.endPropertyStmt();

			string asType = null;
			var args = context.argList();
			if (args != null && args.arg().Length > 0) {
				var arg = args.arg()[0];
				var argIdent = arg.asTypeClause()?.identifier();
				if (GetVariant(argIdent)) {
					asType = "Object";
				} else {
					asType = argIdent?.GetText();
				}
			}
			rewriteVBA.AddPropertyName(name.Start.Line - 1, name.GetText(), asType);

			var name_s = name.Start;
			rewriteVBA.AddChange(name_s.Line - 1, (0, name_s.Column),
				"Private Sub Set", name_s.Column);
			var end_s = end_stm.Start;
			rewriteVBA.AddChange(end_s.Line - 1, "End Sub");
		}

		public override void ExitPropertySetStmt([NotNull] VBAParser.PropertySetStmtContext context) {
			var setTokens = context.GetTokens(VBAParser.SET);
			var letTokens = context.GetTokens(VBAParser.LET);
			var tokens = setTokens.Concat(letTokens);
			GetLetSet(tokens);
			GetVBAFunction(context.children);
			GetPredefined(context.children);
			GetFilenumber(context);
			GetVariant(context);

			var name = context.identifier();
			var end_stm = context.endPropertyStmt();

			string asType = null;
			var args = context.argList();
			if (args != null && args.arg().Length > 0) {
				var arg = args.arg()[0];
				var argIdent = arg.asTypeClause()?.identifier();
				GetVariant(argIdent);
				asType = argIdent?.GetText();
			}
			rewriteVBA.AddPropertyName(name.Start.Line - 1, name.GetText(), asType);
			
			var name_s = name.Start;
			rewriteVBA.AddChange(name_s.Line - 1, (0, name_s.Column),
				"Private Sub Set", name_s.Column);
			var end_s = end_stm.Start;
			rewriteVBA.AddChange(end_s.Line - 1, "End Sub");
		}

		public override void ExitOpenStmt([NotNull] VBAParser.OpenStmtContext context) {
			base.ExitOpenStmt(context);
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
			var traget_list = new List<string> { 
				"range", "cells", "workbooks", "worksheets" };
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
