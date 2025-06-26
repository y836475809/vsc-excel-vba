using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using VBAAntlr;


namespace VBARewrite {
	using static VBAAntlr.VBAParser;

	internal class ChangeVBA {
		public List<ChangeData> ChangeDataList { get; set; }
		private Settings settings;

		public ChangeVBA() {
			ChangeDataList = [];
			settings = new();
		}

		public void Change(StartRuleContext context) {
			var setTokens = context.GetTokens(SET);
			var letTokens = context.GetTokens(LET);
			var tokens = setTokens.Concat(letTokens);
			GetLetSet(tokens);
			GetVBAFunction(context.children);
			GetPredefined(context.children);
			GetFilenumber(context);
			GetVariant(context);
		}

		private void GetLetSet(IEnumerable<ITerminalNode> tokens) {
			foreach (var item in tokens) {
				var sym = item.Symbol;
				var lineIndex = sym.Line - 1;
				var s = sym.Column;
				var e = s + sym.Text.Length;
				ChangeDataList.Add(new(lineIndex, (s, e), new string(' ', e - s), e));
			}
		}

		private void GetVBAFunction(IList<IParseTree> children) {
			var nl = Environment.NewLine;
			var traget_list = settings.VBAFunction.Targets;
			var funcModule = $"{settings.VBAFunction.Module}.";

			var traget_set = new List<string> { "as", "new" };
			var vbaFuncList = children.Where((x, index) => {
				if (!Util.Contains(x.GetText(), traget_list)) {
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
				if (pre_text1 == ".") {
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
				ChangeDataList.Add(new(ct.Line - 1, (s, s), funcModule, s));
			}
		}

		private void GetPredefined(IList<IParseTree> children) {
			var traget_list = settings.VBAPredefined.Targets;
			var funcModule = $"{settings.VBAPredefined.Module}.";
			var predefinedList = children.Where(x => Util.Contains(x.GetText(), traget_list));
			foreach (var item in predefinedList) {
				var ct = item.Payload as CommonToken;
				var s = ct.Column;
				ChangeDataList.Add(new(ct.Line - 1, (s, s), funcModule, s));
			}
			// TODO
			var lineInputItems = children.Where(
				x => (x.Payload as CommonToken)?.Type == VBAParser.LINE_INPUT);
			foreach (var item in lineInputItems) {
				var ct = item.Payload as CommonToken;
				var text = ct.Text;
				var s = ct.Column;
				var e = ct.Column + text.Length;
				var len = text.Length - "line_input".Length;
				var repText = $"line_input{new string(' ', len)}";
				ChangeDataList.Add(new(ct.Line - 1, (s, e), repText, s, false));
			}
		}

		private void GetFilenumber(ParserRuleContext context) {
			var fnTokens = context.children.Select((x, Index) => (Index, x))
				.Where(x => x.x.Payload is CommonToken { Type: IDENTIFIER })
				.Where(x => x.x.GetText().StartsWith('#'));
			foreach (var item in fnTokens) {
				var ni = item.Index + 1;
				var neide = context.GetChild(ni);
				var next_ct = neide?.Payload as CommonToken;
				if (next_ct?.Type != IDENTIFIER) {
					var ct = item.x.Payload as CommonToken;
					var s = ct.Column;
					ChangeDataList.Add(new(ct.Line - 1, (s, s + 1), " ", s + 1));
				}
			}
		}

		private void GetVariant(ParserRuleContext context) {
			var fnTokens = context.children
				.Where(x => x.Payload is CommonToken { Type: IDENTIFIER })
				.Where(x => Util.Eq(x.GetText(), "variant"));
			foreach (var item in fnTokens) {
				var ct = item.Payload as CommonToken;
				var s = ct.Column;
				var e = s + ct.Text.Length;
				ChangeDataList.Add(new(ct.Line - 1, (s, e), "Object ", s, false));
			}
		}
	}
}
