using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static VBAAntlr.VBAParser;

namespace VBARewrite {
	internal class ChangeVBAType {
		public List<ChangeData> ChangeDataList { get; set; }

		public ChangeVBAType() {
			ChangeDataList = [];
		}

		public void ChangeTypeStmt(TypeStmtContext context) {
			ChangeTypeMember(context);

			var typeStmt = context.TYPE();
			var sym = typeStmt.Symbol;

			var name = context.identifier();
			var st = name.Start;
			var s = sym.Column;
			var e = name.Start.Column;
			var ws_count = name.Start.Column - (s + sym.Text.Length);
			var text = $"Structure{new string(' ', ws_count)}";
			ChangeDataList.Add(new(st.Line - 1, (s, e), text, st.Column));

			var end_stm = context.typeEndStmt();
			var end_s = end_stm.Start;
			ChangeDataList.Add(new(end_s.Line - 1, "End Structure"));
		}

		private void ChangeTypeMember(TypeStmtContext context) {
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
					ChangeDataList.Add(new(st.Line - 1, (s, s), "Public ", s));
				}

				var asType = item.asTypeClause()?.identifier();
				if (asType == null) {
					continue;
				}
				var asTypeName = asType.GetText();
				if (Util.Eq(asTypeName, "variant")) {
					var st = asType.Start;
					var s = st.Column;
					var e = s + asTypeName.Length;
					ChangeDataList.Add(new(st.Line - 1, (s, e), "Object ", s, false));
				}
			}
		}
	}
}
