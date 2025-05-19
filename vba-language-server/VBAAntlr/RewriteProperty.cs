using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using VBAAntlr;
using static VBAAntlr.VBAParser;

namespace AntlrTemplate {
	public enum PropertyType {
		End,
		Get,
		Set
	}

	internal enum PropertyDataType {
		None,
		GetSet,
		Get,
		Set
	}
	internal class PropertyData {
		public string Name { get; set; }
		public PropertyGetStmtContext GetStmt { get; set; }
		public EndPropertyStmtContext GetEndStmt { get; set; }

		public PropertySetStmtContext SetStmt { get; set; }
		public EndPropertyStmtContext SetEndStmt { get; set; }


		public PropertyDataType DataType() {
			var hasGetPorp = GetStmt != null && GetEndStmt != null;
			var hasSetPorp = SetStmt != null && SetEndStmt != null;
			if (hasGetPorp && hasSetPorp) {
				return PropertyDataType.GetSet;
			}
			if (hasSetPorp) {
				return PropertyDataType.Set;
			}
			return PropertyDataType.None;
		}
	}


	internal class RewriteGetProperty {
		private List<(PropertyType, ParserRuleContext)> PropLines;
		//private List<(PropertyType, string, int)> PropBlockLines;

		public RewriteGetProperty() {
			PropLines = [];
			//PropBlockLines = [];
		}
		public void AddProperty(PropertyType propType, ParserRuleContext stmt) {
			//PropBlockLines.Add((propType, name, line));
			PropLines.Add((propType, stmt));
		}

		//public void AddPropertyEnd(ParserRuleContext stmt) {
		//	PropLines.Add((PropertyType.End, stmt));
		//	//PropLines.Add((propType, stmt, line));
		//}

		public void Rewrite(IRewriteVBA rewriteVBA) {
			var propEndList = new List<(PropertyType, int)>();
			var propDict = new Dictionary<string, PropertyData>();
			for (int i = 0; i < PropLines.Count; i += 2) {
				var propType = PropLines[i].Item1;

				if (propType == PropertyType.End) {
					continue;
				}
				if (PropLines.Count <= i + 1) {
					break;
				}

				var startStmt = PropLines[i].Item2;
				var EndStmt = PropLines[i + 1].Item2;

				if (propType == PropertyType.Get) {
					var getStmt = startStmt as PropertyGetStmtContext;
					var name = getStmt.identifier().GetText();
					if (!propDict.ContainsKey(name)) {
						propDict[name] = new();
					}
					propDict[name].GetStmt = getStmt;
					propDict[name].GetEndStmt = EndStmt as EndPropertyStmtContext;
				}
				if (propType == PropertyType.Set) {
					var setStmt = startStmt as PropertySetStmtContext;
					var name = setStmt.identifier().GetText();
					if (!propDict.ContainsKey(name)) {
						propDict[name] = new();
					}
					propDict[name].SetStmt = setStmt;
					propDict[name].SetEndStmt = EndStmt as EndPropertyStmtContext;
				}
			}
			foreach (var (name, propData) in propDict) {
				if(propData.DataType() == PropertyDataType.GetSet) {

				}
				if (propData.DataType() == PropertyDataType.Get) {

				}
				if (propData.DataType() == PropertyDataType.Set) {

				}
			}

			//foreach (var propEnd in propEndList) {
			//	var (pre, line) = propEnd;
			//	if (pre == PropertyType.Get) {
			//		rewriteVBA.AddChange(line - 1, "End Function");
			//	} else {
			//		rewriteVBA.AddChange(line - 1, "End Sub");
			//	}
			//}

			//foreach (var token in letsetTokens) {
			//	if (GetPropertyToken(token, rangeDict, out string propName)) {
			//		var sym = token.Symbol;
			//		var line = sym.Line - 1;
			//		var sc = sym.Column;
			//		var ec = sc + sym.Text.Length;
			//		rewriteVBA.AddChange(line, (sc, sc), "Get", sc);
			//	}
			//}
		}

		private void RewriteGetSet(
			IRewriteVBA rewriteVBA,
			PropertyGetStmtContext getStmt, EndPropertyStmtContext getEndStmt,
			PropertySetStmtContext setStmt, EndPropertyStmtContext setEndStmt) {
			//getStmt.

		}

		private bool GetPropertyToken(Antlr4.Runtime.Tree.ITerminalNode token, Dictionary<string, (int, int)> rangeDict, out string name) {
			foreach (var item in rangeDict) {
				var s = item.Value.Item1;
				var e = item.Value.Item2;
				//if (Enumerable.Range(s, e).Contains(token.Symbol.Line)) {
				if (s < token.Symbol.Line && token.Symbol.Line < e) {
					name = item.Key;
					return true;
				}
			}
			name = "";
			return false;
		}
	}
}
