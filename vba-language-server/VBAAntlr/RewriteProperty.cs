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
			if (hasGetPorp) {
				return PropertyDataType.Get;
			}
			if (hasSetPorp) {
				return PropertyDataType.Set;
			}
			return PropertyDataType.None;
		}
	}


	internal class RewriteGetProperty {
		private List<(PropertyType, ParserRuleContext)> PropLines;

		public RewriteGetProperty() {
			PropLines = [];
		}

		public void AddProperty(PropertyType propType, ParserRuleContext stmt) {
			PropLines.Add((propType, stmt));
		}

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
			// TODO slow
			foreach (var (name, propData) in propDict) {
				var dataType = propData.DataType();
				if (dataType == PropertyDataType.GetSet) {
					RewriteGetSetProp(rewriteVBA, propData);
				}
				if (dataType == PropertyDataType.Get) {
					RewriteGetProp(rewriteVBA, propData);
				}
				if (dataType == PropertyDataType.Set) {
					RewriteSetProp(rewriteVBA, propData);
				}
			}
		}

		private void RewriteGetSetProp(IRewriteVBA rewriteVBA, PropertyData propertyData) {
			var getPropStmt = propertyData.GetStmt;
			var getPropEndStmt = propertyData.GetEndStmt;
			var getSym = getPropStmt.GET().Symbol;
			var getStartCol = getSym.Column;
			rewriteVBA.AddChange(getSym.Line - 1,
				(getStartCol, getStartCol + getSym.Text.Length + 1),
				"", getPropStmt.identifier().Start.Column);

			var setPropStmt = propertyData.SetStmt;
			var setPropEndStmt = propertyData.SetEndStmt;
			var setPropIdeStart = setPropStmt.identifier().Start;
			rewriteVBA.AddChange(setPropIdeStart.Line - 1, (0, setPropIdeStart.Column),
				"Private Sub set_p_", setPropIdeStart.Column);
			rewriteVBA.AddChange(
				setPropEndStmt.Start.Line - 1, "End Sub");

			var sCol = getPropStmt.Start.Column + getPropStmt.GetText().Length;
			rewriteVBA.AddChange(getPropStmt.Start.Line - 1, 
				(sCol, sCol), " : Set : End Set : Get", sCol, false);
			rewriteVBA.AddChange(getPropEndStmt.Start.Line - 1, 
				"End Get : End Property");
		}

		private void RewriteGetProp(IRewriteVBA rewriteVBA, PropertyData propertyData) {
			var getPropStmt = propertyData.GetStmt;
			var getPropEndStmt = propertyData.GetEndStmt;
			var porpSym = getPropStmt.PROPERTY().Symbol;
			var getSym = getPropStmt.GET().Symbol;
			var getStartCol = getSym.Column;
			var identStartCol = getPropStmt.identifier().Start.Column;
			var startCol = getPropStmt.Start.Column;
			var propLen = getPropStmt.GetText().Length;

			rewriteVBA.AddChange(porpSym.Line - 1,
				(porpSym.Column, getStartCol + getSym.Text.Length),
				"ReadOnly Property", identStartCol);
			rewriteVBA.AddChange(porpSym.Line - 1,
				(startCol + propLen, startCol + propLen),
				" : Get", startCol + propLen, false);

			rewriteVBA.AddChange(getPropEndStmt.Start.Line - 1, "End Get : End Property");
		}

		private void RewriteSetProp(IRewriteVBA rewriteVBA, PropertyData propertyData) {
			var setPropStmt = propertyData.SetStmt;
			var setPropEndStmt = propertyData.SetEndStmt;
			var startCol = setPropStmt.Start.Column;
			var propLen = setPropStmt.GetText().Length;
			var porpSym = setPropStmt.PROPERTY().Symbol;
			var identStartCol = setPropStmt.identifier().Start.Column;
			var setSym = setPropStmt.SET();
			if(setSym == null) {
				setSym = setPropStmt.LET();
			}
			var setStartCol = setSym.Symbol.Column;
			var setEndCol = setStartCol + setSym.Symbol.Text.Length;

			rewriteVBA.AddChange(porpSym.Line - 1,
				(porpSym.Column, setEndCol),
				"WriteOnly Property", identStartCol);

			var lpCol = setPropStmt.LPAREN().Symbol.Column;
			var argName = "value";
			var asType = "";
			if (setPropStmt.arg().asTypeClause() != null) {
				var argIdets = setPropStmt.arg().identifier();
				var asStmt = setPropStmt.arg().asTypeClause();
				asType = asStmt.identifier().GetText();
				if(Util.Eq(asType, "variant")) {
					asType = "Object";
				}
				if (argIdets.Length > 0) {
					argName = argIdets[0].GetText();
				}
				lpCol = argIdets[0].Start.Column;
			}
			var rep1 = "";
			var rep2 = "";
			if (asType == "") {
				rep1 = $") : Set(";
				rep2 = $")";
			} else {
				rep1 = $") As {asType} : Set(";
				rep2 = $"{argName} As {asType})";
			}

			rewriteVBA.AddChange(porpSym.Line - 1,
				(lpCol, startCol + propLen),
				$"{rep1}{rep2}",
				lpCol, rep1.Length);
			rewriteVBA.AddChange(setPropEndStmt.Start.Line - 1, "End Set : End Property");
		}
	}
}
