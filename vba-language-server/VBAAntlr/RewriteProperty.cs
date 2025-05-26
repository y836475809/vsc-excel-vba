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
				(getStartCol, getStartCol + getSym.Text.Length),
				"", getPropStmt.identifier().Start.Column);
			rewriteVBA.InsertLines(getPropStmt.Start.Line, ["Set : End Set", "Get"]);
			rewriteVBA.AddChange(getPropEndStmt.Start.Line - 1, "End Get : End Property");
			
			var setPropStmt = propertyData.SetStmt;
			var setPropEndStmt = propertyData.SetEndStmt;
			var setPropIdeStart = setPropStmt.identifier().Start;
			rewriteVBA.AddChange(setPropIdeStart.Line - 1, (0, setPropIdeStart.Column),
				"Private Sub set_p_", setPropIdeStart.Column);
			rewriteVBA.AddChange(
				setPropEndStmt.Start.Line - 1, "End Sub");
		}

		private void RewriteGetProp(IRewriteVBA rewriteVBA, PropertyData propertyData) {
			var getPropStmt = propertyData.GetStmt;
			var getPropEndStmt = propertyData.GetEndStmt;
			var porpSym = getPropStmt.PROPERTY().Symbol;
			var getSym = getPropStmt.GET().Symbol;
			var getStartCol = getSym.Column;
			var identStartCol = getPropStmt.identifier().Start.Column;
			rewriteVBA.AddChange(porpSym.Line - 1,
				(porpSym.Column, getStartCol + getSym.Text.Length),
				"ReadOnly Property", identStartCol);
			rewriteVBA.InsertLines(getPropStmt.Start.Line, ["Get"]);
			rewriteVBA.AddChange(getPropEndStmt.Start.Line - 1, "End Get : End Property");
		}

		private void RewriteSetProp(IRewriteVBA rewriteVBA, PropertyData propertyData) {
			var setPropStmt = propertyData.SetStmt;
			var setPropEndStmt = propertyData.SetEndStmt;
			var setPropIdeStart = setPropStmt.identifier().Start;
			rewriteVBA.AddChange(setPropIdeStart.Line - 1, (0, setPropIdeStart.Column),
				"Private Sub set_", setPropIdeStart.Column);
			rewriteVBA.AddChange(
				setPropEndStmt.Start.Line - 1, "End Sub");
			var asType = "";
			if (setPropStmt.arg().asTypeClause() != null) {
				var asStmt = setPropStmt.arg().asTypeClause();
				var asTypeName = asStmt.identifier().GetText();
				if(Util.Eq(asTypeName, "variant")) {
					asTypeName = "Object";
				}
				asType = $" As {asTypeName}";
			}
			rewriteVBA.InsertLines(
				setPropEndStmt.Start.Line,
				[$"Public Property {setPropStmt.identifier().GetText()}{asType}"]);
			rewriteVBA.AddLineMap(
				setPropEndStmt.Start.Line,
				 setPropIdeStart.Line - 1);
		}
	}
}
