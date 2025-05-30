using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

			{
				var getSym = getPropStmt.GET().Symbol;
				var rangeStartCol = getSym.Column;
				var rangeEndCol = rangeStartCol + getSym.Text.Length + 1;
				var startCol = getPropStmt.identifier().Start.Column;
				rewriteVBA.AddChange(getSym.Line - 1,
					(rangeStartCol, rangeEndCol), "", startCol);
			}
			{
				var rangeCol = getPropStmt.Start.Column + getPropStmt.GetText().Length;
				var startCol = rangeCol;
				rewriteVBA.AddChange(getPropStmt.Start.Line - 1,
					(rangeCol, rangeCol), 
					" : Set : End Set : Get", startCol, false);

				var getPropEndStmt = propertyData.GetEndStmt;
				rewriteVBA.AddChange(getPropEndStmt.Start.Line - 1,
					"End Get : End Property");
			}
			{
				var setPropStmt = propertyData.SetStmt;
				var rangeStartCol = setPropStmt.Start.Column;
				var rangeEndCol = setPropStmt.identifier().Start.Column;
				rewriteVBA.AddChange(setPropStmt.Start.Line - 1,
					(rangeStartCol, rangeEndCol),
					"Private Sub set_p_", rangeEndCol);

				var argList = setPropStmt.argList();
				foreach (var arg in argList.arg()) {
					var asTypeClause = arg.asTypeClause();
					if (asTypeClause == null) {
						continue;
					}
					var asTypeIdent = asTypeClause.identifier();
					var asType = asTypeIdent.GetText();
					if (Util.Eq(asType, "variant")) {
						var startCol = asTypeIdent.Start.Column;
						rewriteVBA.AddChange(asTypeClause.Start.Line - 1,
							(startCol, startCol + asType.Length),
							"Object ", startCol, false);
					}
				}

				var setPropEndStmt = propertyData.SetEndStmt;
				rewriteVBA.AddChange(
					setPropEndStmt.Start.Line - 1, "End Sub");
			}
		}

		private void RewriteGetProp(IRewriteVBA rewriteVBA, PropertyData propertyData) {
			var getPropStmt = propertyData.GetStmt;
			var propStartLine = getPropStmt.Start.Line;

			{ 
				var rangeStartCol = getPropStmt.PROPERTY().Symbol.Column;
				var getSym = getPropStmt.GET().Symbol;
				var rangeEndCol = getSym.Column + getSym.Text.Length;
				var startCol = getPropStmt.identifier().Start.Column;
				rewriteVBA.AddChange(propStartLine - 1,
					(rangeStartCol, rangeEndCol),
					"ReadOnly Property", startCol);
			}
			{
				var startCol = getPropStmt.Start.Column;
				var rangeCol = startCol + getPropStmt.GetText().Length;
				rewriteVBA.AddChange(propStartLine - 1,
					(rangeCol, rangeCol), " : Get", rangeCol, false);
			}

			var propEndStmt = propertyData.GetEndStmt;
			rewriteVBA.AddChange(propEndStmt.Start.Line - 1, 
				"End Get : End Property");
		}

		private void RewriteSetProp(IRewriteVBA rewriteVBA, PropertyData propertyData) {
			var setPropStmt = propertyData.SetStmt;
			var startLine = setPropStmt.Start.Line;
			
			{	
				var setSym = setPropStmt.SET();
				if (setSym == null) {
					setSym = setPropStmt.LET();
				}
				var rangeStartCol = setPropStmt.PROPERTY().Symbol.Column;
				var rangeEndCol = setSym.Symbol.Column + setSym.Symbol.Text.Length;
				var startCol = setPropStmt.identifier().Start.Column;
				rewriteVBA.AddChange(startLine - 1,
					(rangeStartCol, rangeEndCol),
					"WriteOnly Property", startCol);
			}
			{
				var argList = setPropStmt.argList();
				var argStartCol = argList.Start.Column;
				var argName = "";
				var asType = "";
				if (argList.arg().Any()) {
					var arg = argList.arg().First();
					argStartCol = arg.Start.Column;
					argName = arg.identifier().GetText();
					if (arg.asTypeClause() != null) {
						asType = arg.asTypeClause().identifier().GetText();
						if (Util.Eq(asType, "variant")) {
							asType = "Object ";
						}
					}
				}
				var argText = argList.GetText().Replace("variant", "Object ", StringComparison.OrdinalIgnoreCase);
				var insertText1 = ") : Set";
				var insertText2 = $"{argText}";
				if (asType != "") {
					insertText1 = $") As {asType} : Set";
				}
				var lCol = argList.LPAREN().Symbol.Column;
				var startCol = setPropStmt.Start.Column;
				var propLen = setPropStmt.GetText().Length;
				// rep1に "("の分+1する
				var shiftCol = insertText1.Length + 1;
				rewriteVBA.AddChange(startLine - 1,
					(lCol+1, startCol + propLen),
					$"{insertText1}{insertText2}",
					argStartCol, shiftCol);
			}

			var setPropEndStmt = propertyData.SetEndStmt;
			rewriteVBA.AddChange(setPropEndStmt.Start.Line - 1, "End Set : End Property");
		}
	}
}
