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
	internal class PropertyDiagnostic : IPropertyDiagnostic {
		public string Id { get; set; }

		public string Code { get; set; }

		public string Severity { get; set; }

		public int Line { get; set; }
	}

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
		private List<PropertyData> PropDataList;

		public RewriteGetProperty() {
			PropDataList = [];
		}

		public void AddProperty(PropertyType propType, ParserRuleContext stmt) {
			if (propType == PropertyType.End) {
				if (!PropDataList.Any()) {
					return;
				}
				var porpData = PropDataList.Last();
				var propStmt = stmt as EndPropertyStmtContext;
				if (porpData.GetStmt != null && porpData.GetEndStmt == null) {
					porpData.GetEndStmt = propStmt;
				}
				if (porpData.SetStmt != null && porpData.SetEndStmt == null) {
					porpData.SetEndStmt = propStmt;
				}
			} else if(propType == PropertyType.Get) {
				var propStmt = stmt as PropertyGetStmtContext;
				var name = propStmt.identifier().GetText();
				var propData = PropDataList.Find(x => x.Name == name);
				if (propData == null) {
					PropDataList.Add(new PropertyData {
						Name = name,
						GetStmt = propStmt
					});
				} else {
					propData.GetStmt = propStmt;
				}
			} else if (propType == PropertyType.Set) {
				var propStmt = stmt as PropertySetStmtContext;
				var name = propStmt.identifier().GetText();
				var propData = PropDataList.Find(x => x.Name == name);
				if (propData == null) {
					PropDataList.Add(new PropertyData {
						Name = propStmt.identifier().GetText(),
						SetStmt = propStmt
					});
				} else {
					propData.SetStmt = propStmt;
				}
			}
		}

		public void Rewrite(IRewriteVBA rewriteVBA) {
			foreach (var propData in PropDataList) {
				if (propData.Name == null) {
					continue;
				}
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
				var startLine = getPropStmt.Start.Line - 1;
				var rangeCol = getPropStmt.Start.Column + getPropStmt.GetText().Length;
				var startCol = rangeCol;
				rewriteVBA.AddChange(startLine,
					(rangeCol, rangeCol), 
					" : Set : End Set : Get", startCol, false);

				AddIgnoreDiagnostic(rewriteVBA, "Set", startLine);
				AddIgnoreDiagnostic(rewriteVBA, "Get", startLine);

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

				AddIgnoreDiagnostic(rewriteVBA, "Get", propStartLine - 1);
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

				AddIgnoreDiagnostic(rewriteVBA, "Set", startLine - 1);
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

		private void AddIgnoreDiagnostic(IRewriteVBA rewriteVBA, string code, int line) {
			rewriteVBA.AddIgnoreDiagnostic(new PropertyDiagnostic {
				Id = "BC32009",
				Code = code,
				Severity = "Error",
				Line = line
			});
		}
	}
}
