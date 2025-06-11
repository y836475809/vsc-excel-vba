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
using static VBAAntlr.VBAParser;

namespace VBARewrite {
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

	internal class RewriteProperty {
		private List<PropertyData> PropDataList;

		public RewriteProperty() {
			PropDataList = [];
		}

		public void AddProperty(PropertyType propType, ParserRuleContext stmt) {
			if (propType == PropertyType.End) {
				if (PropDataList.Count == 0) {
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

			RewriteVariant(rewriteVBA, getPropStmt);

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

			RewriteSet(rewriteVBA, 
				propertyData.SetStmt, propertyData.SetEndStmt);
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

			RewriteVariant(rewriteVBA, getPropStmt);

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
				RewriteSet(rewriteVBA,
					propertyData.SetStmt, propertyData.SetEndStmt);

				var propAsType = "";
				var argList = setPropStmt.argList();
				if (argList.arg().Length != 0) {
					var asType = argList.arg().First().asTypeClause()?.identifier()?.GetText();
					if(asType != null) {
						if (Util.Eq(asType, "variant")) {
							asType = "Object ";
						}
						propAsType = $" As {asType}";
					}
				}

				var porpVisibility = "";
				var visibility = setPropStmt.visibility();
				if (visibility != null) {
					if (visibility.PUBLIC() != null) {
						porpVisibility = "Public ";
					} else if (visibility.PRIVATE() != null) {
						porpVisibility = "Private ";
					}
				}

				var propDim = "";
				if (argList.arg().Length != 0) {
					var dimIdents = argList.arg().First().arrayStmt()?.GetText();
					if (dimIdents != null) {
						propDim = dimIdents;
					}
				}

				var propName = setPropStmt.identifier().GetText();
				rewriteVBA.AddPropertyMember(
					$"{porpVisibility}WriteOnly Property {propName}{propAsType}{propDim}",
					startLine - 1);
			}
		}

		private void RewriteSet(IRewriteVBA rewriteVBA, 
			PropertySetStmtContext setPropStmt,
			EndPropertyStmtContext setPropEndStmt) {
			var rangeStartCol = setPropStmt.Start.Column;
			var rangeEndCol = setPropStmt.identifier().Start.Column;
			rewriteVBA.AddChange(setPropStmt.Start.Line - 1,
				(rangeStartCol, rangeEndCol),
				"Private Sub R__", rangeEndCol);

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

			rewriteVBA.AddChange(
				setPropEndStmt.Start.Line - 1, "End Sub");
		}

		private void RewriteVariant(IRewriteVBA rewriteVBA, PropertyGetStmtContext context) {
			var asType = context.asTypeClause()?.identifier();
			if (asType == null) {
				return;
			}
			var text = asType.GetText();
			if (!Util.Eq(text, "variant")) {
				return;
			}
			var start = asType.Start;
			var startCol = start.Column;
			rewriteVBA.AddChange(start.Line - 1,
				(startCol, startCol + text.Length),
				"Object ", startCol, false);
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
