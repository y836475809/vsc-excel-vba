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
	internal enum PropertyType {
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

	internal class ChangeVBAProperty {
		public List<ChangeData> ChangeDataList { get; set; }
		public List<PropertyMember> PropertyMembers { get; set; }
		public List<PropertyDiagnostic> IgnorePropDiags { get; set; }

		private List<PropertyData> PropDataList;

		public ChangeVBAProperty() {
			ChangeDataList = [];
			PropertyMembers = [];
			IgnorePropDiags = [];
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

		public void Change() {
			foreach (var propData in PropDataList) {
				if (propData.Name == null) {
					continue;
				}
				var dataType = propData.DataType();
				if (dataType == PropertyDataType.GetSet) {
					RewriteGetSetProp(propData);
				}
				if (dataType == PropertyDataType.Get) {
					RewriteGetProp(propData);
				}
				if (dataType == PropertyDataType.Set) {
					RewriteSetProp(propData);
				}
			}
		}

		private void RewriteGetSetProp(PropertyData propertyData) {
			var getPropStmt = propertyData.GetStmt;

			{
				var getSym = getPropStmt.GET().Symbol;
				var rangeStartCol = getSym.Column;
				var rangeEndCol = rangeStartCol + getSym.Text.Length + 1;
				var startCol = getPropStmt.identifier().Start.Column;
				ChangeDataList.Add(new (getSym.Line - 1,
					(rangeStartCol, rangeEndCol), "", startCol));
			}

			RewriteVariant(getPropStmt);

			{
				var startLine = getPropStmt.Start.Line - 1;
				var rangeCol = getPropStmt.Start.Column + getPropStmt.GetText().Length;
				var startCol = rangeCol;
				ChangeDataList.Add(new(startLine,
					(rangeCol, rangeCol),
					" : Set : End Set : Get", startCol, false));

				AddIgnoreDiagnostic("Set", startLine);
				AddIgnoreDiagnostic("Get", startLine);

				var getPropEndStmt = propertyData.GetEndStmt;
				ChangeDataList.Add(new(getPropEndStmt.Start.Line - 1,
					"End Get : End Property"));
			}

			RewriteSet(propertyData.SetStmt, propertyData.SetEndStmt);
		}

		private void RewriteGetProp(PropertyData propertyData) {
			var getPropStmt = propertyData.GetStmt;
			var propStartLine = getPropStmt.Start.Line;

			{ 
				var rangeStartCol = getPropStmt.PROPERTY().Symbol.Column;
				var getSym = getPropStmt.GET().Symbol;
				var rangeEndCol = getSym.Column + getSym.Text.Length;
				var startCol = getPropStmt.identifier().Start.Column;
				ChangeDataList.Add(new(propStartLine - 1,
					(rangeStartCol, rangeEndCol),
					"ReadOnly Property", startCol));
			}

			RewriteVariant(getPropStmt);

			{
				var startCol = getPropStmt.Start.Column;
				var rangeCol = startCol + getPropStmt.GetText().Length;
				ChangeDataList.Add(new(propStartLine - 1,
					(rangeCol, rangeCol), " : Get", rangeCol, false));
				AddIgnoreDiagnostic("Get", propStartLine - 1);
			}

			var propEndStmt = propertyData.GetEndStmt;
			ChangeDataList.Add(new(propEndStmt.Start.Line - 1,
				"End Get : End Property"));
		}

		private void RewriteSetProp(PropertyData propertyData) {
			var setPropStmt = propertyData.SetStmt;
			var startLine = setPropStmt.Start.Line;

			{
				RewriteSet(propertyData.SetStmt, propertyData.SetEndStmt);

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
				PropertyMembers.Add(new (
					$"{porpVisibility}WriteOnly Property {propName}{propAsType}{propDim}",
					startLine - 1
				));
			}
		}

		private void RewriteSet( 
			PropertySetStmtContext setPropStmt,
			EndPropertyStmtContext setPropEndStmt) {
			var rangeStartCol = setPropStmt.Start.Column;
			var rangeEndCol = setPropStmt.identifier().Start.Column;
			ChangeDataList.Add(new(setPropStmt.Start.Line - 1,
				(rangeStartCol, rangeEndCol),
				"Private Sub R__", rangeEndCol));

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
					ChangeDataList.Add(new(asTypeClause.Start.Line - 1,
						(startCol, startCol + asType.Length),
						"Object ", startCol, false));
				}
			}

			ChangeDataList.Add(new(
				setPropEndStmt.Start.Line - 1, "End Sub"));
		}

		private void RewriteVariant(PropertyGetStmtContext context) {
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
			ChangeDataList.Add(new(start.Line - 1,
				(startCol, startCol + text.Length),
				"Object ", startCol, false));
		}

		private void AddIgnoreDiagnostic(string code, int line) {
			IgnorePropDiags.Add(new(
				"BC32009", code, "Error", line
			));		
		}
	}
}
