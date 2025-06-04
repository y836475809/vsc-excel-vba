using Antlr4.Runtime;
using AntlrTemplate;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static VBAAntlr.VBAParser;

namespace VBAAntlr {
	internal class DynamicArray {
		public string Name { get; set; }
		public DimStmtContext DimStmt;
		public List<RedimStmtContext> ReDimStmts;

		public DynamicArray() {
			DimStmt = null;
			ReDimStmts = [];
		}
	}

	internal class RewriteDynamicArray {
		private Dictionary<string, DimStmtContext> FieldDimDict;
		private Dictionary<string, DimStmtContext> DimDict;
		private Dictionary<string, List<RedimStmtContext>> ReDimStmtsDict;
		private List<Dictionary<string, DynamicArray>> DynaArrayDictList;

		public RewriteDynamicArray() {
			FieldDimDict = [];
			DimDict = [];
			ReDimStmtsDict = [];
			DynaArrayDictList = [];
		}

		public void Add(DimStmtContext stmt) {
			var name = stmt.identifier().GetText();
			DimDict[name] = stmt;
		}

		public void Add(RedimStmtContext stmt) {
			var name = stmt.identifier().GetText();
			if (!ReDimStmtsDict.ContainsKey(name)) {
				ReDimStmtsDict[name] = [];
			}
			ReDimStmtsDict[name].Add(stmt);
		}

		public void StartBlockStmt() {
			foreach (var (dimName, stmt) in DimDict) {
				FieldDimDict[dimName] = stmt;
			}
			DimDict.Clear();
			ReDimStmtsDict.Clear();
		}

		public void EndBlockStmt() {
			Dictionary<string, DynamicArray> DynaArrayDict = [];
			foreach (var (dimName, stmt) in DimDict) {
				if (!DynaArrayDict.ContainsKey(dimName)) {
					DynaArrayDict[dimName] = new();
				}
				var da = DynaArrayDict[dimName];
				da.DimStmt = stmt;
			}
			foreach (var (reDimName, stmt) in ReDimStmtsDict) {
				if (!DynaArrayDict.ContainsKey(reDimName)) {
					DynaArrayDict[reDimName] = new();
				}
				var da = DynaArrayDict[reDimName];
				da.ReDimStmts = stmt;
			}
			DynaArrayDictList.Add(DynaArrayDict);
			DimDict.Clear();
			ReDimStmtsDict.Clear();
		}

		public void Rewrite(IRewriteVBA rewriteVBA) {
			foreach (var (dimName, stmt) in FieldDimDict) {
				foreach (var DynaArrayDict in DynaArrayDictList) {
					if (DynaArrayDict.TryGetValue(dimName, out DynamicArray da)) {
						if(da.DimStmt == null) {
							da.DimStmt = stmt;
						}
					}
				}
			}
			foreach (var DynaArrayDict in DynaArrayDictList) {
				foreach (var (varName, da) in DynaArrayDict) {
					var dimStmt = da.DimStmt;
					var reDimStmts = da.ReDimStmts;
					if (dimStmt == null && reDimStmts.Count > 0) {
						// redim a(2) -> dim a():redim a(2) 
						// redim a(2) As Long -> dim a() As Long:redim a(2) 
						// redim a(2, 2) -> dim a(,):redim a(2,2) 
						var reDimStmt = reDimStmts[0];
						var redimArgs = reDimStmt.redimArgList()?.identifier();
						var redimToArgs = reDimStmt.redimToArgList()?.redimToArg();
						var c = "";
						if (redimArgs != null && redimArgs.Length > 1) {
							c = new string(',', redimArgs.Length - 1);
						}
						if (redimToArgs != null && redimToArgs.Length > 1) {
							c = new string(',', redimToArgs.Length - 1);
						}
						var asTypeClause = "";
						var asType = reDimStmt.asTypeClause()?.GetText();
						if (asType != null) {
							asTypeClause = $" {asType}";
						}
						rewriteVBA.AddChange(
							reDimStmt.Start.Line - 1,
							(reDimStmt.Start.Column, reDimStmt.Start.Column),
							$"Dim {varName}({c}){asTypeClause}:",
							reDimStmt.Start.Column);
						if (asTypeClause != "") {
							var asStmt = reDimStmt.asTypeClause();
							var asText = asStmt.GetText();
							var sc = asStmt.Start.Column;
							var ec = sc + asText.Length;
							rewriteVBA.AddChange(
								asStmt.Start.Line - 1,
								(sc, ec), new string(' ', asText.Length), sc);
						}
					}
					if (dimStmt != null && reDimStmts.Count > 0) {
						// dim a() redim a(2, 2) -> dim a(,):redim a(2, 2) 
						var reDimStmt = reDimStmts[0];
						var redimArgs = reDimStmt.redimArgList()?.identifier();
						var redimToArgs = reDimStmt.redimToArgList()?.redimToArg();
						var c = "";
						if (redimArgs != null && redimArgs.Length > 1) {
							c = new string(',', redimArgs.Length - 1);
						}
						if (redimToArgs != null && redimToArgs.Length > 1) {
							c = new string(',', redimToArgs.Length - 1);
						}
						var sc = dimStmt.LPAREN().Symbol.Column + 1;
						rewriteVBA.AddChange(
							dimStmt.Start.Line - 1,
							(sc, sc), $"{c}", sc);
					}
				}
			}
		}
	}
}
