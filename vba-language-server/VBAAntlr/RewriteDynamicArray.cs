using Antlr4.Runtime;
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
		private List<(string, int)> MethodLines;
		private List<DimStmtContext> DimStmts;
		private List<RedimStmtContext> ReDimStmts;

		public RewriteDynamicArray() {
			MethodLines = [];
			DimStmts = [];
			ReDimStmts = [];
		}

		public void Add(DimStmtContext stmt) {
			DimStmts.Add(stmt);
		}

		public void Add(RedimStmtContext stmt) {
			ReDimStmts.Add(stmt);
		}

		public void AddMethodStart(string name, int line) {
			if(name == null) {
				MethodLines.Add(("", line));
			} else {
				MethodLines.Add((name, line));
			}
		}

		public void AddMethodEnd(int line) {
			MethodLines.Add(("end", line));
		}

		public Dictionary<string, DynamicArray> GetDynamicArrayDict() {
			var rangeDitc = new Dictionary<string, (int, int)>();
			for (int i = 0; i < MethodLines.Count; i+=2) {
				var name = MethodLines[i].Item1;
				if(name == "") {
					continue;
				}
				var s = MethodLines[i].Item2;
				var e = MethodLines[i+1].Item2;
				rangeDitc[name] = (s, e);
			}

			var FieldDimStmtDict = new Dictionary<string, DimStmtContext>();
			var DimStmtDict = new Dictionary<string, DimStmtContext>();
			var ReDimStmtsDict = new Dictionary<string, List<RedimStmtContext>>();
			foreach (var stmt in DimStmts) {
				if(GetStmes(stmt, rangeDitc, out string methodName)) {
					DimStmtDict[methodName] = stmt;
				} else {
					FieldDimStmtDict[stmt.identifier().GetText()] = stmt;
				}
			}
			foreach (var stmt in ReDimStmts) {
				if (GetStmes(stmt, rangeDitc, out string methodName)) {
					if (!ReDimStmtsDict.ContainsKey(methodName)) {
						ReDimStmtsDict[methodName] = [];
					}
					ReDimStmtsDict[methodName].Add(stmt);
				}
			}

			var dynamicArrayDict = new Dictionary<string, DynamicArray>();
			foreach (var methodName in rangeDitc.Keys) {
				if (DimStmtDict.TryGetValue(methodName, out DimStmtContext dimstmt)) {
					var varName = dimstmt.identifier().GetText();
					if (!dynamicArrayDict.ContainsKey(varName)) {
						dynamicArrayDict[varName] = new DynamicArray();
						dynamicArrayDict[varName].Name = varName;
					}
					dynamicArrayDict[varName].DimStmt = dimstmt;
				}
				if (ReDimStmtsDict.TryGetValue(methodName, out List<RedimStmtContext> redimstmts)) {
					foreach (var redimstmt in redimstmts) {
						var varName = redimstmt.identifier().GetText();
						if (!dynamicArrayDict.ContainsKey(varName)) {
							dynamicArrayDict[varName] = new DynamicArray();
							dynamicArrayDict[varName].Name = varName;
						}
						dynamicArrayDict[varName].ReDimStmts.Add(redimstmt);
					}
				}
			}
			foreach (var pair in dynamicArrayDict) {
				var varName = pair.Key;
				if (FieldDimStmtDict.TryGetValue(varName, out DimStmtContext stmt)) {
					if (pair.Value.DimStmt == null) {
						pair.Value.DimStmt = stmt;
					}
				}
			}
			return dynamicArrayDict;
		}

		public void Rewrite(IRewriteVBA rewriteVBA) {
			var dynamicArrayDict = GetDynamicArrayDict();
			foreach (var pair in dynamicArrayDict) {
				var varName = pair.Key;
				var dimStmt = pair.Value.DimStmt;
				var reDimStmts = pair.Value.ReDimStmts;
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

		private bool GetStmes<T>(T stmt, Dictionary<string, (int, int)> rangeDict, out string name) where T : ParserRuleContext {
			foreach (var item in rangeDict) {
				var s = item.Value.Item1;
				var e = item.Value.Item2;
				if(Enumerable.Range(s, e).Contains(stmt.Start.Line)) {
					name = item.Key;
					return true;
				}
			}
			name = "";
			return false;
		}
	}
}
