using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Antlr4.Runtime.Misc;
using VBAAntlr;


namespace VBARewrite {
	internal class VBAListener : VBABaseListener {
		public ChangeVBA changeVBA;
		private ChangeVBADynamicArray changeVBADynamicArray;
		private ChangeVBAProperty changeVBAProperty;
		private ChangeVBAType changeVBAType;
		private ChnageVBAHeader changeVBAHeader;
		private ChangeVBAFileIO chnageVBAFileIO;

		public VBAListener() {
			changeVBA = new();
			changeVBADynamicArray = new();
			changeVBAProperty = new();
			changeVBAType = new();
			changeVBAHeader = new();
			chnageVBAFileIO = new();
		}

		public override void ExitStartRule([NotNull] VBAParser.StartRuleContext context) {
			changeVBA.Change(context);
			changeVBAProperty.Change();
			changeVBADynamicArray.Change();
		}

		public override void ExitDimStmt([NotNull] VBAParser.DimStmtContext context) {
			changeVBADynamicArray.Add(context);
		}

		public override void ExitRedimStmt([NotNull] VBAParser.RedimStmtContext context) {
			changeVBADynamicArray.Add(context);
		}

		public override void ExitSubStmt([NotNull] VBAParser.SubStmtContext context) {
			changeVBADynamicArray.StartBlockStmt();
		}

		public override void ExitEndSubStmt([NotNull] VBAParser.EndSubStmtContext context) {
			changeVBADynamicArray.EndBlockStmt();
		}

		public override void ExitFunctionStmt([NotNull] VBAParser.FunctionStmtContext context) {
			changeVBADynamicArray.StartBlockStmt();
		}

		public override void ExitEndFunctionStmt([NotNull] VBAParser.EndFunctionStmtContext context) {
			changeVBADynamicArray.EndBlockStmt();
		}
		
		public override void ExitModuleAttributes([NotNull] VBAParser.ModuleAttributesContext context) {
			changeVBAHeader.GetModuleAttributes(context);
		}

		public override void ExitModuleOption([NotNull] VBAParser.ModuleOptionContext context) {
			changeVBAHeader.GetModuleOption(context);
		}

		public override void ExitTypeStmt([NotNull] VBAParser.TypeStmtContext context) {
			changeVBAType.ChangeTypeStmt(context);
		}

		public override void ExitPropertyGetStmt([NotNull] VBAParser.PropertyGetStmtContext context) {
			changeVBAProperty.AddProperty(PropertyType.Get, context);
			changeVBADynamicArray.StartBlockStmt();
		}
		public override void ExitPropertySetStmt([NotNull] VBAParser.PropertySetStmtContext context) {
			changeVBAProperty.AddProperty(PropertyType.Set, context);
			changeVBADynamicArray.StartBlockStmt();
		}

		public override void ExitEndPropertyStmt([NotNull] VBAParser.EndPropertyStmtContext context) {
			changeVBAProperty.AddProperty(PropertyType.End, context);
			changeVBADynamicArray.EndBlockStmt();
		}

		public override void ExitOpenStmt([NotNull] VBAParser.OpenStmtContext context) {
			chnageVBAFileIO.ChangeOpenStmt(context);
		}

		public override void ExitOutputStmt([NotNull] VBAParser.OutputStmtContext context) {
			chnageVBAFileIO.ChangeOutputStmt(context);
		}

		public override void ExitInputStmt([NotNull] VBAParser.InputStmtContext context) {
			chnageVBAFileIO.ChangeInputStmt(context);
		}

		public override void ExitLineInputStmt([NotNull] VBAParser.LineInputStmtContext context) {
			chnageVBAFileIO.ChangeLineInputStmt(context);
		}

		public override void ExitCloseStmt([NotNull] VBAParser.CloseStmtContext context) {
			chnageVBAFileIO.ChangeCloseStmt(context);
		}

		public VBCode ApplyChange(string fp, string vbaCode) {
			ChangeDataDict changeDataDict = [];
			ColumnShiftDict colShiftDict = [];
			LineMapDict lineMapDict = [];
			HashSet<int> ignoreDiagnosticLineSet = [];

			AddChange(ref changeDataDict, changeVBADynamicArray.ChangeDataList);
			AddChange(ref changeDataDict, changeVBAProperty.ChangeDataList);
			AddChange(ref changeDataDict, changeVBAType.ChangeDataList);
			AddChange(ref changeDataDict, changeVBAHeader.ChangeDataList);
			AddChange(ref changeDataDict, changeVBA.ChangeDataList);
			AddChange(ref changeDataDict, chnageVBAFileIO.ChangeDataList);

			var lines = vbaCode.Split(Environment.NewLine).ToList();

			ReverseSortChangeVBA(changeDataDict);

			foreach (var item in changeDataDict) {
				var lineIndex = item.Key;
				var line = lines[lineIndex];
				foreach (var changeData in item.Value) {
					var (colShift, repLine) = changeData.Apply(line);
					line = repLine;
					if (colShift.ShiftCol == 0) {
						continue;
					}
					if (!colShiftDict.TryGetValue(lineIndex, out List<ColumnShift> value)) {
						value = ([]);
						colShiftDict.Add(lineIndex, value);
					}
					value.Add(colShift);
				}
				lines[lineIndex] = line;
			}
			SortColumnShift(colShiftDict);

			foreach (var propMember in changeVBAProperty.PropertyMembers) {
				lines.Add(propMember.Name);
				var porpLine = lines.Count - 1;
				lineMapDict[porpLine] = propMember.Line;
				ignoreDiagnosticLineSet.Add(porpLine);
			}

			changeVBAHeader.Change(ref lines);

			var name = Path.GetFileNameWithoutExtension(fp);
			var vbCode = string.Join(Environment.NewLine, lines);
			var code = new VBCode() {
				//VBACode = vbaCode,
				Code = vbCode,
				ColShiftDict = colShiftDict,
				LineMapDict = lineMapDict,
				DiagnosticList = changeVBAHeader.GetAttributeDiagnosticList(name),
				IgnoreDiagnosticList = chnageVBAFileIO.IgnoreDiagnosticList,
				IgnorePropertyDiagnosticList = changeVBAProperty.IgnorePropDiags,
				IgnoreDiagnosticLineSet = ignoreDiagnosticLineSet
			};
			return code;
		}


		private void AddChange(ref ChangeDataDict changeDataDict, List<ChangeData> changeDataList) {
			foreach (var changeData in changeDataList) {
				var line = changeData.Line;
				if (changeDataDict.TryGetValue(line, out List<ChangeData> value)) {
					var f = value.FindIndex(x => {
						return x.Eq(changeData);
					});
					if (f < 0) {
						value.Add(changeData);
					}
				} else {
					changeDataDict[line] = [changeData];
				}
			}
		}

		private void ReverseSortChangeVBA(ChangeDataDict changeDataDict) {
			foreach (var item in changeDataDict) {
				item.Value.Sort((a, b) => -(a.StartCol - b.StartCol));
			}
		}

		private void SortColumnShift(ColumnShiftDict dict) {
			foreach (var item in dict) {
				item.Value.Sort((a, b) => a.StartCol - b.StartCol);
			}
		}
	}
}
