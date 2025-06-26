using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using VBAAntlr;


namespace VBARewrite {
	public class VBARewriter {
		protected Dictionary<string, VBCode> vbCodeDict;

		public VBARewriter() {
			vbCodeDict = [];
		}

		public int GetColShift(string name, int line, int col) {
			if (!vbCodeDict.TryGetValue(name, out VBCode code)) {
				return 0;
			}
			if (!code.ColShiftDict.TryGetValue(line, out List<ColumnShift> colShifts)) {
				return 0;
			}
			var colShift = colShifts.Where(x => x.StartCol <= col).Select(x => x.ShiftCol).Sum();
			return colShift;
		}

		public int GetReMapLineIndex(string name, int line) {
			if (!vbCodeDict.TryGetValue(name, out VBCode code)) {
				return -1;
			}
			if (!code.LineMapDict.TryGetValue(line, out int lineIndex)) {
				return -1;
			}
			return lineIndex;
		}

		public List<VBADiagnostic> GetDiagnostics(string name) {
			if (!vbCodeDict.TryGetValue(name, out VBCode code)) {
				return [];
			}
			return code.DiagnosticList;
		}

		public List<VBADiagnostic> GetIgnoreDiagnostics(string name) {
			if (!vbCodeDict.TryGetValue(name, out VBCode code)) {
				return [];
			}
			return code.IgnoreDiagnosticList;
		}

		public List<PropertyDiagnostic> GetIgnorePropertyDiagnostics(string name) {
			if (!vbCodeDict.TryGetValue(name, out VBCode code)) {
				return [];
			}
			return code.IgnorePropertyDiagnosticList;
		}

		public HashSet<int> GetIgnoreLineDiagnosticsSet(string name) {
			if (!vbCodeDict.TryGetValue(name, out VBCode code)) {
				return [];
			}
			return code.IgnoreDiagnosticLineSet;
		}

		public string Rewrite(string name, string vbaCode) {
			if (name.EndsWith(".d.vb")) {
				return vbaCode;
			}

			var lexer = new VBALexer(new AntlrInputStream(vbaCode));
			var tokens = new CommonTokenStream(lexer);
			var parser = new VBAParser(tokens);
			parser.Interpreter.PredictionMode = PredictionMode.SLL;
			lexer.RemoveErrorListeners();
			parser.RemoveErrorListeners();

			var nn = new VBAListener();
			parser.AddParseListener(nn);
			try {
				parser.startRule();
			} catch (Exception) {
				tokens.Reset();
				parser.Reset();
				parser.Interpreter.PredictionMode = PredictionMode.LL;
				parser.startRule();
			}

			var vbCode = nn.ApplyChange(name, vbaCode);
			vbCodeDict[name] = vbCode;
			return vbCode.Code;
		}
	}
}
