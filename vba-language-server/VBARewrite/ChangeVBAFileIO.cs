using System.Collections.Generic;
using static VBAAntlr.VBAParser;

namespace VBARewrite {
	internal class ChangeVBAFileIO {
		public List<ChangeData> ChangeDataList { get; set; }
		public List<VBADiagnostic> IgnoreDiagnosticList { get; set; }

		public ChangeVBAFileIO() {
			ChangeDataList = [];
			IgnoreDiagnosticList = [];
		}

		public void ChangeOpenStmt(OpenStmtContext context) {
			var open = context.OPEN();
			var line = open.Symbol.Line - 1;
			var startCol = open.Symbol.Column;
			var endtCol = startCol + open.Symbol.Text.Length;
			AddIgnoreDiagnostic(
				(line, startCol), (line, endtCol),
				open.GetText());

			var fileNum = context.fileNumber();
			var fnText = fileNum.identifier().GetText();
			if (fnText.StartsWith('#')) {
				var st = fileNum.Start;
				ChangeDataList.Add(new(st.Line - 1,
					(st.Column, st.Column + 1), ",", st.Column, false));
			} else {
				var st = fileNum.Start;
				ChangeDataList.Add(new(st.Line - 1,
					(st.Column - 1, st.Column), ",", st.Column, false));
			}
		}

		public void ChangeOutputStmt(OutputStmtContext context) {
			var method = context.OUTPUT();
			var line = method.Symbol.Line - 1;
			var startCol = method.Symbol.Column;
			var endtCol = startCol + method.Symbol.Text.Length;
			AddIgnoreDiagnostic(
				(line, startCol), (line, endtCol),
				method.GetText());

			var fileNum = context.fileNumber();
			var fnText = fileNum.identifier().GetText();
			if (fnText.StartsWith('#')) {
				var st = fileNum.Start;
				ChangeDataList.Add(new(st.Line - 1,
					(st.Column, st.Column + 1), " ", st.Column, false));
			}
		}

		public void ChangeInputStmt(InputStmtContext context) {
			var method = context.INPUT();
			var line = method.Symbol.Line - 1;
			var startCol = method.Symbol.Column;
			var endtCol = startCol + method.Symbol.Text.Length;
			AddIgnoreDiagnostic(
				(line, startCol), (line, endtCol),
				method.GetText());

			var fileNum = context.fileNumber();
			var fnText = fileNum.identifier().GetText();
			if (fnText.StartsWith('#')) {
				var st = fileNum.Start;
				ChangeDataList.Add(new(st.Line - 1,
					(st.Column, st.Column + 1), " ", st.Column, false));
			}
		}

		public void ChangeLineInputStmt(LineInputStmtContext context) {
			var method = context.LINE_INPUT();
			var line = method.Symbol.Line - 1;
			var startCol = method.Symbol.Column;
			var endtCol = startCol + method.Symbol.Text.Length;
			AddIgnoreDiagnostic(
				(line, startCol), (line, endtCol),
				"Line_Input");

			ChangeDataList.Add(new(line,
				(startCol, endtCol), "Line_Input", startCol, false));
			var fileNum = context.fileNumber();
			var fnText = fileNum.identifier().GetText();
			if (fnText.StartsWith('#')) {
				var st = fileNum.Start;
				ChangeDataList.Add(new(st.Line - 1,
					(st.Column, st.Column + 1), " ", st.Column, false));
			}
		}

		public void ChangeCloseStmt(CloseStmtContext context) {
			var method = context.CLOSE();
			var line = method.Symbol.Line - 1;
			var startCol = method.Symbol.Column;
			var endtCol = startCol + method.Symbol.Text.Length;
			AddIgnoreDiagnostic(
				(line, startCol), (line, endtCol),
				method.GetText());

			var fileNums = context.fileNumber();
			foreach (var item in fileNums) {
				var fnText = item.identifier().GetText();
				if (fnText.StartsWith('#')) {
					var st = item.Start;
					ChangeDataList.Add(new(st.Line - 1,
						(st.Column, st.Column + 1), " ", st.Column, false));
				}
			}
		}

		public void AddIgnoreDiagnostic((int, int) start, (int, int) end, string text) {
			IgnoreDiagnosticList.Add(new() {
				Code = text,
				Start = start,
				End = end
			});
		}
	}
}
