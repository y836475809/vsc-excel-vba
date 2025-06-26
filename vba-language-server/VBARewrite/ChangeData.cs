using System;

namespace VBARewrite {
	public class ChangeData {
		public int Line { get; set; }
		private int _lineIndex;
		private (int, int) _repColRange;
		private string _text;
		public int StartCol;
		public int? ShiftCol;
		private bool _enableShift;

		public ChangeData(int lineIndex, (int, int) repColRange, string text, int startCol, bool enableShift = true) {
			_lineIndex = lineIndex;
			Line = lineIndex;
			_repColRange = repColRange;
			_text = text;
			StartCol = startCol;
			ShiftCol = null;
			_enableShift = enableShift;
		}

		public ChangeData(int lineIndex, (int, int) repColRange, string text, int startCol, int shiftCol) {
			_lineIndex = lineIndex;
			Line = lineIndex;
			_repColRange = repColRange;
			_text = text;
			StartCol = startCol;
			ShiftCol = shiftCol;
			_enableShift = true;
		}

		public ChangeData(int lineIndex, string text) {
			_lineIndex = lineIndex;
			Line = lineIndex;
			_repColRange = (-1, -1);
			_text = text;
			StartCol = 0;
			ShiftCol = null;
			_enableShift = true;
		}

		public (ColumnShift, string) Apply(string line) {
			var (rStart, rEnd) = _repColRange;
			var text = _text;
			if (rStart < 0 && rEnd < 0) {
				var colShift = new ColumnShift(_lineIndex, StartCol, 0);
				return (colShift, text);
			}
			var t1 = line[0..rStart];
			var t2 = line[rEnd..];
			var repText = $"{t1}{text}";
			var orgText = line[..rEnd];
			var shiftCol = repText.Length - orgText.Length;
			if (ShiftCol != null) {
				shiftCol = (int)ShiftCol;
			}

			var rep_line = $"{repText}{t2}";
			if (_enableShift) {
				var colShift = new ColumnShift(_lineIndex, StartCol, shiftCol);
				return (colShift, rep_line);
			} else {
				var colShift = new ColumnShift(_lineIndex, StartCol, 0);
				return (colShift, rep_line);
			}
		}

		public bool Eq(int lineIndex, (int, int) repColRange, string text, int startCol, bool enableShift) {
			return _lineIndex == lineIndex
				&& _repColRange.Equals(repColRange)
				&& _text == text
				&& StartCol == startCol
				&& _enableShift == enableShift;
		}

		public bool Eq(ChangeData changeData) {
			return _lineIndex == changeData._lineIndex
				&& _repColRange.Equals(changeData._repColRange)
				&& _text == changeData._text
				&& StartCol == changeData.StartCol
				&& _enableShift == changeData._enableShift;
		}
	}
}
