global using ChangeDataDict = System.Collections.Generic.Dictionary<int, System.Collections.Generic.List<VBARewrite.ChangeData>>;
global using ColumnShiftDict = System.Collections.Generic.Dictionary<int, System.Collections.Generic.List<VBARewrite.ColumnShift>>;
global using LineMapDict = System.Collections.Generic.Dictionary<int, int>;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBARewrite {
	public class ColumnShift(int lineIndex, int startCol, int shiftCol) {
		public int LineIndex = lineIndex;
		public int StartCol = startCol;
		public int ShiftCol = shiftCol;
	}

	public class PropertyName(int lineIndex, string prefix, string text, string asType) {
		public int LineIndex = lineIndex;
		public string Prefix = prefix;
		public string Text = text;
		public string AsType = asType;
	}

	public class PropertyMember(string name, int line) {
		public string Name = name;
		public int Line = line;
	}

	public class PropertyDiagnostic(string id, string code, string severity, int line) {
		public string Id = id;
		public string Code = code;
		public string Severity = severity;
		public int Line = line;
	}

	public class AttributeVBName(int line, int startChara, int endChara, string vbaName) {
		public int Line = line;
		public int StartChara = startChara;
		public int EndChara = endChara;
		public string VBAName = vbaName;
	}

	public class VBADiagnostic {
		public string ID { get; set; }
		public string Code { get; set; }
		public string Severity { get; set; }
		public string Message { get; set; }
		public (int, int) Start { get; set; }
		public (int, int) End { get; set; }

		public bool Eq(VBADiagnostic obj) {
			return this.ID == obj.ID
				&& this.Severity == obj.Severity
				&& this.Start.Item1 == obj.Start.Item1
				&& this.Start.Item2 == obj.Start.Item2
				&& this.End.Item1 == obj.End.Item1
				&& this.End.Item2 == obj.End.Item2;
		}
	}

	

	public class VBCode {
		public string Code { get; set; }
		public ColumnShiftDict ColShiftDict { get; set; }
		public LineMapDict LineMapDict { get; set; }
		public List<VBADiagnostic> DiagnosticList { get; set; }
		public List<VBADiagnostic> IgnoreDiagnosticList { get; set; }
		public List<PropertyDiagnostic> IgnorePropertyDiagnosticList { get; set; }
		public HashSet<int> IgnoreDiagnosticLineSet { get; set; }
	}
}
