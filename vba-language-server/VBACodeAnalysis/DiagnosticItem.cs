using System;
using System.Collections.Generic;
using System.Text;

namespace VBACodeAnalysis {
    public class DiagnosticItem {
        public string Id;
		public string Severity { get; set; }
        public string Message { get; set; }
        public int StartLine { get; set; }
        public int StartChara { get; set; }
        public int EndLine { get; set; }
        public int EndChara { get; set; }

        public DiagnosticItem(string id, string Severity, string Message, 
            int StartLine, int StartChara,
            int EndLine, int EndChara) {
            this.Id = id;
            this.Severity = Severity;
            this.Message = Message;
            this.StartLine = StartLine;
            this.StartChara = StartChara;
            this.EndLine = EndLine;
            this.EndChara = EndChara;
        }

        public bool Eq(
            int StartLine, int StartChara,
            int EndLine, int EndChara) {
            return this.StartLine == StartLine
                && this.StartChara == StartChara
                && this.EndLine == EndLine
                && this.EndChara == EndChara;
        }
		public bool Eq(DiagnosticItem obj) {
			return this.Id == obj.Id
				&& this.Severity == obj.Severity
				&& this.StartLine == obj.StartLine
				&& this.StartChara == obj.StartChara
				&& this.EndLine == obj.EndLine
				&& this.EndChara == obj.EndChara;
		}
	}
}
