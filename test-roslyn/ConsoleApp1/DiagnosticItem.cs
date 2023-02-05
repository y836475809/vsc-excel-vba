using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1 {
    public class DiagnosticItem {
        public string Severity { get; set; }
        public string Message { get; set; }
        public int Start { get; set; }
        public int End { get; set; }

        public DiagnosticItem(string Severity, string Message, int Start, int End) {
            this.Severity = Severity;
            this.Message = Message;
            this.Start = Start;
            this.End = End;
        }
    }
}
