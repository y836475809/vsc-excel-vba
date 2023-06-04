﻿using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1 {
    public class DiagnosticItem {
        public string Severity { get; set; }
        public string Message { get; set; }
        public int StartLine { get; set; }
        public int StartChara { get; set; }
        public int EndLine { get; set; }
        public int EndChara { get; set; }

        public DiagnosticItem(string Severity, string Message, 
            int StartLine, int StartChara,
            int EndLine, int EndChara) {
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
    }
}
