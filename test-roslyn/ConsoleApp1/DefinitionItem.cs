using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1 {
    public class DefinitionItem {
        public string FilePath { get; set; }
        public int Start { get; set; }
        public int End { get; set; }

        public DefinitionItem(string FilePath, int Start, int End) {
            this.FilePath = FilePath;
            this.Start = Start;
            this.End = End;
        }
    }
}
