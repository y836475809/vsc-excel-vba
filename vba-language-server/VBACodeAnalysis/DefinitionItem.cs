using System;
using System.Collections.Generic;
using System.Text;

namespace VBACodeAnalysis {
    public class Location {
        public int Positon { get; set; }
        public int Line { get; set; }
        public int Character { get; set; }

        public Location(int Positon, int Line, int Character) {
            this.Positon = Positon;
            this.Line = Line;
            this.Character = Character;
        }
    }

    public class DefinitionItem {
        public string FilePath { get; set; }
        public Location Start { get; set; }
        public Location End { get; set; }

        private bool IsClass;
        public bool IsKindClass() {
            return this.IsClass;
        }

        public DefinitionItem(string FilePath, Location Start, Location End, bool IsClass) {
            this.FilePath = FilePath;
            this.Start = Start;
            this.End = End;
            this.IsClass = IsClass;
        }
    }
}
