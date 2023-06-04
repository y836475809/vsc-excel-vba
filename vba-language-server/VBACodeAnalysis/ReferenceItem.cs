using System;
using System.Collections.Generic;
using System.Text;

namespace VBACodeAnalysis {
	public class ReferenceItem {
		public string FilePath { get; set; }
		public Location Start { get; set; }
		public Location End { get; set; }

		public ReferenceItem(string FilePath, Location Start, Location End) {
			this.FilePath = FilePath;
			this.Start = Start;
			this.End = End;
		}
	}
}
