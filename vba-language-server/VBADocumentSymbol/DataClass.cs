using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBADocumentSymbol {
	//public class IDocumentSymbol {
	//	string Name { get; set; }
	//	int StartLine { get; set; }
	//	int StartColumn { get; set; }
	//	int EndLine { get; set; }
	//	int EndColumn { get; set; }
	//}

	public class IDocumentSymbol {
		public string Name { get; set; }
		public string Kind { get; set; }
		public int StartLine { get; set; }
		public int StartColumn { get; set; }
		public int EndLine { get; set; }
		public int EndColumn { get; set; }
		public List<IDocumentSymbol> Variables { get; set; }
	}
}
