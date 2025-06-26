using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBACodeAnalysis {
	public class VBACompletionItem {
		public string Label { get; set; }
		public string Display { get; set; }
		public string Doc { get; set; }
		public string Kind { get; set; }
	}

	public class VBAParameterInfo {
		public string Label { get; set; }
		public string Doc { get; set; }
	}

	public class VBASignatureInfo {
		public string Label { get; set; }
		public string Doc { get; set; }
		public List<VBAParameterInfo> ParameterInfos { get; set; }
	}

	public class VBALocation {
		public Uri Uri { get; set; }
		public (int, int) Start { get; set; }
		public (int, int) End { get; set; }
	}

	public class VBContent {
		public string Language { get; set; }
		public string Value { get; set; }
	}

	public class VBAHover {
		public Uri Uri { get; set; }
		public (int, int) Start { get; set; }
		public List<VBContent> Contents { get; set; }
	}
}
