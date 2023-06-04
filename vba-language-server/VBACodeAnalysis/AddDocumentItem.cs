using System;
using System.Collections.Generic;
using System.Text;

namespace VBACodeAnalysis {
    public class AddDocumentItem {
        public List<string> FilePaths { get; set; }
        public List<string> Texts { get; set; }

        public AddDocumentItem(List<string> FilePaths, List<string> Texts) {
            this.FilePaths = FilePaths;
            this.Texts = Texts;
        }
    }
}
