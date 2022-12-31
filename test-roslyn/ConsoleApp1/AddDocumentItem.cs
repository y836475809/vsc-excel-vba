using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1 {
    public class AddDocumentItem {
        public string FilePath { get; set; }
        public string Text { get; set; }

        public AddDocumentItem(string FilePath, string Text) {
            this.FilePath = FilePath;
            this.Text = Text;
        }
    }
}
