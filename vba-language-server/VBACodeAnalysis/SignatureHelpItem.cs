using System.Collections.Generic;

namespace VBACodeAnalysis {
    public class ArgumentItem {
        public string Name { get; set; }
        public string AsType { get; set; }

        public ArgumentItem(string Name, string AsType) {
            this.Name = Name;
            this.AsType = AsType;
        }
    }

    public class SignatureHelpItem {
        public string DisplayText { get; set; }
        public string Description { get; set; }
        public string ReturnType { get; set; }
        public string Kind { get; set; }
        public int ActiveParameter { get; set; }

        public List<ArgumentItem> Args { get; set; }

        public SignatureHelpItem() {
            this.Args = new List<ArgumentItem>();

        }
    }
}
