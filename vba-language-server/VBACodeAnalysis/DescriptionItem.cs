using System;
using System.Collections.Generic;
using System.Text;

namespace VBACodeAnalysis {
    public class DescriptionParam {
        public string Name { get; set; }
        public string Text { get; set; }

        public DescriptionParam(string Name, string Text) {
            this.Name = Name;
            this.Text = Text;
        }
    }

    public class DescriptionItem {
        public string Summary { get; set; }
        public List<DescriptionParam> Params { get; set; }
        public string Returns { get; set; }

        public DescriptionItem(string Summary, List<DescriptionParam> Params, string Returns) {
            this.Summary = Summary;
            this.Params = Params;
            this.Returns = Returns;
        }
    }
}
