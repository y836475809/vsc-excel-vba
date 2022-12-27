using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1 {
    public class CompletionItem {
        public string DisplayText { get; set; }
        public string CompletionText { get; set; }
        public string Description { get; set; }
        public string ReturnType { get; set; }
        public string Kind { get; set; }

        public override bool Equals(object other) {
            var otherItem = other as CompletionItem;
            return otherItem != null
                && otherItem.DisplayText == DisplayText
                && otherItem.CompletionText == CompletionText
                && otherItem.Description == Description
                && otherItem.ReturnType == ReturnType
                && otherItem.Kind == Kind;
        }

        public override int GetHashCode() {
            return new {
                DisplayText,
                CompletionText,
                Description,
                ReturnType,
                Kind
            }.GetHashCode();
        }
    }
}
