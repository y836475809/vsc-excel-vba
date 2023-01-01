using ConsoleApp1;
using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleAppServer {
    class Command {
        public string Id { get; set; }

        public string FilePath { get; set; }

        public int Position { get; set; }

        public string Text { get; set; }
    }

    class ResponseCompletion {
        public List<CompletionItem> items { get; set; }
    }
    class ResponseDefinition {
        public List<DefinitionItem> items { get; set; }
    }

    //class ResponseCompletion {
    //    public string FilePath { get; set; }

    //    public int Position { get; set; }
    //}


    public class DocumentAddedEventArgs : EventArgs {
        public string FilePath { get; set; }
        public string Text { get; set; }

        public DocumentAddedEventArgs(string FilePath) {
            this.FilePath = FilePath;
        }
    }

    public class DocumentChangedEventArgs : EventArgs {
        public string FilePath { get; set; }
        public string Text { get; set; }

        public DocumentChangedEventArgs(string FilePath, string Text) {
            this.FilePath = FilePath;
            this.Text = Text;
        }
    }

    public class CompletionEventArgs : EventArgs {
        public string FilePath { get; set; }
        public string Text { get; set; }
        public int Position { get; set; }
        public List<CompletionItem> Items { get; set; }

        public CompletionEventArgs(string FilePath, string Text, int Position) {
            this.FilePath = FilePath;
            this.Text = Text;
            this.Position = Position;
        }
    }

    //public class DefinitionItem {
    //    public string FilePath { get; set; }
    //    public int Start { get; set; }
    //    public int End { get; set; }
    //}

    public class DefinitionEventArgs : EventArgs {
        public string FilePath { get; set; }
        public string Text { get; set; }
        public int Position { get; set; }

        public List<DefinitionItem> Items { get; set; }

        public DefinitionEventArgs(string FilePath, string Text, int Position) {
            this.FilePath = FilePath;
            this.Text = Text;
            this.Position = Position;
        }
    }
}
