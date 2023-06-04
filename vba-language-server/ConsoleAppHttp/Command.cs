using ConsoleApp1;
using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleAppServer {
    class Command {
        public string id { get; set; }

        public List<string> filepaths { get; set; }

        public int line { get; set; }
        public int chara { get; set; }

        public string text { get; set; }
    }

    //public class AddDocuments {
    //    public List<string> FilePaths { get; set; }
    //}

    //public class ChangeDocument {
    //    public string FilePath { get; set; }
    //    public string Text { get; set; }
    //}

    //public class Completion {
    //    public string FilePath { get; set; }
    //    public int Position { get; set; }
    //    public string Text { get; set; }
    //}

    //class ResponseCompletion {
    //    public string FilePath { get; set; }

    //    public int Position { get; set; }
    //}

    //public class Document {
    //    public string FilePath { get; set; }
    //    public string Text { get; set; }

    //    public Document(string FilePath, string Text) {
    //        this.FilePath = FilePath;
    //        this.Text = Text;
    //    }

    //}


    public class DocumentAddedEventArgs : EventArgs {
        public List<string> FilePaths { get; set; }

        public DocumentAddedEventArgs(List<string> FilePaths) {
            this.FilePaths = FilePaths;
        }
    }

    public class DocumentDeletedEventArgs : EventArgs {
        public List<string> FilePaths { get; set; }

        public DocumentDeletedEventArgs(List<string> FilePaths) {
            this.FilePaths = FilePaths;
        }
    }

    public class DocumentRenamedEventArgs : EventArgs {
        public string OldFilePath { get; set; }
        public string NewFilePath { get; set; }

        public DocumentRenamedEventArgs(string OldFilePath, string NewFilePath) {
            this.OldFilePath = OldFilePath;
            this.NewFilePath = NewFilePath;
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
        public int Line { get; set; }
        public int Chara { get; set; }
        public List<CompletionItem> Items { get; set; }

        public CompletionEventArgs(string FilePath, string Text, int Line, int Chara) {
            this.FilePath = FilePath;
            this.Text = Text;
            this.Line = Line;
            this.Chara = Chara;
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
        public int Line { get; set; }
        public int Chara { get; set; }

        public List<DefinitionItem> Items { get; set; }

        public DefinitionEventArgs(string FilePath, string Text, int Line, int Chara) {
            this.FilePath = FilePath;
            this.Text = Text;
            this.Line = Line;
            this.Chara = Chara;
        }
    }

    public class DiagnosticEventArgs : EventArgs {
        public string FilePath { get; set; }

        public List<DiagnosticItem> Items { get; set; }

        public DiagnosticEventArgs(string FilePath) {
            this.FilePath = FilePath;
        }
    }

    public class DebugEventArgs : EventArgs {
        public string Text { get; set; }
        public DebugEventArgs() {
        }
    }

    public class ReferencesEventArgs : EventArgs {
        public string FilePath { get; set; }
        public int Line { get; set; }
        public int Chara { get; set; }
        public List<ReferenceItem> Items { get; set; }

        public ReferencesEventArgs(string FilePath, int Line, int Chara) {
            this.FilePath = FilePath;
            this.Line = Line;
            this.Chara = Chara;
        }
    }
}
